# main.py
import os
import re
import json
import requests
import docx
import PyPDF2
import pyodbc
import tempfile
import magic
import fitz  # PyMuPDF
import pytesseract
import io
import time
from PIL import Image
from flask import Flask, request, jsonify
from flask_cors import CORS
from werkzeug.datastructures import FileStorage
from collections import Counter
from datetime import datetime, timedelta
from dotenv import load_dotenv
import logging

# Import services
from compare_service import compare_service, initialize_compare_service
from document_generation_service import document_generation_service, initialize_document_generation_service
from case_analysis_service import case_analysis_service, initialize_case_analysis_service
from auth_middleware import require_auth, check_document_quota, get_user_from_token, require_role, log_user_action

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

load_dotenv()

# --- Initialization ---
app = Flask(__name__)
app.register_blueprint(compare_service, url_prefix='/service-plus')
app.register_blueprint(document_generation_service, url_prefix='/service-plus')
app.register_blueprint(case_analysis_service, url_prefix='/service-plus')

CORS(app, resources={
    r"/*": {
        "origins": ["http://localhost:3001", "http://localhost:4028", "http://localhost:5173", "http://localhost:8080"],
        "methods": ["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
        "allow_headers": ["Content-Type", "Authorization", "X-Requested-With"],
        "supports_credentials": True
    }
})

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

SQLSERVER_CONN_STRING = os.getenv("SQLSERVER_CONN_STRING")

def get_db_connection():
    if not SQLSERVER_CONN_STRING:
        raise ValueError("SQLSERVER_CONN_STRING not found")
    for attempt in range(3):
        try:
            return pyodbc.connect(SQLSERVER_CONN_STRING)
        except pyodbc.OperationalError as e:
            if attempt == 2:
                raise e
            time.sleep(1)

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        IF NOT EXISTS(SELECT * FROM sysobjects WHERE name='cases' AND xtype='U')
        CREATE TABLE cases(
            id INT IDENTITY(1,1) PRIMARY KEY,
            filename NVARCHAR(255) NOT NULL,
            title NVARCHAR(255),
            summary NVARCHAR(MAX),
            parties NVARCHAR(MAX),
            document_date NVARCHAR(50),
            court NVARCHAR(255),
            arguments NVARCHAR(MAX),
            document_type NVARCHAR(100),
            creation_date NVARCHAR(50),
            document_language NVARCHAR(50),
            file_size BIGINT,
            extracted_text NVARCHAR(MAX),
            analysis_duration_ms BIGINT,
            classification NVARCHAR(50),
            enable_ocr BIT DEFAULT 1,
            file_path NVARCHAR(500),
            status NVARCHAR(50) DEFAULT 'Pending',
            practice_area NVARCHAR(100),
            priority NVARCHAR(20) DEFAULT 'normal',
            confidence_score FLOAT,
            needs_review BIT DEFAULT 0,
            tags NVARCHAR(MAX),
            updated_at DATETIME DEFAULT GETDATE(),
            analyzed_at DATETIME,
            original_name NVARCHAR(255),
            user_id NVARCHAR(450),
            client_id INT
        );
    """)

    new_columns = [
        ("practice_area", "NVARCHAR(100)"),
        ("priority", "NVARCHAR(20) DEFAULT 'normal'"),
        ("confidence_score", "FLOAT"),
        ("needs_review", "BIT DEFAULT 0"),
        ("tags", "NVARCHAR(MAX)"),
        ("updated_at", "DATETIME DEFAULT GETDATE()"),
        ("analyzed_at", "DATETIME"),
        ("original_name", "NVARCHAR(255)"),
        ("user_id", "NVARCHAR(450)"),  # ASP.NET Identity UserId
        ("client_id", "INT")
    ]

    for col, defn in new_columns:
        try:
            cursor.execute(f"IF COL_LENGTH('cases', '{col}') IS NULL ALTER TABLE cases ADD {col} {defn}")
            logger.info(f"Added column: {col}")
        except Exception as e:
            logger.debug(f"Column {col} already exists: {e}")

    # Add indexes for performance
    try:
        cursor.execute("""
            IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_cases_user_id')
            CREATE INDEX IX_cases_user_id ON cases(user_id)
        """)
        logger.info("Created index on user_id")
    except Exception as e:
        logger.debug(f"Index IX_cases_user_id already exists: {e}")

    try:
        cursor.execute("""
            IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_cases_client_id')
            CREATE INDEX IX_cases_client_id ON cases(client_id)
        """)
        logger.info("Created index on client_id")
    except Exception as e:
        logger.debug(f"Index IX_cases_client_id already exists: {e}")

    # Create Clients table if not exists
    cursor.execute("""
        IF NOT EXISTS(SELECT * FROM sysobjects WHERE name='Clients' AND xtype='U')
        CREATE TABLE Clients(
            Id INT IDENTITY(1,1) PRIMARY KEY,
            Name NVARCHAR(255) NOT NULL,
            Email NVARCHAR(255),
            Phone NVARCHAR(50),
            Address NVARCHAR(500),
            UserId NVARCHAR(450) NOT NULL,
            CreatedAt DATETIME DEFAULT GETDATE(),
            UpdatedAt DATETIME DEFAULT GETDATE()
        );
    """)

    # Add index on Clients.UserId
    try:
        cursor.execute("""
            IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Clients_UserId')
            CREATE INDEX IX_Clients_UserId ON Clients(UserId)
        """)
        logger.info("Created index on Clients.UserId")
    except Exception as e:
        logger.debug(f"Index IX_Clients_UserId already exists: {e}")

    conn.commit()
    conn.close()
    logger.info("SQL Server database initialized.")

# --- Helper: Verify client ownership ---
def verify_client_ownership(client_id, user_id):
    """Verify that a client belongs to the specified user"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT Id FROM Clients WHERE Id = ? AND UserId = ?", (client_id, user_id))
        result = cursor.fetchone()
        conn.close()
        return result is not None
    except Exception as e:
        logger.error(f"Error verifying client ownership: {e}")
        return False

# --- Practice Area, Confidence, Priority logic ---
def detect_practice_area(text, document_type=None, parties=None):
    text_lower = text.lower()
    practice_areas = {
        'corporate': ['merger', 'acquisition', 'shareholder', 'board of directors', 'corporate governance',
                      'securities', 'stock', 'shares', 'corporation', 'company formation', 'bylaws',
                      'articles of incorporation', 'dividend', 'proxy', 'corporate resolution'],
        'litigation': ['plaintiff', 'defendant', 'lawsuit', 'court', 'trial', 'hearing', 'motion',
                       'discovery', 'deposition', 'subpoena', 'settlement', 'judgment', 'appeal',
                       'civil procedure', 'evidence', 'testimony', 'cross-examination'],
        'contract': ['agreement', 'contract', 'terms and conditions', 'breach', 'performance',
                     'consideration', 'party obligations', 'warranty', 'indemnification',
                     'force majeure', 'termination clause', 'non-disclosure', 'confidentiality'],
        'employment': ['employee', 'employer', 'workplace', 'discrimination', 'harassment',
                       'wrongful termination', 'wage', 'salary', 'benefits', 'labor law',
                       'union', 'collective bargaining', 'overtime', 'workers compensation'],
        'intellectual_property': ['patent', 'trademark', 'copyright', 'trade secret', 'infringement',
                                  'intellectual property', 'licensing', 'royalty', 'brand', 'invention',
                                  'proprietary', 'patent application', 'copyright registration'],
        'real_estate': ['property', 'real estate', 'lease', 'landlord', 'tenant', 'mortgage',
                        'deed', 'title', 'zoning', 'easement', 'property tax', 'closing',
                        'purchase agreement', 'rental agreement', 'eviction'],
        'family': ['divorce', 'custody', 'alimony', 'child support', 'adoption',
                   'marriage', 'prenuptial', 'domestic', 'family court', 'visitation',
                   'separation agreement', 'parental rights'],
        'criminal': ['criminal', 'felony', 'misdemeanor', 'prosecution', 'defense',
                     'guilty', 'innocent', 'sentencing', 'probation', 'parole',
                     'criminal code', 'arrest', 'indictment', 'plea bargain'],
        'tax': ['tax', 'taxation', 'irs', 'tax return', 'audit', 'deduction',
                'tax liability', 'tax planning', 'estate tax', 'income tax',
                'tax evasion', 'tax compliance'],
        'immigration': ['visa', 'immigration', 'citizenship', 'green card', 'deportation',
                        'asylum', 'refugee', 'naturalization', 'immigration law',
                        'border', 'immigrant', 'permanent residence'],
        'regulatory': ['regulation', 'compliance', 'regulatory', 'agency', 'licensing',
                       'permit', 'environmental', 'safety', 'health regulations',
                       'government oversight', 'administrative law'],
        'bankruptcy': ['bankruptcy', 'debtor', 'creditor', 'insolvency', 'liquidation',
                       'reorganization', 'chapter 7', 'chapter 11', 'discharge',
                       'trustee', 'bankruptcy court']
    }

    scores = {}
    for area, keywords in practice_areas.items():
        score = sum(text_lower.count(kw) * len(kw.split()) for kw in keywords)
        scores[area] = score

    if document_type:
        doc_type_lower = document_type.lower()
        if 'contract' in doc_type_lower or 'agreement' in doc_type_lower:
            scores['contract'] += 10
        elif 'court' in doc_type_lower or 'judgment' in doc_type_lower:
            scores['litigation'] += 10
        elif 'patent' in doc_type_lower or 'trademark' in doc_type_lower:
            scores['intellectual_property'] += 10

    if scores:
        max_area = max(scores.items(), key=lambda x: x[1])
        if max_area[1] > 0:
            return max_area[0].replace('_', ' ').title()
    return 'General Legal'

def calculate_confidence_score(analysis_result, text_length):
    if not analysis_result:
        return 0.0
    score = 0.0
    required_fields = ['court', 'document_date', 'parties', 'summary', 'arguments', 'document_type']
    for field in required_fields:
        if analysis_result.get(field) and len(str(analysis_result[field]).strip()) > 0:
            score += 15.0
    if text_length > 5000:
        score += 5.0
    elif text_length > 1000:
        score += 3.0
    elif text_length < 100:
        score -= 10.0
    if analysis_result.get('summary'):
        summary_len = len(analysis_result['summary'])
        if summary_len > 200:
            score += 5.0
        elif summary_len > 50:
            score += 2.0
    return min(score / 100.0, 1.0)

def determine_priority(text, document_type=None, parties=None):
    text_lower = text.lower()
    high_priority_keywords = ['urgent', 'emergency', 'immediate', 'asap', 'deadline', 'time sensitive',
                              'court order', 'injunction', 'restraining order', 'motion for', 'appeal',
                              'bankruptcy', 'foreclosure', 'eviction', 'criminal', 'arrest warrant']
    medium_priority_keywords = ['contract', 'agreement', 'settlement', 'negotiation', 'dispute',
                                'litigation', 'lawsuit', 'legal action', 'compliance', 'audit']
    high_score = sum(text_lower.count(kw) for kw in high_priority_keywords)
    medium_score = sum(text_lower.count(kw) for kw in medium_priority_keywords)
    if high_score > 0:
        return 'high'
    elif medium_score > 2:
        return 'medium'
    else:
        return 'normal'

# --- Text Extraction ---
def extract_text_with_ocr(filepath):
    logger.info(f"Processing file: {filepath}")
    if filepath.endswith(".pdf"):
        return extract_pdf_text_with_ocr(filepath)
    elif filepath.endswith(".docx"):
        return extract_docx_text(filepath)
    elif filepath.endswith(".txt"):
        return extract_txt_text(filepath)
    else:
        logger.error(f"Unsupported file type: {filepath}")
        return None

def extract_pdf_text_with_ocr(filepath):
    try:
        with open(filepath, 'rb') as f:
            content = f.read()
        mime_type = magic.from_buffer(content, mime=True)
        if mime_type != "application/pdf":
            logger.error(f"Invalid file type: {filepath}, MIME type: {mime_type}")
            return None
        pdf_doc = fitz.open(filepath)
        extracted_text = []
        logger.info(f"PDF has {pdf_doc.page_count} pages")
        for page_num in range(pdf_doc.page_count):
            page = pdf_doc[page_num]
            text = page.get_text().strip()
            if text:
                text = re.sub(r"(\n\s*)+\n", "\n\n", text)
                text = re.sub(r"(?<=\w)\s+(?=[A-Z][a-z])", "\n\n", text)
                extracted_text.append(text)
            else:
                pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                languages = "eng+fra+deu+spa+ita+por+nld"
                ocr_text = pytesseract.image_to_string(img, lang=languages)
                if ocr_text.strip():
                    ocr_text = re.sub(r"(\n\s*)+\n", "\n\n", ocr_text)
                    ocr_text = re.sub(r"(?<=\w)\s+(?=[A-Z][a-z])", "\n\n", ocr_text)
                    extracted_text.append(ocr_text)
                else:
                    extracted_text.append(f"[OCR Error: No text extracted from page {page_num + 1}]")
        pdf_doc.close()
        result = "\n\n".join([t.strip() for t in extracted_text if t.strip()])
        return result if result else None
    except Exception as e:
        logger.error(f"OCR failed for {filepath}: {str(e)}")
        return None

def extract_docx_text(filepath):
    try:
        doc = docx.Document(filepath)
        return "\n".join(para.text for para in doc.paragraphs)
    except Exception as e:
        logger.error(f"Error reading DOCX {filepath}: {e}")
        return None

def extract_txt_text(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8') as file:
            return file.read()
    except Exception as e:
        logger.error(f"Error reading TXT {filepath}: {e}")
        return None

def clean_extracted_text(text):
    text = re.sub(r'Digitally signed by.*', '', text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'NC:\d{4}:KHC:\d+', '', text)
    text = re.sub(r'RSA No\.\d+ of \d+', '', text)
    text = re.sub(r'--- PAGE \d+ ---', '', text)
    text = re.sub(r'-\d+-', '', text)
    text = "\n".join(line.strip() for line in text.splitlines() if line.strip())
    return text

def perform_gemini_analysis(text):
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        logger.error("GEMINI_API_KEY not found.")
        return None
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    json_schema = {
        "type": "OBJECT",
        "properties": {
            "court": {"type": "STRING"},
            "document_date": {"type": "STRING"},
            "parties": {"type": "STRING"},
            "summary": {"type": "STRING"},
            "arguments": {"type": "ARRAY", "items": {"type": "STRING"}},
            "document_type": {"type": "STRING"},
            "document_language": {"type": "STRING"},
            "key_entities": {"type": "ARRAY", "items": {"type": "STRING"}},
            "legal_citations": {"type": "ARRAY", "items": {"type": "STRING"}}
        },
        "required": ["court", "document_date", "parties", "summary", "arguments", "document_type", "document_language"]
    }
    prompt = f"""Analyze the following legal document text.
Extract the specified information and always respond in the SAME LANGUAGE as the source document.
Do not translate or switch languages.
For key_entities, include important people, organizations, or legal entities mentioned.
For legal_citations, include any case law, statutes, or legal references cited.
Document Text:
---{text[:20000]}---
Return ONLY in the specified JSON format."""
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "responseMimeType": "application/json",
            "responseSchema": json_schema
        }
    }
    try:
        response = requests.post(url, json=payload)
        
        if response.status_code != 200:
            logger.error(f"Gemini API error {response.status_code}: {response.text}")
        
        response.raise_for_status()
        outer_json = response.json()
        
        if 'candidates' in outer_json and len(outer_json['candidates']) > 0:
            json_string = outer_json['candidates'][0]['content']['parts'][0]['text']
            return json.loads(json_string)
        else:
            logger.error(f"Unexpected Gemini response structure: {outer_json}")
            return None

    except requests.exceptions.RequestException as e:
        logger.error(f"HTTP request failed: {e}")
        return None
    except json.JSONDecodeError as e:
        logger.error(f"Failed to parse JSON from Gemini response: {e}")
        return None
    except Exception as e:
        logger.error(f"An unexpected error occurred during Gemini analysis: {e}")
        return None

# --- Routes ---
@app.route('/analyze', methods=['POST'])
@require_auth
def analyze_document_route():
    """
    Upload and analyze a document with proper client isolation and duplicate handling.
    Returns status='duplicate' when the same file is uploaded for the same client.
    """
    user = get_user_from_token()
    if not user:
        return jsonify({"error": "Authentication required"}), 401
        
    user_id = user.get('user_id')
    user_email = user.get('email')
    
    # 🔥 DEBUG: Log exactly what Flask receives
    logger.info(f"=== Request Debug ===")
    logger.info(f"Content-Type header: {request.headers.get('Content-Type')}")
    logger.info(f"Form keys: {list(request.form.keys())}")
    logger.info(f"Files keys: {list(request.files.keys())}")
    for key in request.form:
        logger.info(f"Form[{key}]: {request.form[key][:100] if len(request.form[key]) > 100 else request.form[key]}")
    for key in request.files:
        f = request.files[key]
        logger.info(f"File[{key}]: {f.filename}, size: {len(f.read())}, mimetype: {f.content_type}")
        f.seek(0)  # Reset file pointer
    logger.info(f"=== End Debug ===")
    
    client_id = request.form.get('clientId')
    if not client_id:
        logger.error(f"Missing clientId. Form data: {dict(request.form)}")
        return jsonify({
            "error": "Client ID is required",
            "received_fields": list(request.form.keys()),
            "received_files": list(request.files.keys())
        }), 400
    
    # Verify client belongs to user
    try:
        client_id = int(client_id)
    except ValueError:
        return jsonify({
            "error": f"Invalid client ID format: {client_id}", 
            "code": "INVALID_CLIENT_ID"
        }), 400
        
    if not verify_client_ownership(client_id, user_id):
            return jsonify({"error": "Invalid client ID or access denied"}), 403
        
    # Check document quota
    can_process, error_msg = check_document_quota()
    if not can_process:
        return jsonify({
            "error": error_msg,
            "quota_exceeded": True,
            "documents_processed": user.get('documents_processed', 0),
            "quota": user.get('document_quota', 0)
        }), 403
    
    # Validate file upload
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    
    # Extract form parameters
    title = request.form.get('title', file.filename)
    classification = request.form.get('classification', 'auto')
    language = request.form.get('language', 'en')
    enable_ocr = request.form.get('enableOCR', 'true').lower() == 'true'
    priority = request.form.get('priority', 'normal')
    practice_area = request.form.get('practiceArea', '')
    tags = request.form.get('tags', '')
    
    start_time = time.time()
    filename = file.filename
    original_name = filename
    file_size = len(file.read())
    file.seek(0)  # Reset file pointer after reading size
    
    # Save file to disk
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)
    
    # Extract text from file
    raw_text = extract_text_with_ocr(filepath)
    if raw_text is None:
        # Clean up file if extraction fails
        if os.path.exists(filepath):
            os.remove(filepath)
        return jsonify({"error": "Could not read file or extract text"}), 500
    
    cleaned_text = clean_extracted_text(raw_text)
    conn = get_db_connection()
    cursor = conn.cursor()
    creation_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # 🔥 FIX: Try to insert with composite key (user_id + client_id + filename)
    try:
        cursor.execute("""
            INSERT INTO cases(
                filename, title, file_size, extracted_text, classification, 
                enable_ocr, file_path, status, creation_date, original_name, 
                priority, tags, user_id, client_id
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, 'Processing', ?, ?, ?, ?, ?, ?)
        """, (filename, title, file_size, raw_text, classification, enable_ocr, 
              filepath, creation_date, original_name, priority, tags, user_id, client_id))
        conn.commit()
        cursor.execute("SELECT SCOPE_IDENTITY()")
        document_id = cursor.fetchone()[0]
        
    except pyodbc.IntegrityError:
        conn.rollback()  # Rollback failed transaction
        
        # 🔥 Check if duplicate exists for SAME user + SAME client + SAME filename
        cursor.execute(
            "SELECT id, status FROM cases WHERE filename = ? AND user_id = ? AND client_id = ?", 
            (filename, user_id, client_id)
        )
        result = cursor.fetchone()
        
        if result:
            # ✅ Document already exists for this client - return friendly duplicate response
            existing_id, existing_status = result
            conn.close()
            
            # 🔥 Return 200 OK with special status for frontend to handle
            return jsonify({
                "status": "duplicate",  # ← Special status for frontend
                "message": f"Document '{filename}' already exists for this client",
                "id": existing_id,
                "filename": filename,
                "existing_status": existing_status,
                "user_id": user_id,
                "client_id": client_id
            }), 200  # ← Return 200, not an error!
        else:
            # Filename exists for DIFFERENT user or client - allow it with unique name
            unique_filename = f"{filename.rsplit('.', 1)[0]}_{int(time.time())}.{filename.rsplit('.', 1)[-1]}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            file.seek(0)
            file.save(filepath)  # Re-save with unique name
            
            # Retry insert with unique filename
            cursor.execute("""
                INSERT INTO cases(filename, title, file_size, extracted_text, classification, 
                    enable_ocr, file_path, status, creation_date, original_name, 
                    priority, tags, user_id, client_id)
                VALUES (?, ?, ?, ?, ?, ?, ?, 'Processing', ?, ?, ?, ?, ?, ?)
            """, (unique_filename, title, file_size, raw_text, classification, enable_ocr, 
                  filepath, creation_date, original_name, priority, tags, user_id, client_id))
            conn.commit()
            cursor.execute("SELECT SCOPE_IDENTITY()")
            document_id = cursor.fetchone()[0]
            filename = unique_filename  # Update filename for response
    
    # Perform Gemini AI analysis
    analysis = perform_gemini_analysis(cleaned_text)
    analysis_duration = int((time.time() - start_time) * 1000)
    analyzed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Handle analysis failure
    if analysis is None:
        cursor.execute("""
            UPDATE cases 
            SET status = 'Error', analysis_duration_ms = ?, updated_at = GETDATE() 
            WHERE id = ?
        """, (analysis_duration, document_id))
        conn.commit()
        conn.close()
        # Clean up file on error
        if os.path.exists(filepath):
            os.remove(filepath)
        return jsonify({"error": "AI analysis failed"}), 500
    
    # Auto-detect practice area if not provided
    if not practice_area:
        practice_area = detect_practice_area(cleaned_text, analysis.get('document_type'), 
                                            analysis.get('parties'))
    
    # Auto-detect priority if normal
    if priority == 'normal':
        detected_priority = determine_priority(cleaned_text, analysis.get('document_type'), 
                                              analysis.get('parties'))
        if detected_priority != 'normal':
            priority = detected_priority
    
    # Calculate confidence score
    confidence_score = calculate_confidence_score(analysis, len(cleaned_text))
    needs_review = confidence_score < 0.7
    
    # Format arguments for storage
    arguments_str = "||".join(analysis.get('arguments', []))
    
    # Update document record with analysis results
    cursor.execute("""
        UPDATE cases 
        SET summary = ?, parties = ?, document_date = ?, court = ?, arguments = ?, 
            document_type = ?, document_language = ?, analysis_duration_ms = ?, 
            status = 'Analyzed', practice_area = ?, priority = ?, confidence_score = ?, 
            needs_review = ?, updated_at = GETDATE(), analyzed_at = ?
        WHERE id = ?
    """, (analysis.get('summary'), analysis.get('parties'), analysis.get('document_date'), 
          analysis.get('court'), arguments_str, analysis.get('document_type'), 
          analysis.get('document_language'), analysis_duration, practice_area, priority, 
          confidence_score, needs_review, analyzed_at, document_id))
    conn.commit()
    
    # Update user's document count in .NET database
    try:
        update_user_document_count(user_id)
    except Exception as e:
        logger.error(f"Failed to update user document count: {e}")
        # Continue even if this fails - it's not critical
    
    conn.close()
    
    # Log the action for audit trail
    log_user_action('document_analyzed', {
        'document_id': document_id,
        'filename': filename,
        'file_size': file_size,
        'client_id': client_id,
        'status': 'success'
    })
    
    # Return success response
    return jsonify({
        "status": "success",
        "filename": filename,
        "id": document_id,
        "analysis_duration_ms": analysis_duration,
        "file_size": file_size,
        "practice_area": practice_area,
        "priority": priority,
        "confidence_score": confidence_score,
        "needs_review": needs_review,
        "user_id": user_id,
        "client_id": client_id
    }), 200

def update_user_document_count(user_id):
    """Update the user's document count in the .NET database"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        cursor.execute("""
            UPDATE Users 
            SET DocumentsProcessedThisMonth = DocumentsProcessedThisMonth + 1
            WHERE Id = ?
        """, (user_id,))
        
        conn.commit()
        conn.close()
    except Exception as e:
        logger.error(f"Error updating user document count: {e}")
        raise

@app.route('/cases', methods=['GET'])
@require_auth
def get_all_cases():
    return get_documents()

@app.route('/documents', methods=['GET'])
@require_auth
def get_documents():
    """
    Get all documents with proper client isolation and filtering.
    Non-admin users only see documents for their selected client.
    """
    user = get_user_from_token()
    if not user:
        return jsonify({"error": "Authentication required"}), 401
        
    user_id = user.get('user_id')
    user_roles = user.get('roles', [])
    
    # Admin and Manager can see all documents; regular users see only their client's docs
    is_admin = 'Admin' in user_roles or 'Manager' in user_roles
    
    # Get all filter parameters from query string
    date_range = request.args.get('dateRange', 'all')
    document_type = request.args.get('documentType', 'all')
    practice_area = request.args.get('practiceArea', 'all')
    status = request.args.get('status', 'all')
    priority = request.args.get('priority', 'all')
    needs_review = request.args.get('needsReview', 'all')
    client_id = request.args.get('clientId')  # ← Critical: Client filter
    
    conn = get_db_connection()
    cursor = conn.cursor()
    where_conditions = []
    params = []
    
    # 🔐 CRITICAL: Role-based filtering - non-admin users only see their own documents
    if not is_admin:
        where_conditions.append("user_id = ?")
        params.append(user_id)
        
        # 🔥 FIX: If no clientId provided for non-admin, return empty list
        # This enforces client selection before viewing documents
        if not client_id:
            conn.close()
            return jsonify([])
    
    # 🔥 FIX: Client filter - ALWAYS apply when clientId is provided
    if client_id:
        try:
            client_id_int = int(client_id)
            # Verify client belongs to user (unless admin)
            if not is_admin and not verify_client_ownership(client_id_int, user_id):
                conn.close()
                return jsonify({"error": "Access denied to this client"}), 403
            where_conditions.append("client_id = ?")
            params.append(client_id_int)
        except ValueError:
            conn.close()
            return jsonify({"error": "Invalid client ID"}), 400
    
    # Date range filter
    if date_range != 'all':
        days = {'7days': 7, '30days': 30, '90days': 90, '6months': 180, '1year': 365}.get(date_range)
        if days:
            where_conditions.append("creation_date >= ?")
            params.append((datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d"))
    
    # Document type filter
    if document_type != 'all':
        where_conditions.append("LOWER(document_type) LIKE ?")
        params.append(f"%{document_type.lower()}%")
    
    # Practice area filter
    if practice_area != 'all':
        where_conditions.append("LOWER(practice_area) LIKE ?")
        params.append(f"%{practice_area.lower()}%")
    
    # Status filter
    if status != 'all':
        where_conditions.append("status = ?")
        params.append(status)
    
    # Priority filter
    if priority != 'all':
        where_conditions.append("priority = ?")
        params.append(priority)
    
    # Needs review filter
    if needs_review == 'true':
        where_conditions.append("needs_review = 1")
    elif needs_review == 'false':
        where_conditions.append("needs_review = 0")
    
    # Build the base query
    base_query = """
        SELECT id, filename, original_name, title, summary, parties, document_date, court, 
               arguments, document_type, creation_date, document_language, file_size, 
               analysis_duration_ms, classification, status, practice_area, priority, 
               confidence_score, needs_review, tags, updated_at, analyzed_at, user_id, client_id
        FROM cases
    """
    
    # Add WHERE clause if we have conditions
    if where_conditions:
        query = base_query + " WHERE " + " AND ".join(where_conditions) + " ORDER BY creation_date DESC"
    else:
        query = base_query + " ORDER BY creation_date DESC"
    
    cursor.execute(query, params)
    rows = cursor.fetchall()
    
    # Convert rows to list of dictionaries
    cases = []
    for row in rows:
        case_dict = dict(zip([col[0] for col in cursor.description], row))
        # Convert datetime objects to ISO format strings for JSON serialization
        for key, value in case_dict.items():
            if isinstance(value, datetime):
                case_dict[key] = value.isoformat()
        cases.append(case_dict)
    
    conn.close()
    return jsonify(cases)

@app.route('/analytics', methods=['GET'])
@require_auth
def get_analytics():
    user = get_user_from_token()
    user_id = user.get('user_id')
    user_roles = user.get('roles', [])
    is_admin = 'Admin' in user_roles or 'Manager' in user_roles
    
    client_id = request.args.get('clientId')
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Build WHERE clause for user/client filtering
    where_clause = ""
    params = []
    
    if not is_admin:
        where_clause = "WHERE user_id = ?"
        params = [user_id]
    
    if client_id:
        client_id_int = int(client_id)
        if not is_admin and not verify_client_ownership(client_id_int, user_id):
            return jsonify({"error": "Access denied"}), 403
        
        if where_clause:
            where_clause += " AND client_id = ?"
        else:
            where_clause = "WHERE client_id = ?"
        params.append(client_id_int)
    
    # Document types
    query = f"SELECT document_type, COUNT(*) FROM cases {where_clause} AND status='Analyzed' GROUP BY document_type"
    cursor.execute(query, params)
    doc_types = [{"type": row[0], "count": row[1]} for row in cursor.fetchall()]
    
    # Practice areas
    query = f"SELECT practice_area, COUNT(*) FROM cases {where_clause} AND status='Analyzed' AND practice_area IS NOT NULL GROUP BY practice_area"
    cursor.execute(query, params)
    practice_areas = [{"area": row[0], "count": row[1]} for row in cursor.fetchall()]
    
    # Languages
    query = f"SELECT document_language, COUNT(*) FROM cases {where_clause} AND status='Analyzed' GROUP BY document_language"
    cursor.execute(query, params)
    languages = [{"language": row[0], "count": row[1]} for row in cursor.fetchall()]
    
    # Priorities
    query = f"SELECT priority, COUNT(*) FROM cases {where_clause} GROUP BY priority"
    cursor.execute(query, params)
    priorities = [{"priority": row[0], "count": row[1]} for row in cursor.fetchall()]
    
    # Status counts
    query = f"SELECT status, COUNT(*) FROM cases {where_clause} GROUP BY status"
    cursor.execute(query, params)
    status_counts = {row[0]: row[1] for row in cursor.fetchall()}
    
    # Average duration
    query = f"SELECT AVG(CAST(analysis_duration_ms AS FLOAT)) FROM cases {where_clause} AND status='Analyzed' AND analysis_duration_ms IS NOT NULL"
    cursor.execute(query, params)
    avg_duration_row = cursor.fetchone()
    avg_duration = int(avg_duration_row[0]) if avg_duration_row and avg_duration_row[0] is not None else 0
    
    # Average confidence
    query = f"SELECT AVG(confidence_score) FROM cases {where_clause} AND status='Analyzed' AND confidence_score IS NOT NULL"
    cursor.execute(query, params)
    avg_confidence_row = cursor.fetchone()
    avg_confidence = float(avg_confidence_row[0]) if avg_confidence_row and avg_confidence_row[0] is not None else 0.0
    
    # Total size
    query = f"SELECT SUM(file_size) FROM cases {where_clause}"
    cursor.execute(query, params)
    total_size_row = cursor.fetchone()
    total_size = total_size_row[0] if total_size_row and total_size_row[0] is not None else 0
    
    # Needs review count
    query = f"SELECT COUNT(*) FROM cases {where_clause} AND needs_review = 1"
    cursor.execute(query, params)
    needs_review_count_row = cursor.fetchone()
    needs_review_count = needs_review_count_row[0] if needs_review_count_row else 0
    
    total_documents = sum(status_counts.values())
    
    conn.close()
    return jsonify({
        "document_types": doc_types,
        "practice_areas": practice_areas,
        "languages": languages,
        "priorities": priorities,
        "average_analysis_time_ms": avg_duration,
        "average_confidence_score": avg_confidence,
        "total_file_size_bytes": total_size,
        "document_counts": {
            "total": total_documents,
            "analyzed": status_counts.get('Analyzed', 0),
            "processing": status_counts.get('Processing', 0),
            "errors": status_counts.get('Error', 0),
            "pending": status_counts.get('Pending', 0),
            "needs_review": needs_review_count
        }
    })

@app.route('/clients', methods=['GET'])
@require_auth
def get_clients():
    """Get all clients belonging to the current user"""
    user = get_user_from_token()
    user_id = user.get('user_id')
    user_roles = user.get('roles', [])
    is_admin = 'Admin' in user_roles or 'Manager' in user_roles
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    if is_admin:
        # Admins see all clients
        cursor.execute("""
            SELECT Id, Name, Email, Phone, Address, UserId, CreatedAt, UpdatedAt
            FROM Clients
            ORDER BY Name
        """)
    else:
        # Regular users only see their own clients
        cursor.execute("""
            SELECT Id, Name, Email, Phone, Address, UserId, CreatedAt, UpdatedAt
            FROM Clients
            WHERE UserId = ?
            ORDER BY Name
        """, (user_id,))
    
    rows = cursor.fetchall()
    clients = []
    for row in rows:
        clients.append({
            'id': row[0],
            'name': row[1],
            'email': row[2],
            'phone': row[3],
            'address': row[4],
            'userId': row[5],
            'createdAt': row[6].isoformat() if row[6] else None,
            'updatedAt': row[7].isoformat() if row[7] else None
        })
    
    conn.close()
    return jsonify(clients)

@app.route('/clients', methods=['POST'])
@require_auth
def create_client():
    """Create a new client for the current user"""
    user = get_user_from_token()
    user_id = user.get('user_id')
    
    data = request.get_json()
    if not data or not data.get('name'):
        return jsonify({"error": "Client name is required"}), 400
    
    name = data.get('name')
    email = data.get('email', '')
    phone = data.get('phone', '')
    address = data.get('address', '')
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute("""
            INSERT INTO Clients (Name, Email, Phone, Address, UserId, CreatedAt, UpdatedAt)
            VALUES (?, ?, ?, ?, ?, GETDATE(), GETDATE())
        """, (name, email, phone, address, user_id))
        
        conn.commit()
        
        cursor.execute("SELECT SCOPE_IDENTITY()")
        client_id = cursor.fetchone()[0]
        
        conn.close()
        
        log_user_action('client_created', {
            'client_id': client_id,
            'name': name
        })
        
        return jsonify({
            "status": "success",
            "id": client_id,
            "name": name,
            "userId": user_id
        }), 201
        
    except Exception as e:
        conn.close()
        logger.error(f"Error creating client: {e}")
        return jsonify({"error": "Failed to create client"}), 500

@app.route('/clients/<int:client_id>', methods=['GET'])
@require_auth
def get_client_by_id(client_id):
    """Get a specific client (with ownership verification)"""
    user = get_user_from_token()
    user_id = user.get('user_id')
    user_roles = user.get('roles', [])
    is_admin = 'Admin' in user_roles or 'Manager' in user_roles
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    if is_admin:
        cursor.execute("""
            SELECT Id, Name, Email, Phone, Address, UserId, CreatedAt, UpdatedAt
            FROM Clients WHERE Id = ?
        """, (client_id,))
    else:
        cursor.execute("""
            SELECT Id, Name, Email, Phone, Address, UserId, CreatedAt, UpdatedAt
            FROM Clients WHERE Id = ? AND UserId = ?
        """, (client_id, user_id))
    
    row = cursor.fetchone()
    conn.close()
    
    if not row:
        return jsonify({"error": "Client not found or access denied"}), 404
    
    return jsonify({
        'id': row[0],
        'name': row[1],
        'email': row[2],
        'phone': row[3],
        'address': row[4],
        'userId': row[5],
        'createdAt': row[6].isoformat() if row[6] else None,
        'updatedAt': row[7].isoformat() if row[7] else None
    })

@app.route('/clients/<int:client_id>', methods=['PUT'])
@require_auth
def update_client(client_id):
    """Update a client (with ownership verification)"""
    user = get_user_from_token()
    user_id = user.get('user_id')
    
    if not verify_client_ownership(client_id, user_id):
        return jsonify({"error": "Access denied"}), 403
    
    data = request.get_json()
    if not data:
        return jsonify({"error": "No data provided"}), 400
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    update_fields = []
    params = []
    
    if 'name' in data:
        update_fields.append("Name = ?")
        params.append(data['name'])
    if 'email' in data:
        update_fields.append("Email = ?")
        params.append(data['email'])
    if 'phone' in data:
        update_fields.append("Phone = ?")
        params.append(data['phone'])
    if 'address' in data:
        update_fields.append("Address = ?")
        params.append(data['address'])
    
    if not update_fields:
        conn.close()
        return jsonify({"error": "No valid fields to update"}), 400
    
    update_fields.append("UpdatedAt = GETDATE()")
    params.append(client_id)
    
    query = f"UPDATE Clients SET {', '.join(update_fields)} WHERE Id = ?"
    cursor.execute(query, params)
    conn.commit()
    conn.close()
    
    return jsonify({"status": "success"}), 200

@app.route('/clients/<int:client_id>', methods=['DELETE'])
@require_auth
def delete_client(client_id):
    """Delete a client (with ownership verification)"""
    user = get_user_from_token()
    user_id = user.get('user_id')
    
    if not verify_client_ownership(client_id, user_id):
        return jsonify({"error": "Access denied"}), 403
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Check if client has documents
    cursor.execute("SELECT COUNT(*) FROM cases WHERE client_id = ?", (client_id,))
    doc_count = cursor.fetchone()[0]
    
    if doc_count > 0:
        conn.close()
        return jsonify({
            "error": f"Cannot delete client with {doc_count} documents. Please delete or reassign documents first."
        }), 400
    
    cursor.execute("DELETE FROM Clients WHERE Id = ?", (client_id,))
    conn.commit()
    conn.close()
    
    return jsonify({"status": "success"}), 200

@app.route('/microservices/health', methods=['GET'])
def check_microservices_health():
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM cases")
        db_count = cursor.fetchone()[0]
        conn.close()

        gemini_status = "healthy"
        api_key = os.getenv("GEMINI_API_KEY")
        if not api_key:
            gemini_status = "no_api_key"
        else:
            url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
            try:
                response = requests.get(url, timeout=10)
                gemini_status = "healthy" if response.status_code == 200 else "api_error"
            except:
                gemini_status = "connection_error"

        upload_folder_writable = os.access(UPLOAD_FOLDER, os.W_OK)
        overall_status = "healthy" if gemini_status == "healthy" and upload_folder_writable else "degraded"

        return jsonify({
            "overall_status": overall_status,
            "services": {
                "database": {"status": "healthy", "document_count": db_count},
                "gemini_api": {"status": gemini_status},
                "file_system": {"status": "healthy" if upload_folder_writable else "error", "upload_folder": UPLOAD_FOLDER, "writable": upload_folder_writable}
            },
            "timestamp": datetime.now().isoformat()
        }), 200
    except Exception as e:
        return jsonify({"overall_status": "unhealthy", "error": str(e), "timestamp": datetime.now().isoformat()}), 500

@app.route('/practice-areas', methods=['GET'])
@require_auth
def get_practice_areas():
    user = get_user_from_token()
    user_id = user.get('user_id')
    user_roles = user.get('roles', [])
    is_admin = 'Admin' in user_roles or 'Manager' in user_roles
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    if is_admin:
        cursor.execute("""
            SELECT practice_area, COUNT(*) as count
            FROM cases
            WHERE practice_area IS NOT NULL AND practice_area != ''
            GROUP BY practice_area
            ORDER BY count DESC
        """)
    else:
        cursor.execute("""
            SELECT practice_area, COUNT(*) as count
            FROM cases
            WHERE practice_area IS NOT NULL AND practice_area != '' AND user_id = ?
            GROUP BY practice_area
            ORDER BY count DESC
        """, (user_id,))
    
    rows = cursor.fetchall()
    practice_areas = [{"area": row[0], "count": row[1]} for row in rows]
    conn.close()
    return jsonify(practice_areas)

@app.route('/quick-filters', methods=['GET'])
@require_auth
def get_quick_filter_stats():
    user = get_user_from_token()
    user_id = user.get('user_id')
    user_roles = user.get('roles', [])
    is_admin = 'Admin' in user_roles or 'Manager' in user_roles
    
    client_id = request.args.get('clientId')
    
    conn = get_db_connection()
    cursor = conn.cursor()
    today = datetime.now().strftime("%Y-%m-%d")
    tomorrow = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")
    
    where_clause = ""
    params = []
    
    if not is_admin:
        where_clause = "WHERE user_id = ?"
        params = [user_id]
    
    if client_id:
        client_id_int = int(client_id)
        if not is_admin and not verify_client_ownership(client_id_int, user_id):
            return jsonify({"error": "Access denied"}), 403
        
        if where_clause:
            where_clause += " AND client_id = ?"
        else:
            where_clause = "WHERE client_id = ?"
        params.append(client_id_int)
    
    stats = {}
    
    # High priority
    query = f"SELECT COUNT(*) FROM cases {where_clause}"
    if where_clause:
        query += " AND priority IN ('high', 'urgent')"
    else:
        query += " WHERE priority IN ('high', 'urgent')"
    cursor.execute(query, params)
    stats['high_priority'] = cursor.fetchone()[0]
    
    # Completed today
    query = f"SELECT COUNT(*) FROM cases {where_clause}"
    today_params = params + [today, tomorrow]
    if where_clause:
        query += " AND analyzed_at >= ? AND analyzed_at < ?"
    else:
        query += " WHERE analyzed_at >= ? AND analyzed_at < ?"
    cursor.execute(query, today_params)
    stats['completed_today'] = cursor.fetchone()[0]
    
    # Needs review
    query = f"SELECT COUNT(*) FROM cases {where_clause}"
    if where_clause:
        query += " AND needs_review = 1"
    else:
        query += " WHERE needs_review = 1"
    cursor.execute(query, params)
    stats['needs_review'] = cursor.fetchone()[0]
    
    # Processing errors
    query = f"SELECT COUNT(*) FROM cases {where_clause}"
    if where_clause:
        query += " AND status = 'Error'"
    else:
        query += " WHERE status = 'Error'"
    cursor.execute(query, params)
    stats['processing_errors'] = cursor.fetchone()[0]
    
    # Low confidence
    query = f"SELECT COUNT(*) FROM cases {where_clause}"
    if where_clause:
        query += " AND confidence_score < 0.6 AND confidence_score IS NOT NULL"
    else:
        query += " WHERE confidence_score < 0.6 AND confidence_score IS NOT NULL"
    cursor.execute(query, params)
    stats['low_confidence'] = cursor.fetchone()[0]
    
    conn.close()
    return jsonify(stats)

@app.route('/cases/<int:case_id>/update-metadata', methods=['PATCH'])
@require_auth
def update_case_enhanced_metadata(case_id):
    user = get_user_from_token()
    user_id = user.get('user_id')
    user_roles = user.get('roles', [])
    is_admin = 'Admin' in user_roles or 'Manager' in user_roles
    
    data = request.get_json()
    if not data:
        return jsonify({"error": "No data provided"}), 400
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Verify ownership
    if is_admin:
        cursor.execute("SELECT id FROM cases WHERE id = ?", (case_id,))
    else:
        cursor.execute("SELECT id FROM cases WHERE id = ? AND user_id = ?", (case_id, user_id))
    
    if not cursor.fetchone():
        conn.close()
        return jsonify({"error": "Case not found or access denied"}), 404
    
    allowed_fields = {
        'title': 'title', 
        'classification': 'classification', 
        'enable_ocr': 'enable_ocr',
        'practice_area': 'practice_area', 
        'priority': 'priority', 
        'needs_review': 'needs_review', 
        'tags': 'tags'
    }
    
    update_fields = []
    update_values = []
    for field, db_col in allowed_fields.items():
        if field in data:
            update_fields.append(f"{db_col} = ?")
            update_values.append(data[field])
    
    if not update_fields:
        conn.close()
        return jsonify({"error": "No valid fields to update"}), 400
    
    update_fields.append("updated_at = GETDATE()")
    update_values.append(case_id)
    
    query = f"UPDATE cases SET {', '.join(update_fields)} WHERE id = ?"
    cursor.execute(query, update_values)
    conn.commit()
    conn.close()
    return jsonify({"status": "success", "message": "Case metadata updated successfully"}), 200

@app.route('/batch-upload', methods=['POST'])
@require_auth
def batch_upload_documents():
    user = get_user_from_token()
    user_id = user.get('user_id')
    
    client_id = request.form.get('clientId')
    if not client_id:
        return jsonify({"error": "Client ID is required"}), 400
    
    try:
        client_id = int(client_id)
        if not verify_client_ownership(client_id, user_id):
            return jsonify({"error": "Invalid client ID or access denied"}), 403
    except ValueError:
        return jsonify({"error": "Invalid client ID format"}), 400
    
    if 'files' not in request.files:
        return jsonify({"error": "No files provided"}), 400
    
    files = request.files.getlist('files')
    if not files or all(f.filename == '' for f in files):
        return jsonify({"error": "No files selected"}), 400

    titles = request.form.getlist('titles')
    classifications = request.form.getlist('classifications')
    languages = request.form.getlist('languages')
    priorities = request.form.getlist('priorities')
    practice_areas = request.form.getlist('practiceAreas')
    enable_ocr = request.form.get('enableOCR', 'true').lower() == 'true'

    results = []
    for i, file in enumerate(files):
        if file.filename == '':
            continue
        try:
            title = titles[i] if i < len(titles) else file.filename
            classification = classifications[i] if i < len(classifications) else 'auto'
            language = languages[i] if i < len(languages) else 'en'
            priority = priorities[i] if i < len(priorities) else 'normal'
            practice_area = practice_areas[i] if i < len(practice_areas) else ''

            start_time = time.time()
            filename = file.filename
            original_name = filename
            file_size = len(file.read())
            file.seek(0)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)

            raw_text = extract_text_with_ocr(filepath)
            if raw_text is None:
                results.append({"filename": filename, "status": "error", "error": "Could not extract text"})
                continue

            cleaned_text = clean_extracted_text(raw_text)
            analysis = perform_gemini_analysis(cleaned_text)
            analysis_duration = int((time.time() - start_time) * 1000)
            analyzed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            if analysis is None:
                results.append({"filename": filename, "status": "error", "error": "AI analysis failed"})
                continue

            if not practice_area:
                practice_area = detect_practice_area(cleaned_text, analysis.get('document_type'), analysis.get('parties'))
            if priority == 'normal':
                detected_priority = determine_priority(cleaned_text, analysis.get('document_type'), analysis.get('parties'))
                if detected_priority != 'normal':
                    priority = detected_priority

            confidence_score = calculate_confidence_score(analysis, len(cleaned_text))
            needs_review = confidence_score < 0.7

            conn = get_db_connection()
            cursor = conn.cursor()
            creation_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            arguments_str = "||".join(analysis.get('arguments', []))
            try:
                cursor.execute("""
                    INSERT INTO cases(filename, original_name, title, summary, parties, document_date, court, arguments, document_type,
                    creation_date, document_language, file_size, extracted_text, analysis_duration_ms, classification, enable_ocr, file_path,
                    status, practice_area, priority, confidence_score, needs_review, analyzed_at, user_id, client_id)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Analyzed', ?, ?, ?, ?, ?, ?, ?)
                """, (filename, original_name, title, analysis.get('summary'), analysis.get('parties'), analysis.get('document_date'),
                      analysis.get('court'), arguments_str, analysis.get('document_type'), creation_date, analysis.get('document_language'),
                      file_size, raw_text, analysis_duration, classification, enable_ocr, filepath, practice_area, priority,
                      confidence_score, needs_review, analyzed_at, user_id, client_id))
                conn.commit()
                cursor.execute("SELECT SCOPE_IDENTITY()")
                document_id = cursor.fetchone()[0]
                results.append({
                    "filename": filename,
                    "status": "success",
                    "id": document_id,
                    "analysis_duration_ms": analysis_duration,
                    "practice_area": practice_area,
                    "priority": priority,
                    "confidence_score": confidence_score
                })
            except pyodbc.IntegrityError:
                results.append({"filename": filename, "status": "error", "error": "File already exists"})
            conn.close()
        except Exception as e:
            results.append({"filename": filename, "status": "error", "error": str(e)})
    return jsonify({"results": results}), 200

@app.route('/cases/<int:case_id>/reanalyze', methods=['POST'])
@require_auth
def reanalyze_single_document(case_id):
    user = get_user_from_token()
    user_id = user.get('user_id')
    user_roles = user.get('roles', [])
    is_admin = 'Admin' in user_roles or 'Manager' in user_roles
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Verify ownership
    if is_admin:
        cursor.execute("SELECT id, filename, file_path FROM cases WHERE id = ?", (case_id,))
    else:
        cursor.execute("SELECT id, filename, file_path FROM cases WHERE id = ? AND user_id = ?", (case_id, user_id))
    
    row = cursor.fetchone()
    if not row:
        conn.close()
        return jsonify({"error": "Case not found or access denied"}), 404
    
    _, _, file_path = row
    if not file_path or not os.path.exists(file_path):
        conn.close()
        return jsonify({"error": "Source file not found"}), 404

    try:
        start_time = time.time()
        raw_text = extract_text_with_ocr(file_path)
        if raw_text is None:
            raise Exception("Could not re-extract text from file")
        cleaned_text = clean_extracted_text(raw_text)
        analysis = perform_gemini_analysis(cleaned_text)
        if analysis is None:
            raise Exception("AI re-analysis failed")

        analysis_duration = int((time.time() - start_time) * 1000)
        analyzed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        practice_area = detect_practice_area(cleaned_text, analysis.get('document_type'), analysis.get('parties'))
        priority = determine_priority(cleaned_text, analysis.get('document_type'), analysis.get('parties'))
        confidence_score = calculate_confidence_score(analysis, len(cleaned_text))
        needs_review = confidence_score < 0.7
        arguments_str = "||".join(analysis.get('arguments', []))

        cursor.execute("""
            UPDATE cases SET summary = ?, parties = ?, document_date = ?, court = ?, arguments = ?, document_type = ?, document_language = ?,
            extracted_text = ?, analysis_duration_ms = ?, status = 'Analyzed', practice_area = ?, priority = ?, confidence_score = ?,
            needs_review = ?, updated_at = GETDATE(), analyzed_at = ?
            WHERE id = ?
        """, (analysis.get('summary'), analysis.get('parties'), analysis.get('document_date'), analysis.get('court'),
              arguments_str, analysis.get('document_type'), analysis.get('document_language'), raw_text, analysis_duration,
              practice_area, priority, confidence_score, needs_review, analyzed_at, case_id))
        conn.commit()
        conn.close()
        return jsonify({
            "status": "success",
            "message": "Document re-analyzed successfully",
            "id": case_id,
            "analysis_duration_ms": analysis_duration,
            "practice_area": practice_area,
            "priority": priority,
            "confidence_score": confidence_score
        }), 200
    except Exception as e:
        cursor.execute("UPDATE cases SET status = 'Error' WHERE id = ?", (case_id,))
        conn.commit()
        conn.close()
        return jsonify({"error": str(e)}), 500

@app.route('/reanalyze-all', methods=['POST'])
@require_auth
@require_role('Admin')
def reanalyze_all_documents():
    """Admin-only: Re-analyze all documents"""
    user = get_user_from_token()
    user_id = user.get('user_id')
    
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, filename, file_path, enable_ocr FROM cases")
    rows = cursor.fetchall()
    updated_count = 0
    error_count = 0
    for row in rows:
        case_id, filename, file_path, _ = row
        if not file_path or not os.path.exists(file_path):
            error_count += 1
            continue
        try:
            start_time = time.time()
            raw_text = extract_text_with_ocr(file_path)
            if raw_text is None:
                cursor.execute("UPDATE cases SET status = 'Error' WHERE id = ?", (case_id,))
                error_count += 1
                continue
            cleaned_text = clean_extracted_text(raw_text)
            analysis = perform_gemini_analysis(cleaned_text)
            if analysis is None:
                cursor.execute("UPDATE cases SET status = 'Error' WHERE id = ?", (case_id,))
                error_count += 1
                continue
            analysis_duration = int((time.time() - start_time) * 1000)
            analyzed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            practice_area = detect_practice_area(cleaned_text, analysis.get('document_type'), analysis.get('parties'))
            priority = determine_priority(cleaned_text, analysis.get('document_type'), analysis.get('parties'))
            confidence_score = calculate_confidence_score(analysis, len(cleaned_text))
            needs_review = confidence_score < 0.7
            arguments_str = "||".join(analysis.get('arguments', []))
            update_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cursor.execute("""
                UPDATE cases SET summary = ?, parties = ?, document_date = ?, court = ?, arguments = ?, document_type = ?, document_language = ?,
                extracted_text = ?, analysis_duration_ms = ?, status = 'Analyzed', creation_date = ?, practice_area = ?, priority = ?,
                confidence_score = ?, needs_review = ?, updated_at = GETDATE(), analyzed_at = ?
                WHERE id = ?
            """, (analysis.get('summary'), analysis.get('parties'), analysis.get('document_date'), analysis.get('court'),
                  arguments_str, analysis.get('document_type'), analysis.get('document_language'), raw_text, analysis_duration,
                  update_date, practice_area, priority, confidence_score, needs_review, analyzed_at, case_id))
            conn.commit()
            updated_count += 1
        except:
            cursor.execute("UPDATE cases SET status = 'Error' WHERE id = ?", (case_id,))
            error_count += 1
    conn.close()
    return jsonify({
        "status": "completed",
        "updated_count": updated_count,
        "error_count": error_count,
        "message": f"Re-analyzed {updated_count} documents with {error_count} errors"
    }), 200

@app.route('/cases/<int:case_id>', methods=['GET'])
@require_auth
def get_case_by_id(case_id):
    user = get_user_from_token()
    user_id = user.get('user_id')
    user_roles = user.get('roles', [])
    is_admin = 'Admin' in user_roles or 'Manager' in user_roles
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    if is_admin:
        cursor.execute("""
            SELECT id, filename, original_name, title, summary, parties, document_date, court, arguments, document_type,
                   creation_date, document_language, file_size, extracted_text, analysis_duration_ms, classification, status,
                   practice_area, priority, confidence_score, needs_review, tags, updated_at, analyzed_at, user_id, client_id
            FROM cases WHERE id = ?
        """, (case_id,))
    else:
        cursor.execute("""
            SELECT id, filename, original_name, title, summary, parties, document_date, court, arguments, document_type,
                   creation_date, document_language, file_size, extracted_text, analysis_duration_ms, classification, status,
                   practice_area, priority, confidence_score, needs_review, tags, updated_at, analyzed_at, user_id, client_id
            FROM cases WHERE id = ? AND user_id = ?
        """, (case_id, user_id))
    
    row = cursor.fetchone()
    if not row:
        conn.close()
        return jsonify({"error": "Case not found or access denied"}), 404
    
    case = dict(zip([col[0] for col in cursor.description], row))
    for key, value in case.items():
        if isinstance(value, datetime):
            case[key] = value.isoformat()
    conn.close()
    return jsonify(case)

@app.route('/cases/<int:case_id>', methods=['DELETE'])
@require_auth
def delete_case(case_id):
    user = get_user_from_token()
    user_id = user.get('user_id')
    user_roles = user.get('roles', [])
    is_admin = 'Admin' in user_roles or 'Manager' in user_roles
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Verify ownership
    if is_admin:
        cursor.execute("SELECT file_path FROM cases WHERE id = ?", (case_id,))
    else:
        cursor.execute("SELECT file_path FROM cases WHERE id = ? AND user_id = ?", (case_id, user_id))
    
    row = cursor.fetchone()
    if not row:
        conn.close()
        return jsonify({"error": "Case not found or access denied"}), 404
    
    if row[0]:
        file_path = row[0]
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
                logger.info(f"Deleted file: {file_path}")
            except Exception as e:
                logger.error(f"Failed to delete file {file_path}: {e}")
    
    cursor.execute("DELETE FROM cases WHERE id = ?", (case_id,))
    conn.commit()
    conn.close()
    return jsonify({"status": "deleted"}), 200

@app.route('/search', methods=['GET'])
@require_auth
def search_cases():
    user = get_user_from_token()
    user_id = user.get('user_id')
    user_roles = user.get('roles', [])
    is_admin = 'Admin' in user_roles or 'Manager' in user_roles
    
    keyword = request.args.get('keyword', '') or request.args.get('query', '')
    client_id = request.args.get('clientId')
    
    if not keyword:
        return jsonify([])

    query_param = f"%{keyword}%"
    conn = get_db_connection()
    cursor = conn.cursor()

    where_conditions = [
        "(filename LIKE ? OR title LIKE ? OR summary LIKE ? OR parties LIKE ? OR court LIKE ? OR document_type LIKE ? OR practice_area LIKE ? OR tags LIKE ?)"
    ]
    params = [query_param] * 8
    
    # User filter
    if not is_admin:
        where_conditions.append("user_id = ?")
        params.append(user_id)
    
    # Client filter
    if client_id:
        try:
            client_id_int = int(client_id)
            if not is_admin and not verify_client_ownership(client_id_int, user_id):
                return jsonify({"error": "Access denied"}), 403
            
            where_conditions.append("client_id = ?")
            params.append(client_id_int)
        except ValueError:
            return jsonify({"error": "Invalid client ID"}), 400

    cursor.execute(f"""
        SELECT id, filename, original_name, title, summary, parties, document_date, court, 
               arguments, document_type, creation_date, document_language, file_size, 
               analysis_duration_ms, status, practice_area, priority, confidence_score, 
               needs_review, tags, user_id, client_id
        FROM cases
        WHERE {' AND '.join(where_conditions)}
        ORDER BY id DESC
    """, params)
    
    rows = cursor.fetchall()

    cases = []
    for row in rows:
        case_dict = dict(zip([col[0] for col in cursor.description], row))
        for key, value in case_dict.items():
            if isinstance(value, datetime):
                case_dict[key] = value.isoformat()
        cases.append(case_dict)

    conn.close()
    return jsonify(cases)

@app.route('/trends', methods=['GET'])
@require_auth
def get_trends():
    user = get_user_from_token()
    user_id = user.get('user_id')
    user_roles = user.get('roles', [])
    is_admin = 'Admin' in user_roles or 'Manager' in user_roles
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    where_clause = "WHERE status='Analyzed'"
    params = []
    if not is_admin:
        where_clause += " AND user_id = ?"
        params.append(user_id)
    
    cursor.execute(f"SELECT parties FROM cases {where_clause}", params)
    all_entities = []
    for row in cursor.fetchall():
        parties = row[0] or ""
        entities = re.split(r',\s*|\s+and\s+', parties)
        cleaned = [re.sub(r'\(.*?\)|^(APPELLANT:|RESPONDENT:|JUDGE:)', '', e).strip() for e in entities]
        all_entities.extend([name for name in cleaned if len(name) > 3])
    entity_counts = Counter(all_entities).most_common(10)
    
    cursor.execute(f"""
        SELECT practice_area, COUNT(*) FROM cases {where_clause} 
        AND practice_area IS NOT NULL 
        GROUP BY practice_area 
        ORDER BY COUNT(*) DESC
    """, params)
    practice_area_trends = [{"area": row[0], "count": row[1]} for row in cursor.fetchall()]
    
    cursor.execute(f"""
        SELECT document_type, COUNT(*) FROM cases {where_clause}
        GROUP BY document_type 
        ORDER BY COUNT(*) DESC
    """, params)
    document_type_trends = [{"type": row[0], "count": row[1]} for row in cursor.fetchall()]
    
    conn.close()
    return jsonify({
        "top_entities": entity_counts,
        "practice_area_trends": practice_area_trends,
        "document_type_trends": document_type_trends
    })

@app.route('/extract-text', methods=['POST'])
def extract_text_endpoint():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    filename = file.filename
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)
    try:
        start_time = time.time()
        text = extract_text_with_ocr(filepath)
        extraction_time = int((time.time() - start_time) * 1000)
        if text is None:
            return jsonify({"error": "Could not extract text from file"}), 500
        return jsonify({
            "text": text,
            "filename": filename,
            "extraction_time_ms": extraction_time,
            "text_length": len(text)
        }), 200
    finally:
        if os.path.exists(filepath):
            os.remove(filepath)

@app.route('/health', methods=['GET'])
def health_check():
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM cases")
        conn.close()
        return jsonify({
            "status": "healthy",
            "database": "connected",
            "timestamp": datetime.now().isoformat()
        }), 200
    except Exception as e:
        return jsonify({
            "status": "unhealthy",
            "database": "disconnected",
            "error": str(e),
            "timestamp": datetime.now().isoformat()
        }), 500

@app.route('/admin/audit-logs', methods=['GET'])
@require_role('Admin', 'Manager')
def get_audit_logs():
    """Get audit logs - Admin/Manager only"""
    start_date = request.args.get('startDate')
    end_date = request.args.get('endDate')
    user_id = request.args.get('userId')
    action = request.args.get('action', 'all')
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    where_conditions = []
    params = []
    
    if start_date:
        where_conditions.append("Timestamp >= ?")
        params.append(start_date)
    
    if end_date:
        where_conditions.append("Timestamp <= ?")
        params.append(end_date)
    
    if user_id:
        where_conditions.append("UserId = ?")
        params.append(user_id)
    
    if action != 'all':
        where_conditions.append("Action = ?")
        params.append(action)
    
    query = """
        SELECT Id, UserId, UserEmail, Action, Details, IpAddress, Timestamp
        FROM AuditLogs
    """
    
    if where_conditions:
        query += " WHERE " + " AND ".join(where_conditions)
    
    query += " ORDER BY Timestamp DESC"
    
    try:
        cursor.execute(query, params)
        rows = cursor.fetchall()
        
        logs = []
        for row in rows:
            logs.append({
                'id': row[0],
                'userId': row[1],
                'userEmail': row[2],
                'action': row[3],
                'details': row[4],
                'ipAddress': row[5],
                'timestamp': row[6].isoformat() if row[6] else None
            })
        
        conn.close()
        return jsonify({'logs': logs, 'total': len(logs)}), 200
    except Exception as e:
        logger.error(f"Error fetching audit logs: {e}")
        return jsonify({'error': 'Failed to fetch audit logs'}), 500

@app.errorhandler(404)
def not_found(error):
    return jsonify({"error": "Endpoint not found"}), 404

@app.errorhandler(500)
def internal_error(error):
    return jsonify({"error": "Internal server error"}), 500

if __name__ == '__main__':
    init_db()
    initialize_compare_service()
    initialize_document_generation_service()
    initialize_case_analysis_service()
    # Disable reloader to prevent double initialization
    app.run(debug=True, port=3001, use_reloader=False)