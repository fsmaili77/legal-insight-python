# compare_service.py
import os
import difflib
import numpy as np
import re
import pyodbc
import json
import requests
from datetime import datetime
from flask import Blueprint, request, jsonify
from werkzeug.utils import secure_filename
from PyPDF2 import PdfReader
import docx
from sentence_transformers import SentenceTransformer, util
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

compare_service = Blueprint('compare_service', __name__)

# Load model once
try:
    model = SentenceTransformer('all-MiniLM-L6-v2')
    logger.info("SentenceTransformer model loaded successfully")
except Exception as e:
    logger.error(f"Failed to load SentenceTransformer model: {e}")
    model = None

def get_db_connection():
    conn_str = os.getenv("SQLSERVER_CONN_STRING")
    if not conn_str:
        raise ValueError("SQLSERVER_CONN_STRING not found in environment variables")
    return pyodbc.connect(conn_str)

def create_document_comparisons_table():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        IF NOT EXISTS (
            SELECT * FROM INFORMATION_SCHEMA.TABLES 
            WHERE TABLE_NAME = 'document_comparisons'
        )
        CREATE TABLE document_comparisons(
            id INT IDENTITY(1,1) PRIMARY KEY,
            doc1_id INT NULL,
            doc2_id INT NULL,
            doc1_name NVARCHAR(255),
            doc2_name NVARCHAR(255),
            compared_at DATETIME NOT NULL DEFAULT GETDATE(),
            comparison_type VARCHAR(50) NOT NULL,
            summary NVARCHAR(MAX),
            semantic_diff NVARCHAR(MAX),
            gemini_summary NVARCHAR(MAX),
            similarity_score FLOAT,
            key_differences NVARCHAR(MAX),
            created_by NVARCHAR(100) DEFAULT 'system',
            status VARCHAR(20) DEFAULT 'completed'
        )
    """)
    conn.commit()

    new_columns = [
        ("doc1_name", "NVARCHAR(255)"),
        ("doc2_name", "NVARCHAR(255)"),
        ("similarity_score", "FLOAT"),
        ("key_differences", "NVARCHAR(MAX)"),
        ("created_by", "NVARCHAR(100) DEFAULT 'system'"),
        ("status", "VARCHAR(20) DEFAULT 'completed'")
    ]

    for col, defn in new_columns:
        try:
            cursor.execute(f"IF COL_LENGTH('document_comparisons', '{col}') IS NULL ALTER TABLE document_comparisons ADD {col} {defn}")
            logger.info(f"Added column: {col}")
        except Exception as e:
            logger.debug(f"Column {col} already exists: {e}")

    conn.commit()
    conn.close()

def get_extracted_text_from_db(doc_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT extracted_text, filename FROM cases WHERE id = ?", (doc_id,))
    row = cursor.fetchone()
    conn.close()
    return (row[0], row[1]) if row else (None, None)

def save_document_comparison(doc1_id, doc2_id, doc1_name, doc2_name, comparison_type, summary, semantic_diff, gemini_summary, similarity_score=None, key_differences=None):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        key_differences_val = "\n".join(key_differences) if isinstance(key_differences, list) else key_differences

        # Use OUTPUT to get the inserted ID in one atomic statement
        cursor.execute("""
            INSERT INTO document_comparisons(
                doc1_id, doc2_id, doc1_name, doc2_name, compared_at, comparison_type,
                summary, semantic_diff, gemini_summary, similarity_score, key_differences
            )
            OUTPUT INSERTED.id
            VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            doc1_id,
            doc2_id,
            doc1_name[:255] if doc1_name else None,  # Enforce NVARCHAR(255) limit
            doc2_name[:255] if doc2_name else None,
            datetime.now(),
            comparison_type[:50] if comparison_type else 'advanced',
            summary[:4000] if summary else None,  # Optional: cap summary
            json.dumps(semantic_diff, ensure_ascii=False) if semantic_diff else None,
            gemini_summary[:8000] if gemini_summary and len(gemini_summary) > 8000 else gemini_summary,  # Prevent extreme size
            similarity_score,
            key_differences_val[:8000] if key_differences_val and len(key_differences_val) > 8000 else key_differences_val
        ))

        row = cursor.fetchone()
        if row is None:
            conn.rollback()
            logger.error("OUTPUT clause returned no row after INSERT.")
            raise RuntimeError("Database insert failed to return an ID.")

        comparison_id = int(row[0])
        conn.commit()
        return comparison_id

    except pyodbc.Error as db_err:
        conn.rollback()
        logger.error(f"Database error in save_document_comparison: {db_err}")
        raise RuntimeError(f"Database error: {str(db_err)}")
    except Exception as e:
        conn.rollback()
        logger.error(f"Unexpected error in save_document_comparison: {e}", exc_info=True)
        raise
    finally:
        conn.close()

def extract_text_from_pdf(pdf_file):
    reader = PdfReader(pdf_file)
    text = ""
    for page in reader.pages:
        ptxt = page.extract_text() or ""
        text += ptxt
    return text, []

def extract_text_from_docx(docx_file):
    doc = docx.Document(docx_file)
    text = "\n".join(p.text for p in doc.paragraphs)
    return text, []

def read_text_and_tables(file_storage):
    filename = secure_filename(file_storage.filename)
    if filename.lower().endswith('.pdf'):
        return extract_text_from_pdf(file_storage)
    elif filename.lower().endswith('.docx'):
        return extract_text_from_docx(file_storage)
    else:
        return file_storage.read().decode("utf-8"), []

def split_into_clauses(text):
    if not text:
        return []
    patterns = [
        r'(?:Article|Section|Clause|Art\.|Sec\.)\s*\d+\.?',
        r'\d+\.\s+',
        r'[A-Z][A-Z\s]{10,}:',
        r'\n\s*\([a-z]\)',
        r'\n\s*\([0-9]+\)'
    ]
    clauses = []
    current_clause = ""
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue
        is_new_clause = any(re.match(pattern, line, re.IGNORECASE) for pattern in patterns)
        if is_new_clause and current_clause:
            clauses.append(current_clause.strip())
            current_clause = line
        else:
            current_clause += " " + line if current_clause else line
    if current_clause:
        clauses.append(current_clause.strip())
    return [c for c in clauses if len(c) > 20]

def semantic_clause_diff(clauses1, clauses2, threshold=0.75):
    if not model or not clauses1 or not clauses2:
        return []
    try:
        emb1 = model.encode(clauses1)
        emb2 = model.encode(clauses2)
        sem_diff = []
        taken = set()
        for idx1, v1 in enumerate(emb1):
            scores = util.cos_sim([v1], emb2)[0].cpu().numpy()
            idx2 = int(np.argmax(scores))
            score = float(scores[idx2])
            if score > threshold and idx2 not in taken:
                taken.add(idx2)
                sem_diff.append({
                    'doc1_clause_idx': idx1,
                    'doc2_clause_idx': idx2,
                    'doc1_clause': clauses1[idx1],
                    'doc2_clause': clauses2[idx2],
                    'semantic_similarity': score,
                    'change_type': 'modified' if score < 0.95 else 'identical'
                })
            else:
                sem_diff.append({
                    'doc1_clause_idx': idx1,
                    'doc2_clause_idx': None,
                    'doc1_clause': clauses1[idx1],
                    'doc2_clause': None,
                    'semantic_similarity': score if idx2 not in taken else 0.0,
                    'change_type': 'removed'
                })
        matched = {item['doc2_clause_idx'] for item in sem_diff if item['doc2_clause_idx'] is not None}
        for idx2 in range(len(clauses2)):
            if idx2 not in matched:
                sem_diff.append({
                    'doc1_clause_idx': None,
                    'doc2_clause_idx': idx2,
                    'doc1_clause': None,
                    'doc2_clause': clauses2[idx2],
                    'semantic_similarity': 0.0,
                    'change_type': 'added'
                })
        return sem_diff
    except Exception as e:
        logger.error(f"Error in semantic clause comparison: {e}")
        return []

def calculate_overall_similarity(semantic_diff):
    if not semantic_diff:
        return 0.0
    total = len(semantic_diff)
    similarity_sum = sum(item['semantic_similarity'] for item in semantic_diff)
    return similarity_sum / total

def extract_key_differences(semantic_diff, limit=10):
    diffs = []
    for item in semantic_diff:
        if item['change_type'] == 'removed':
            diffs.append(f"Removed: {item['doc1_clause'][:100]}...")
        elif item['change_type'] == 'added':
            diffs.append(f"Added: {item['doc2_clause'][:100]}...")
        elif item['change_type'] == 'modified' and item['semantic_similarity'] < 0.85:
            diffs.append(f"Modified: {item['doc1_clause'][:50]}... → {item['doc2_clause'][:50]}...")
    return diffs[:limit]

def gemini_advanced_compare(text1, text2, doc1_name="Document 1", doc2_name="Document 2"):
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
    if not GEMINI_API_KEY:
        return "Gemini API key not configured"
    max_length = 15000
    text1 = text1[:max_length] + "\n\n[Document truncated due to length limits]" if len(text1) > max_length else text1
    text2 = text2[:max_length] + "\n\n[Document truncated due to length limits]" if len(text2) > max_length else text2
    prompt = f"""As a legal AI assistant, compare these two legal documents and provide a comprehensive analysis.
DOCUMENT 1 ({doc1_name}): {text1}
DOCUMENT 2 ({doc2_name}): {text2}
Please provide:
1. Executive Summary of changes
2. Key additions in Document 2
3. Key removals from Document 1
4. Modifications to existing content
5. Legal implications of changes
6. Risk assessment
7. Recommendations for review
Focus on substantive changes that could affect legal interpretation or obligations.
Limit response to 2000 words for clarity."""
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_API_KEY}"
        # url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.1, "topK": 40, "topP": 0.95, "maxOutputTokens": 4000}
        }
        response = requests.post(url, json=payload, timeout=60)
        response.raise_for_status()
        result = response.json()
        if 'candidates' in result and result['candidates']:
            candidate = result['candidates'][0]
            if 'content' in candidate and 'parts' in candidate['content'] and candidate['content']['parts']:
                return candidate['content']['parts'][0].get('text', 'No text in response')
        return "Unexpected response format from Gemini API"
    except requests.exceptions.Timeout:
        return "Gemini API request timed out"
    except requests.exceptions.HTTPError as e:
        error_detail = ""
        try:
            error_response = e.response.json()
            error_detail = f" - {error_response.get('error', {}).get('message', '')}"
        except:
            pass
        return f"Gemini API HTTP Error: {e.response.status_code}{error_detail}"
    except Exception as e:
        return f"Gemini API Error: {str(e)}"

@compare_service.route('/compare', methods=['POST'])
def compare_documents():
    try:
        file1 = request.files.get('file1')
        file2 = request.files.get('file2')
        if not file1 or not file2:
            return jsonify({'error': 'Both documents must be provided!'}), 400
        text1, _ = read_text_and_tables(file1)
        text2, _ = read_text_and_tables(file2)
        if not text1 or not text2:
            return jsonify({'error': 'Could not extract text from one or both documents'}), 400
        clauses1 = split_into_clauses(text1)
        clauses2 = split_into_clauses(text2)
        sem_clause_diff = semantic_clause_diff(clauses1, clauses2)
        similarity_score = calculate_overall_similarity(sem_clause_diff)
        key_differences = extract_key_differences(sem_clause_diff)
        return jsonify({
            'from': file1.filename,
            'to': file2.filename,
            'semantic_clause_diff': sem_clause_diff,
            'overall_similarity': similarity_score,
            'key_differences': key_differences,
            'summary': f"Compared {len(clauses1)} clauses from {file1.filename} with {len(clauses2)} clauses from {file2.filename}. Overall similarity: {similarity_score:.2%}"
        })
    except Exception as e:
        logger.error(f"Error in document comparison: {e}")
        return jsonify({'error': f'Comparison failed: {str(e)}'}), 500

@compare_service.route('/advanced-compare', methods=['POST'])
def advanced_compare_from_db_or_file():
    try:
        logger.info("Advanced compare endpoint called")

        # Safely parse document IDs from form data
        raw_doc1_id = request.form.get('doc1_id')
        raw_doc2_id = request.form.get('doc2_id')

        # Convert to int if valid, otherwise None
        def safe_int(val):
            if val is None or val == '':
                return None
            try:
                return int(val)
            except (ValueError, TypeError):
                logger.warning(f"Invalid document ID provided: {val}")
                return None

        doc1_id = safe_int(raw_doc1_id)
        doc2_id = safe_int(raw_doc2_id)

        # Fetch text from DB if valid IDs are provided
        text1, doc1_name = get_extracted_text_from_db(doc1_id) if doc1_id is not None else (None, "Document 1")
        text2, doc2_name = get_extracted_text_from_db(doc2_id) if doc2_id is not None else (None, "Document 2")

        # Fallback to file uploads if text not retrieved and files are provided
        if not text1 and 'file1' in request.files:
            file1 = request.files['file1']
            text1, _ = read_text_and_tables(file1)
            doc1_name = file1.filename

        if not text2 and 'file2' in request.files:
            file2 = request.files['file2']
            text2, _ = read_text_and_tables(file2)
            doc2_name = file2.filename

        if not text1 or not text2:
            return jsonify({'error': 'Could not retrieve text from both documents'}), 400

        # Perform comparison
        clauses1 = split_into_clauses(text1)
        clauses2 = split_into_clauses(text2)
        sem_clause_diff = semantic_clause_diff(clauses1, clauses2)
        similarity_score = calculate_overall_similarity(sem_clause_diff)
        key_differences = extract_key_differences(sem_clause_diff)
        gemini_summary = gemini_advanced_compare(text1, text2, doc1_name, doc2_name)

        # Save comparison result (doc1_id/doc2_id may be None — that's OK per schema)
        comparison_id = save_document_comparison(
            doc1_id=doc1_id,
            doc2_id=doc2_id,
            doc1_name=doc1_name,
            doc2_name=doc2_name,
            comparison_type='advanced-gemini',
            summary=f"Similarity: {similarity_score:.2%}",
            semantic_diff=sem_clause_diff,
            gemini_summary=gemini_summary,
            similarity_score=float(similarity_score) if similarity_score is not None else None,
            key_differences=key_differences
        )

        return jsonify({
            'comparison_id': comparison_id,
            'doc1_id': doc1_id,
            'doc2_id': doc2_id,
            'doc1_name': doc1_name,
            'doc2_name': doc2_name,
            'semantic_clause_diff': sem_clause_diff,
            'overall_similarity': similarity_score,
            'key_differences': key_differences,
            'gemini_summary': gemini_summary,
            'clauses_compared': {'document1': len(clauses1), 'document2': len(clauses2)}
        })

    except Exception as e:
        logger.error(f"Error in advanced comparison: {e}", exc_info=True)
        return jsonify({'error': f'Advanced comparison failed: {str(e)}'}), 500

@compare_service.route('/comparisons', methods=['GET'])
def get_comparisons():
    try:
        page = int(request.args.get('page', 1))
        per_page = int(request.args.get('per_page', 10))
        comparison_type = request.args.get('type', 'all')
        conn = get_db_connection()
        cursor = conn.cursor()
        where_clause = "WHERE comparison_type = ?" if comparison_type != 'all' else ""
        params = [comparison_type] if comparison_type != 'all' else []
        cursor.execute(f"SELECT COUNT(*) FROM document_comparisons {where_clause}", params)
        total_count = cursor.fetchone()[0]
        offset = (page - 1) * per_page
        query = f"""
            SELECT id, doc1_id, doc2_id, doc1_name, doc2_name, compared_at, comparison_type, similarity_score, status
            FROM document_comparisons {where_clause}
            ORDER BY compared_at DESC
            OFFSET ? ROWS FETCH NEXT ? ROWS ONLY
        """
        cursor.execute(query, params + [offset, per_page])
        comparisons = []
        for row in cursor.fetchall():
            comparisons.append({
                'id': row[0],
                'doc1_id': row[1],
                'doc2_id': row[2],
                'doc1_name': row[3],
                'doc2_name': row[4],
                'compared_at': row[5].isoformat() if row[5] else None,
                'comparison_type': row[6],
                'similarity_score': row[7],
                'status': row[8]
            })
        conn.close()
        return jsonify({
            'comparisons': comparisons,
            'pagination': {
                'page': page,
                'per_page': per_page,
                'total_count': total_count,
                'total_pages': (total_count + per_page - 1) // per_page
            }
        }), 200
    except Exception as e:
        logger.error(f"Error fetching comparisons: {e}")
        return jsonify({'error': 'Failed to fetch comparisons'}), 500

@compare_service.route('/comparison/<int:comparison_id>', methods=['GET'])
def get_comparison(comparison_id):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, doc1_id, doc2_id, doc1_name, doc2_name, compared_at, comparison_type, summary, semantic_diff, gemini_summary, similarity_score, key_differences, status
            FROM document_comparisons WHERE id = ?
        """, (comparison_id,))
        row = cursor.fetchone()
        if not row:
            return jsonify({'error': f'Comparison with ID {comparison_id} not found'}), 404
        comparison = {
            'id': row[0],
            'doc1_id': row[1],
            'doc2_id': row[2],
            'doc1_name': row[3],
            'doc2_name': row[4],
            'compared_at': row[5].isoformat() if row[5] else None,
            'comparison_type': row[6],
            'summary': row[7],
            'semantic_diff': json.loads(row[8]) if row[8] else None,
            'gemini_summary': row[9],
            'similarity_score': row[10],
            'key_differences': row[11].split('\n') if row[11] else [],
            'status': row[12]
        }
        conn.close()
        return jsonify(comparison), 200
    except Exception as e:
        logger.error(f"Error fetching comparison {comparison_id}: {e}")
        return jsonify({'error': 'Failed to fetch comparison details'}), 500

@compare_service.route('/comparison/<int:comparison_id>', methods=['DELETE'])
def delete_comparison(comparison_id):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM document_comparisons WHERE id = ?", (comparison_id,))
        if not cursor.fetchone():
            return jsonify({'error': f'Comparison with ID {comparison_id} not found'}), 404
        cursor.execute("DELETE FROM document_comparisons WHERE id = ?", (comparison_id,))
        conn.commit()
        conn.close()
        return jsonify({
            'message': f'Comparison with ID {comparison_id} deleted successfully',
            'deleted_id': comparison_id
        }), 200
    except Exception as e:
        logger.error(f"Error deleting comparison: {e}")
        return jsonify({'error': 'Failed to delete comparison'}), 500

def initialize_compare_service():
    create_document_comparisons_table()