# case_analysis_service.py
import os
import json
import logging
import requests
import numpy as np
from datetime import datetime, timedelta
from flask import Blueprint, request, jsonify
import pyodbc
from sentence_transformers import SentenceTransformer
import uuid
import re

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

case_analysis_service = Blueprint('case_analysis_service', __name__)

# Load models once
try:
    sentence_model = SentenceTransformer('all-MiniLM-L6-v2')
    legal_model = SentenceTransformer('sentence-transformers/all-mpnet-base-v2')
    logger.info("SentenceTransformer models loaded successfully")
except Exception as e:
    logger.error(f"Failed to load SentenceTransformer models: {e}")
    sentence_model = None
    legal_model = None

def create_case_analysis_tables():
    conn = pyodbc.connect(os.getenv("SQLSERVER_CONN_STRING"))
    cursor = conn.cursor()
    cursor.execute("""
        IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='case_analyses')
        CREATE TABLE case_analyses(
            id INT IDENTITY(1,1) PRIMARY KEY,
            analysis_id NVARCHAR(100) UNIQUE NOT NULL,
            user_id NVARCHAR(100),
            case_title NVARCHAR(255),
            analysis_type NVARCHAR(50) NOT NULL,
            document_ids NVARCHAR(MAX),
            case_context NVARCHAR(MAX),
            analysis_parameters NVARCHAR(MAX),
            status NVARCHAR(50) DEFAULT 'processing',
            created_at DATETIME DEFAULT GETDATE(),
            completed_at DATETIME,
            error_message NVARCHAR(MAX)
        )
    """)
    cursor.execute("""
        IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='case_analysis_results')
        CREATE TABLE case_analysis_results(
            id INT IDENTITY(1,1) PRIMARY KEY,
            analysis_id NVARCHAR(100) NOT NULL,
            result_type NVARCHAR(50) NOT NULL,
            content NVARCHAR(MAX),
            confidence_score FLOAT,
            metadata NVARCHAR(MAX),
            created_at DATETIME DEFAULT GETDATE(),
            FOREIGN KEY(analysis_id) REFERENCES case_analyses(analysis_id)
        )
    """)
    cursor.execute("""
        IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='legal_precedents')
        CREATE TABLE legal_precedents(
            id INT IDENTITY(1,1) PRIMARY KEY,
            case_name NVARCHAR(255),
            citation NVARCHAR(255),
            court NVARCHAR(255),
            jurisdiction NVARCHAR(100),
            date_decided DATE,
            legal_issues NVARCHAR(MAX),
            key_holdings NVARCHAR(MAX),
            relevance_keywords NVARCHAR(MAX),
            summary NVARCHAR(MAX),
            full_text NVARCHAR(MAX),
            created_at DATETIME DEFAULT GETDATE()
        )
    """)
    cursor.execute("""
        IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='risk_factors')
        CREATE TABLE risk_factors(
            id INT IDENTITY(1,1) PRIMARY KEY,
            factor_type NVARCHAR(100),
            factor_name NVARCHAR(255),
            description NVARCHAR(MAX),
            severity_level NVARCHAR(50),
            jurisdiction NVARCHAR(100),
            practice_area NVARCHAR(100),
            mitigation_strategies NVARCHAR(MAX),
            created_at DATETIME DEFAULT GETDATE()
        )
    """)
    conn.commit()
    conn.close()
    logger.info("Case analysis tables created successfully")

def initialize_sample_legal_data():
    """Initialize sample legal precedents and risk factors."""
    try:
        conn = pyodbc.connect(os.getenv("SQLSERVER_CONN_STRING"))
        cursor = conn.cursor()

        # Sample legal precedents
        sample_precedents = [
            {
                'case_name': 'Smith v. Jones Corp',
                'citation': '123 F.3d 456 (9th Cir. 2020)',
                'court': 'US Court of Appeals, 9th Circuit',
                'jurisdiction': 'federal',
                'date_decided': '2020-03-15',
                'legal_issues': 'Contract interpretation, breach of contract, damages',
                'key_holdings': 'Material breach requires substantial performance failure. Consequential damages must be foreseeable.',
                'relevance_keywords': 'contract, breach, damages, foreseeability, material breach',
                'summary': 'Court held that minor deviations from contract terms do not constitute material breach, and consequential damages must be reasonably foreseeable at time of contract formation.'
            },
            {
                'case_name': 'ABC LLC v. State Regulatory Board',
                'citation': '789 F.Supp.2d 123 (S.D.N.Y. 2019)',
                'court': 'US District Court, Southern District of New York',
                'jurisdiction': 'federal',
                'date_decided': '2019-11-08',
                'legal_issues': 'Administrative law, due process, regulatory compliance',
                'key_holdings': 'Administrative agencies must provide adequate notice and opportunity to be heard before imposing sanctions.',
                'relevance_keywords': 'administrative law, due process, regulatory, sanctions, notice',
                'summary': 'Federal court invalidated regulatory sanctions imposed without proper procedural safeguards, emphasizing importance of due process in administrative proceedings.'
            },
            {
                'case_name': 'Johnson v. Medical Center',
                'citation': '456 Cal.App.4th 789 (2021)',
                'court': 'California Court of Appeal',
                'jurisdiction': 'california',
                'date_decided': '2021-07-22',
                'legal_issues': 'Medical malpractice, standard of care, expert testimony',
                'key_holdings': 'Expert testimony must establish both standard of care and deviation therefrom. Res ipsa loquitur doctrine has limited application in complex medical procedures.',
                'relevance_keywords': 'medical malpractice, standard of care, expert testimony, res ipsa loquitur',
                'summary': 'Court affirmed summary judgment for defendant hospital, holding that plaintiff failed to present adequate expert testimony establishing breach of standard of care.'
            }
        ]

        for p in sample_precedents:
            # Check if already exists
            cursor.execute("SELECT COUNT(*) FROM legal_precedents WHERE case_name = ?", (p['case_name'],))
            if cursor.fetchone()[0] == 0:
                cursor.execute("""
                    INSERT INTO legal_precedents (case_name, citation, court, jurisdiction, date_decided, legal_issues, key_holdings, relevance_keywords, summary)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    p['case_name'],
                    p['citation'],
                    p['court'],
                    p['jurisdiction'],
                    p['date_decided'],
                    p['legal_issues'],
                    p['key_holdings'],
                    p['relevance_keywords'],
                    p['summary']
                ))
                logger.info(f"Added precedent: {p['case_name']}")

        # Sample risk factors
        sample_risks = [
            {
                'factor_type': 'procedural',
                'factor_name': 'Statute of Limitations',
                'description': 'Risk that claims may be time-barred by applicable statutes of limitations',
                'severity_level': 'high',
                'jurisdiction': 'general',
                'practice_area': 'litigation',
                'mitigation_strategies': 'Conduct thorough limitations analysis, consider discovery rule exceptions, file protective pleadings if necessary'
            },
            {
                'factor_type': 'evidentiary',
                'factor_name': 'Document Preservation',
                'description': 'Risk of spoliation sanctions due to inadequate document retention',
                'severity_level': 'high',
                'jurisdiction': 'general',
                'practice_area': 'litigation',
                'mitigation_strategies': 'Issue litigation hold notices, implement document retention policies, work with IT to preserve electronic evidence'
            },
            {
                'factor_type': 'substantive',
                'factor_name': 'Causation Proof',
                'description': 'Difficulty establishing causation between defendant\'s actions and plaintiff\'s damages',
                'severity_level': 'medium',
                'jurisdiction': 'general',
                'practice_area': 'tort',
                'mitigation_strategies': 'Retain expert witnesses early, conduct thorough factual investigation, consider alternative causation theories'
            },
            {
                'factor_type': 'financial',
                'factor_name': 'Litigation Costs',
                'description': 'High costs of litigation may exceed potential recovery',
                'severity_level': 'medium',
                'jurisdiction': 'general',
                'practice_area': 'general',
                'mitigation_strategies': 'Conduct cost-benefit analysis, explore alternative fee arrangements, consider litigation funding options'
            }
        ]

        for r in sample_risks:
            cursor.execute("""
                SELECT COUNT(*) FROM risk_factors 
                WHERE factor_name = ? AND practice_area = ?
            """, (r['factor_name'], r['practice_area']))
            if cursor.fetchone()[0] == 0:
                cursor.execute("""
                    INSERT INTO risk_factors (factor_type, factor_name, description, severity_level, jurisdiction, practice_area, mitigation_strategies)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (
                    r['factor_type'],
                    r['factor_name'],
                    r['description'],
                    r['severity_level'],
                    r['jurisdiction'],
                    r['practice_area'],
                    r['mitigation_strategies']
                ))
                logger.info(f"Added risk factor: {r['factor_name']}")

        conn.commit()
        logger.info("Sample legal data initialized successfully")

    except Exception as e:
        logger.error(f"Error initializing sample data: {e}")
        raise e
    finally:
        if 'cursor' in locals():
            cursor.close()
        if 'conn' in locals():
            conn.close()

def extract_legal_issues(document_texts):
    legal_patterns = [
        r'\b(breach of (?:contract|duty|warranty))\b',
        r'\b(negligence|malpractice)\b',
        r'\b(fraud(?:ulent)?(?:\s+(?:misrepresentation|inducement))?)\b',
        r'\b(discrimination|harassment)\b',
        r'\b(defamation|libel|slander)\b',
        r'\b(intellectual property|copyright|trademark|patent)\b',
        r'\b(employment|wrongful termination)\b',
        r'\b(personal injury|wrongful death)\b',
        r'\b(constitutional|civil rights)\b',
        r'\b(administrative|regulatory)\b',
        r'\b(corporate|securities)\b',
        r'\b(real estate|property)\b'
    ]
    legal_issues = set()
    combined_text = ''.join(document_texts).lower()
    for pattern in legal_patterns:
        matches = re.findall(pattern, combined_text, re.IGNORECASE)
        for match in matches:
            legal_issues.add(match[0] if isinstance(match, tuple) else match)
    return list(legal_issues)

def find_relevant_precedents(legal_issues, jurisdiction='general', limit=5):
    conn = pyodbc.connect(os.getenv("SQLSERVER_CONN_STRING"))
    cursor = conn.cursor()
    search_terms = ' '.join(legal_issues).lower()
    where_conditions = []
    params = [limit, jurisdiction]
    for issue in legal_issues[:5]:
        pattern = f"%{issue}%"
        where_conditions.extend(["legal_issues LIKE ?", "key_holdings LIKE ?", "relevance_keywords LIKE ?", "summary LIKE ?"])
        params.extend([pattern] * 4)
    query = f"""
        SELECT TOP(?) case_name, citation, court, jurisdiction, date_decided, legal_issues, key_holdings, summary,
               (CASE WHEN jurisdiction = ? THEN 3 WHEN jurisdiction = 'federal' THEN 2 WHEN jurisdiction = 'general' THEN 1 ELSE 0 END) as jurisdiction_score
        FROM legal_precedents
        WHERE ({' OR '.join(where_conditions)})
        ORDER BY jurisdiction_score DESC, date_decided DESC
    """
    cursor.execute(query, params)
    precedents = []
    for row in cursor.fetchall():
        precedents.append({
            'case_name': row[0],
            'citation': row[1],
            'court': row[2],
            'jurisdiction': row[3],
            'date_decided': row[4].strftime('%Y-%m-%d') if row[4] else None,
            'legal_issues': row[5],
            'key_holdings': row[6],
            'summary': row[7],
            'jurisdiction_score': row[8]
        })
    conn.close()
    return precedents

def assess_case_risks(legal_issues, case_context, practice_area='litigation'):
    conn = pyodbc.connect(os.getenv("SQLSERVER_CONN_STRING"))
    cursor = conn.cursor()
    cursor.execute("""
        SELECT factor_type, factor_name, description, severity_level, mitigation_strategies
        FROM risk_factors
        WHERE (practice_area = ? OR practice_area = 'general')
    """, (practice_area,))
    risk_factors = []
    for row in cursor.fetchall():
        risk = {
            'factor_type': row[0],
            'factor_name': row[1],
            'description': row[2],
            'severity_level': row[3],
            'mitigation_strategies': row[4],
            'relevance_score': calculate_risk_relevance(row[1], legal_issues, case_context)
        }
        risk_factors.append(risk)
    conn.close()
    risk_factors.sort(key=lambda x: x['relevance_score'], reverse=True)
    return risk_factors[:10]

def calculate_risk_relevance(risk_name, legal_issues, case_context):
    score = 0.0
    for issue in legal_issues:
        if issue.lower() in risk_name.lower():
            score += 0.5
    if case_context:
        context_text = ''.join(str(v) for v in case_context.values()).lower()
        if risk_name.lower() in context_text:
            score += 0.3
        if 'statute' in risk_name.lower() and ('time' in context_text or 'limitation' in context_text):
            score += 0.4
        if 'evidence' in risk_name.lower() and ('document' in context_text or 'proof' in context_text):
            score += 0.4
        if 'causation' in risk_name.lower() and ('cause' in context_text or 'damage' in context_text):
            score += 0.4
    return min(score, 1.0)

def generate_with_gemini(prompt):
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_API_KEY}"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.2, "topK": 40, "topP": 0.8, "maxOutputTokens": 4000}
    }
    response = requests.post(url, json=payload, timeout=90)
    response.raise_for_status()
    result = response.json()
    if 'candidates' in result and result['candidates']:
        return result['candidates'][0]['content']['parts'][0]['text']
    raise Exception("Invalid Gemini API response")

def generate_template_recommendations(analysis_type, legal_issues, precedents, risks, case_context):
    strength = get_case_strength_assessment(legal_issues, precedents, risks)
    recommendations = f"""STRATEGIC CASE ANALYSIS - {analysis_type.upper()}
EXECUTIVE SUMMARY: Based on the analysis of legal issues, relevant precedents, and identified risks, this case presents a {strength} prospect for success.
KEY LEGAL ISSUES:
{chr(10).join(f"• {issue.title()}" for issue in legal_issues)}
STRATEGIC RECOMMENDATIONS:
1. IMMEDIATE ACTIONS (Next 30 Days):
• Conduct comprehensive case investigation
• Preserve all relevant documents and evidence
• Identify and retain necessary expert witnesses
• Review applicable statutes of limitations
2. CASE DEVELOPMENT STRATEGY:
• Focus on strongest legal theories: {', '.join(legal_issues[:3])}
• Develop factual foundation for key claims
• Anticipate and prepare for likely defenses
3. RISK MITIGATION PRIORITIES:"""
    for risk in risks[:3]:
        recommendations += f"  • {risk['factor_name']}: {risk['mitigation_strategies'][:100]}...\n"
    recommendations += f""" 4. PRECEDENT ANALYSIS: Based on {len(precedents)} relevant precedents found, key considerations include: """
    for p in precedents[:2]:
        recommendations += f"  • {p['case_name']}: {p['key_holdings'][:150]}...\n"
    recommendations += """ 5. RECOMMENDED TIMELINE:
• Discovery Phase: 6-9 months
• Motion Practice: 3-4 months
• Trial Preparation: 2-3 months
• Alternative Resolution: Consider at 6-month mark
6. BUDGET CONSIDERATIONS:
• Retain qualified experts early to control costs
• Prioritize high-impact discovery activities
• Consider litigation funding if appropriate
This analysis should be updated as new information becomes available and case circumstances evolve.
"""
    return recommendations

def get_case_strength_assessment(legal_issues, precedents, risks):
    score = 0
    if len(legal_issues) >= 3:
        score += 1
    elif len(legal_issues) >= 1:
        score += 0.5
    if len(precedents) >= 3:
        score += 1
    elif len(precedents) >= 1:
        score += 0.5
    high_risk_count = sum(1 for r in risks if r['severity_level'] == 'high')
    if high_risk_count >= 3:
        score -= 1
    elif high_risk_count >= 1:
        score -= 0.5
    if score >= 1.5:
        return "strong"
    elif score >= 0.5:
        return "moderate"
    else:
        return "challenging"

def generate_strategic_recommendations(analysis_type, legal_issues, precedents, risks, case_context):
    if not os.getenv("GEMINI_API_KEY"):
        return generate_template_recommendations(analysis_type, legal_issues, precedents, risks, case_context)
    prompt = f"""As a senior legal strategist, analyze this case and provide comprehensive strategic recommendations.
ANALYSIS TYPE: {analysis_type}
LEGAL ISSUES IDENTIFIED:
{chr(10).join(f"• {issue}" for issue in legal_issues)}
RELEVANT PRECEDENTS:
{chr(10).join(f"• {p['case_name']} - {p['key_holdings'][:200]}..." for p in precedents[:3])}
KEY RISKS IDENTIFIED:
{chr(10).join(f"• {r['factor_name']} ({r['severity_level']} severity): {r['description']}" for r in risks[:5])}
CASE CONTEXT: {json.dumps(case_context, indent=2)}
Please provide:
1. STRATEGIC OVERVIEW - Overall case strength assessment - Primary legal theories to pursue - Potential challenges and opportunities
2. RECOMMENDED ACTIONS - Immediate next steps (30 days) - Medium-term strategy (3-6 months) - Long-term considerations
3. RISK MITIGATION - Priority risks to address - Specific mitigation strategies - Timeline for risk mitigation activities
4. RESOURCE ALLOCATION - Expert witnesses needed - Discovery priorities - Budget considerations
5. SETTLEMENT CONSIDERATIONS - Likelihood of favorable settlement - Optimal timing for settlement discussions - Key leverage points
6. TIMELINE AND MILESTONES - Critical deadlines - Key decision points - Anticipated case duration
Provide practical, actionable recommendations based on the legal issues, precedents, and risks identified.
Limit response to 2000 words for clarity and actionability."""
    try:
        return generate_with_gemini(prompt)
    except Exception as e:
        logger.error(f"AI recommendation generation failed: {e}")
        return generate_template_recommendations(analysis_type, legal_issues, precedents, risks, case_context)

@case_analysis_service.route('/analyze-case', methods=['POST'])
def analyze_case():
    try:
        data = request.get_json()
        document_ids = data.get('document_ids', [])
        analysis_type = data.get('analysis_type', 'case-strategy')
        case_context = data.get('case_context', {})
        user_id = data.get('user_id', 'system')
        if not document_ids:
            return jsonify({'error': 'document_ids are required'}), 400

        analysis_id = str(uuid.uuid4())
        conn = pyodbc.connect(os.getenv("SQLSERVER_CONN_STRING"))
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO case_analyses (analysis_id, user_id, analysis_type, document_ids, case_context, analysis_parameters, status)
            VALUES (?, ?, ?, ?, ?, ?, 'processing')
        """, (analysis_id, user_id, analysis_type, json.dumps(document_ids), json.dumps(case_context), json.dumps(data)))
        conn.commit()

        document_texts = []
        for doc_id in document_ids:
            cursor.execute("SELECT extracted_text, filename FROM cases WHERE id = ?", (doc_id,))
            row = cursor.fetchone()
            if row and row[0]:
                document_texts.append(row[0])

        if not document_texts:
            cursor.execute("UPDATE case_analyses SET status = 'failed', error_message = 'No document texts found' WHERE analysis_id = ?", (analysis_id,))
            conn.commit()
            return jsonify({'error': 'No valid document texts found'}), 400

        legal_issues = extract_legal_issues(document_texts)
        jurisdiction = case_context.get('jurisdiction', 'general')
        practice_area = case_context.get('caseType', 'litigation')
        precedents = find_relevant_precedents(legal_issues, jurisdiction)
        risks = assess_case_risks(legal_issues, case_context, practice_area)
        recommendations = generate_strategic_recommendations(analysis_type, legal_issues, precedents, risks, case_context)

        results = [
            {'result_type': 'legal_issues', 'content': json.dumps(legal_issues), 'confidence_score': 0.8},
            {'result_type': 'precedents', 'content': json.dumps(precedents), 'confidence_score': 0.9},
            {'result_type': 'risks', 'content': json.dumps(risks), 'confidence_score': 0.85},
            {'result_type': 'recommendations', 'content': recommendations, 'confidence_score': 0.9}
        ]
        for r in results:
            cursor.execute("""
                INSERT INTO case_analysis_results (analysis_id, result_type, content, confidence_score)
                VALUES (?, ?, ?, ?)
            """, (analysis_id, r['result_type'], r['content'], r['confidence_score']))

        cursor.execute("UPDATE case_analyses SET status = 'completed', completed_at = GETDATE() WHERE analysis_id = ?", (analysis_id,))
        conn.commit()
        conn.close()

        return jsonify({
            'analysis_id': analysis_id,
            'analysis_type': analysis_type,
            'status': 'completed',
            'legal_issues': legal_issues,
            'precedents': precedents,
            'risks': risks,
            'recommendations': recommendations,
            'case_context': case_context,
            'documents_analyzed': len(document_texts),
            'completed_at': datetime.now().isoformat()
        }), 200
    except Exception as e:
        logger.error(f"Error in case analysis: {e}")
        return jsonify({'error': f'Case analysis failed: {str(e)}'}), 500

@case_analysis_service.route('/research-precedents', methods=['POST'])
def research_precedents():
    try:
        data = request.get_json()
        legal_issues = data.get('legal_issues', [])
        jurisdiction = data.get('jurisdiction', 'general')
        case_type = data.get('case_type', 'litigation')
        limit = int(data.get('limit', 10))
        if not legal_issues:
            return jsonify({'error': 'legal_issues are required'}), 400
        precedents = find_relevant_precedents(legal_issues, jurisdiction, limit)
        return jsonify({
            'precedents': precedents,
            'search_criteria': {'legal_issues': legal_issues, 'jurisdiction': jurisdiction, 'case_type': case_type},
            'total_found': len(precedents)
        }), 200
    except Exception as e:
        logger.error(f"Error researching precedents: {e}")
        return jsonify({'error': f'Precedent research failed: {str(e)}'}), 500

@case_analysis_service.route('/assess-risk', methods=['POST'])
def assess_risk():
    try:
        data = request.get_json()
        legal_issues = data.get('legal_issues', [])
        case_context = data.get('case_context', {})
        practice_area = data.get('practice_area', 'litigation')
        if not legal_issues:
            return jsonify({'error': 'legal_issues are required'}), 400
        risks = assess_case_risks(legal_issues, case_context, practice_area)
        high_risks = sum(1 for r in risks if r['severity_level'] == 'high')
        medium_risks = sum(1 for r in risks if r['severity_level'] == 'medium')
        overall_risk = 'high' if high_risks >= 2 else 'medium' if medium_risks >= 3 else 'low'
        return jsonify({
            'risks': risks,
            'risk_summary': {
                'overall_risk_level': overall_risk,
                'high_severity_count': high_risks,
                'medium_severity_count': medium_risks,
                'total_risks_identified': len(risks)
            },
            'assessment_criteria': {'legal_issues': legal_issues, 'practice_area': practice_area}
        }), 200
    except Exception as e:
        logger.error(f"Error assessing risks: {e}")
        return jsonify({'error': f'Risk assessment failed: {str(e)}'}), 500

@case_analysis_service.route('/analyses', methods=['GET'])
def get_case_analyses():
    try:
        page = int(request.args.get('page', 1))
        per_page = int(request.args.get('per_page', 10))
        user_id = request.args.get('user_id')
        status = request.args.get('status', 'all')
        analysis_type = request.args.get('type', 'all')
        conn = pyodbc.connect(os.getenv("SQLSERVER_CONN_STRING"))
        cursor = conn.cursor()
        where_clauses = []
        params = []
        if user_id:
            where_clauses.append("user_id = ?")
            params.append(user_id)
        if status != 'all':
            where_clauses.append("status = ?")
            params.append(status)
        if analysis_type != 'all':
            where_clauses.append("analysis_type = ?")
            params.append(analysis_type)
        where_clause = "WHERE " + " AND ".join(where_clauses) if where_clauses else ""
        cursor.execute(f"SELECT COUNT(*) FROM case_analyses {where_clause}", params)
        total_count = cursor.fetchone()[0]
        offset = (page - 1) * per_page
        query = f"""
            SELECT analysis_id, case_title, analysis_type, status, created_at, completed_at
            FROM case_analyses {where_clause}
            ORDER BY created_at DESC
            OFFSET ? ROWS FETCH NEXT ? ROWS ONLY
        """
        cursor.execute(query, params + [offset, per_page])
        analyses = []
        for row in cursor.fetchall():
            analyses.append({
                'analysis_id': row[0],
                'case_title': row[1],
                'analysis_type': row[2],
                'status': row[3],
                'created_at': row[4].isoformat() if row[4] else None,
                'completed_at': row[5].isoformat() if row[5] else None
            })
        conn.close()
        return jsonify({
            'analyses': analyses,
            'pagination': {
                'page': page,
                'per_page': per_page,
                'total_count': total_count,
                'total_pages': (total_count + per_page - 1) // per_page
            }
        }), 200
    except Exception as e:
        logger.error(f"Error fetching case analyses: {e}")
        return jsonify({'error': f'Failed to fetch analyses: {str(e)}'}), 500

@case_analysis_service.route('/analysis/<analysis_id>', methods=['GET'])
def get_case_analysis(analysis_id):
    try:
        conn = pyodbc.connect(os.getenv("SQLSERVER_CONN_STRING"))
        cursor = conn.cursor()
        cursor.execute("""
            SELECT analysis_id, case_title, analysis_type, document_ids, case_context, status, created_at, completed_at, error_message
            FROM case_analyses WHERE analysis_id = ?
        """, (analysis_id,))
        analysis_row = cursor.fetchone()
        if not analysis_row:
            return jsonify({'error': f'Analysis {analysis_id} not found'}), 404
        cursor.execute("""
            SELECT result_type, content, confidence_score, metadata
            FROM case_analysis_results WHERE analysis_id = ?
            ORDER BY created_at
        """, (analysis_id,))
        results = {}
        for row in cursor.fetchall():
            result_type = row[0]
            content = row[1]
            if result_type in ['legal_issues', 'precedents', 'risks']:
                try:
                    content = json.loads(content)
                except:
                    pass
            results[result_type] = {
                'content': content,
                'confidence_score': row[2],
                'metadata': json.loads(row[3]) if row[3] else None
            }
        conn.close()
        analysis = {
            'analysis_id': analysis_row[0],
            'case_title': analysis_row[1],
            'analysis_type': analysis_row[2],
            'document_ids': json.loads(analysis_row[3]) if analysis_row[3] else [],
            'case_context': json.loads(analysis_row[4]) if analysis_row[4] else {},
            'status': analysis_row[5],
            'created_at': analysis_row[6].isoformat() if analysis_row[6] else None,
            'completed_at': analysis_row[7].isoformat() if analysis_row[7] else None,
            'error_message': analysis_row[8],
            'results': results
        }
        return jsonify(analysis), 200
    except Exception as e:
        logger.error(f"Error fetching analysis {analysis_id}: {e}")
        return jsonify({'error': f'Failed to fetch analysis: {str(e)}'}), 500

@case_analysis_service.route('/strategy-recommendations', methods=['POST'])
def get_strategy_recommendations():
    try:
        data = request.get_json()
        legal_issues = data.get('legal_issues', [])
        case_context = data.get('case_context', {})
        analysis_type = data.get('analysis_type', 'case-strategy')
        if not legal_issues:
            return jsonify({'error': 'legal_issues are required'}), 400
        jurisdiction = case_context.get('jurisdiction', 'general')
        practice_area = case_context.get('caseType', 'litigation')
        precedents = find_relevant_precedents(legal_issues, jurisdiction, 5)
        risks = assess_case_risks(legal_issues, case_context, practice_area)
        recommendations = generate_strategic_recommendations(analysis_type, legal_issues, precedents, risks, case_context)
        return jsonify({
            'recommendations': recommendations,
            'supporting_data': {
                'legal_issues': legal_issues,
                'precedents_found': len(precedents),
                'risks_identified': len(risks),
                'analysis_type': analysis_type
            },
            'generated_at': datetime.now().isoformat()
        }), 200
    except Exception as e:
        logger.error(f"Error generating strategy recommendations: {e}")
        return jsonify({'error': f'Strategy generation failed: {str(e)}'}), 500

def create_case_timeline(case_type, jurisdiction, case_context):
    base_timeline = [
        {
            'phase': 'Case Initiation',
            'duration_days': 30,
            'description': 'File complaint, serve defendants, initial case setup',
            'key_activities': ['Draft and file complaint', 'Serve process on defendants', 'File proof of service', 'Issue litigation hold notices'],
            'deadlines': ['Statute of limitations', 'Service deadlines'],
            'priority': 'critical'
        },
        {
            'phase': 'Early Case Management',
            'duration_days': 60,
            'description': 'Initial responses, case management conference, discovery planning',
            'key_activities': ['Defendants file answers/motions', 'Case management conference', 'Discovery planning meeting', 'Preliminary injunction motions (if applicable)'],
            'deadlines': ['Response deadlines', 'CMC scheduling'],
            'priority': 'high'
        },
        {
            'phase': 'Discovery Phase',
            'duration_days': 180,
            'description': 'Fact discovery, document production, depositions',
            'key_activities': ['Written discovery (interrogatories, RFPs)', 'Document production and review', 'Fact witness depositions', 'Expert witness designation'],
            'deadlines': ['Discovery cutoff', 'Expert designations'],
            'priority': 'high'
        },
        {
            'phase': 'Expert Discovery',
            'duration_days': 90,
            'description': 'Expert witness discovery and depositions',
            'key_activities': ['Expert witness reports', 'Expert depositions', 'Rebuttal expert reports', 'Expert discovery disputes'],
            'deadlines': ['Expert report deadlines', 'Expert deposition cutoff'],
            'priority': 'medium'
        },
        {
            'phase': 'Motion Practice',
            'duration_days': 120,
            'description': 'Summary judgment and other dispositive motions',
            'key_activities': ['Summary judgment motions', 'Daubert motions (expert challenges)', 'Other dispositive motions', 'Motion hearings'],
            'deadlines': ['Motion filing deadlines', 'Hearing dates'],
            'priority': 'high'
        },
        {
            'phase': 'Trial Preparation',
            'duration_days': 90,
            'description': 'Final trial preparation and pre-trial motions',
            'key_activities': ['Witness and exhibit lists', 'Pre-trial motions in limine', 'Settlement conferences', 'Trial preparation'],
            'deadlines': ['Pre-trial deadlines', 'Settlement conference'],
            'priority': 'critical'
        },
        {
            'phase': 'Trial',
            'duration_days': 14,
            'description': 'Trial proceedings',
            'key_activities': ['Jury selection', 'Opening statements', 'Witness testimony', 'Closing arguments and verdict'],
            'deadlines': ['Trial date'],
            'priority': 'critical'
        }
    ]
    if case_type.lower() in ['employment', 'employment law']:
        base_timeline.insert(0, {
            'phase': 'Administrative Prerequisites',
            'duration_days': 90,
            'description': 'EEOC charge filing and processing',
            'key_activities': ['File EEOC charge', 'EEOC investigation', 'Obtain right-to-sue letter', 'Evaluate settlement opportunities'],
            'deadlines': ['EEOC filing deadline', 'Right-to-sue expiration'],
            'priority': 'critical'
        })
    elif case_type.lower() in ['personal injury', 'medical malpractice']:
        for phase in base_timeline:
            if phase['phase'] == 'Discovery Phase':
                phase['key_activities'].extend(['Medical record collection', 'Medical expert evaluations', 'Accident reconstruction analysis'])
    current_date = datetime.now()
    for phase in base_timeline:
        phase['start_date'] = current_date.strftime('%Y-%m-%d')
        phase['end_date'] = (current_date + timedelta(days=phase['duration_days'])).strftime('%Y-%m-%d')
        current_date += timedelta(days=phase['duration_days'])
    return base_timeline

@case_analysis_service.route('/case-timeline', methods=['POST'])
def generate_case_timeline():
    try:
        data = request.get_json()
        case_type = data.get('case_type', 'litigation')
        jurisdiction = data.get('jurisdiction', 'federal')
        case_context = data.get('case_context', {})
        timeline = create_case_timeline(case_type, jurisdiction, case_context)
        return jsonify({
            'timeline': timeline,
            'case_type': case_type,
            'jurisdiction': jurisdiction,
            'generated_at': datetime.now().isoformat()
        }), 200
    except Exception as e:
        logger.error(f"Error generating case timeline: {e}")
        return jsonify({'error': f'Timeline generation failed: {str(e)}'}), 500

def assess_case_strength(legal_issues, case_context):
    return {
        'overall_strength': 'moderate',
        'strongest_claims': legal_issues[:3] if legal_issues else [],
        'potential_weaknesses': ['Discovery risks', 'Causation challenges'],
        'probability_of_success': 0.65
    }

def assess_litigation_risks(legal_issues, case_context):
    return {
        'procedural_risks': ['Statute of limitations', 'Jurisdiction challenges'],
        'substantive_risks': ['Proof of damages', 'Causation requirements'],
        'strategic_risks': ['Adverse precedent', 'Counterclaims'],
        'overall_risk_level': 'medium'
    }

def analyze_cost_factors(damages_estimate, litigation_costs):
    return {
        'estimated_damages': damages_estimate,
        'estimated_litigation_costs': litigation_costs,
        'cost_benefit_ratio': 0.75 if damages_estimate and litigation_costs else None,
        'break_even_settlement': damages_estimate * 0.6 if damages_estimate else None
    }

def assess_settlement_timing(case_context):
    return {
        'current_phase': 'discovery',
        'optimal_timing': 'After key depositions',
        'timing_advantages': ['Information leverage', 'Cost control'],
        'timing_risks': ['Evidence development', 'Momentum loss']
    }

def identify_leverage_points(case_context, legal_issues):
    return [
        'Strong document evidence',
        'Adverse publicity risk for defendant',
        'Regulatory exposure',
        'Insurance coverage limits'
    ]

def generate_settlement_strategy(settlement_factors, case_context):
    return """SETTLEMENT STRATEGY ANALYSIS
RECOMMENDED APPROACH: Based on the case analysis, a measured settlement approach is recommended with the following considerations:
1. TIMING STRATEGY:
- Initiate discussions after completing key depositions
- Leverage information asymmetries from discovery
- Consider formal mediation in 4-6 months
2. NEGOTIATION POSITION:
- Initial demand should account for litigation risks
- Emphasize strongest legal theories and evidence
- Prepare for extended negotiation process
3. LEVERAGE POINTS:
- Document strength provides significant leverage
- Defendant's regulatory exposure creates pressure
- Insurance coverage limits may cap exposure
4. RISK MITIGATION:
- Maintain litigation pressure throughout negotiations
- Preserve all legal theories and claims
- Document settlement discussions carefully
5. SUCCESS METRICS:
- Target settlement range: 60-75% of estimated damages
- Timeline: Resolution within 6-9 months
- Cost efficiency: Minimize discovery expenses
This strategy should be adjusted based on evolving case circumstances and opposing party responses."""

@case_analysis_service.route('/settlement-analysis', methods=['POST'])
def settlement_analysis():
    try:
        data = request.get_json()
        case_context = data.get('case_context', {})
        legal_issues = data.get('legal_issues', [])
        damages_estimate = data.get('damages_estimate')
        litigation_costs = data.get('litigation_costs')
        settlement_factors = {
            'strength_of_case': assess_case_strength(legal_issues, case_context),
            'litigation_risks': assess_litigation_risks(legal_issues, case_context),
            'cost_considerations': analyze_cost_factors(damages_estimate, litigation_costs),
            'timing_factors': assess_settlement_timing(case_context),
            'leverage_points': identify_leverage_points(case_context, legal_issues)
        }
        settlement_recommendation = generate_settlement_strategy(settlement_factors, case_context)
        return jsonify({
            'settlement_analysis': settlement_factors,
            'recommendations': settlement_recommendation,
            'analysis_date': datetime.now().isoformat()
        }), 200
    except Exception as e:
        logger.error(f"Error in settlement analysis: {e}")
        return jsonify({'error': f'Settlement analysis failed: {str(e)}'}), 500

def initialize_case_analysis_service():
    create_case_analysis_tables()
    initialize_sample_legal_data()
    logger.info("Case analysis service initialized successfully")