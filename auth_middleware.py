# auth_middleware.py - Enhanced with client verification
import os
import jwt
import pyodbc
from functools import wraps
from flask import request, jsonify
from dotenv import load_dotenv
from datetime import datetime
import traceback

load_dotenv()

JWT_SECRET = os.getenv("JWT_SECRET", "your_actual_32+_character_secret_from_.net")
JWT_ALGORITHM = "HS256"
SQLSERVER_CONN_STRING = os.getenv("SQLSERVER_CONN_STRING")

def get_db_connection():
    """Get database connection"""
    if not SQLSERVER_CONN_STRING:
        raise ValueError("SQLSERVER_CONN_STRING not found")
    return pyodbc.connect(SQLSERVER_CONN_STRING)

def get_user_from_token():
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        print("❌ No Authorization header")
        return None

    token = auth_header.replace('Bearer ', '')
    try:
        payload = jwt.decode(
            token,
            JWT_SECRET,
            algorithms=[JWT_ALGORITHM],
            options={"verify_aud": False}  # ← Keep this!
        )
        print("✅ Decoded JWT payload:", payload)

        # Extract user_id
        user_id = (
            payload.get('http://schemas.xmlsoap.org/ws/2005/05/identity/claims/nameidentifier')
            or payload.get('user_id')
            or payload.get('sub')
        )

        email = payload.get('http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress') or payload.get('email')
        username = payload.get('http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name') or payload.get('name') or email

        # SAFE: Always returns a list
        roles = []
        for key, value in payload.items():
            if isinstance(key, str) and 'role' in key.lower():
                if isinstance(value, list):
                    roles.extend([str(v) for v in value])
                else:
                    roles.append(str(value))
        roles = list(set(roles))

        print("🔍 Extracted roles (type: %s): %s", type(roles), roles)

        user_dict = {
            'user_id': user_id,
            'email': email,
            'username': username,
            'roles': roles,  # ← MUST be list
            'subscription_tier': payload.get('subscriptionTier'),
            'document_quota': int(payload.get('documentQuota', 50)),
            'documents_processed': int(payload.get('documentsProcessed', 0))
        }

        print("✅ Returning user dict:", user_dict)
        return user_dict

    except jwt.ExpiredSignatureError:
        print("❌ JWT expired")
        return None
    except jwt.InvalidTokenError as e:
        print("❌ Invalid JWT:", str(e))
        return None

def get_current_user_quota():
    """Get real-time quota from database"""
    user = get_user_from_token()
    if not user:
        return None, None
    
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT DocumentQuota, DocumentsProcessedThisMonth, SubscriptionTier
            FROM Users
            WHERE Id = ?
        """, (user['user_id'],))
        
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return row[0], row[1]  # quota, processed
        return None, None
    except Exception as e:
        print(f"Error getting user quota: {e}")
        return None, None

def verify_client_ownership(client_id, user_id=None):
    """
    Verify that a client belongs to a specific user.
    If user_id is None, uses the current authenticated user.
    Returns True if the client belongs to the user, False otherwise.
    
    Args:
        client_id (int): The client ID to verify
        user_id (str, optional): The user ID to check against. Defaults to current user.
    
    Returns:
        bool: True if client belongs to user, False otherwise
    """
    if user_id is None:
        user = get_user_from_token()
        if not user:
            return False
        user_id = user['user_id']
    
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT Id FROM Clients WHERE Id = ? AND UserId = ?", (client_id, user_id))
        result = cursor.fetchone()
        conn.close()
        return result is not None
    except Exception as e:
        print(f"Error verifying client ownership: {e}")
        return False

def require_client_access(f):
    """
    Decorator to require client access verification.
    Expects 'client_id' in request args, form, or json.
    """
    @wraps(f)
    def decorated_function(*args, **kwargs):
        user = get_user_from_token()
        
        if not user:
            return jsonify({"error": "Authentication required"}), 401
        
        # Check if user is admin/manager (bypass client check)
        user_roles = user.get('roles', [])
        is_admin = 'Admin' in user_roles or 'Manager' in user_roles
        
        if not is_admin:
            # Get client_id from various sources
            client_id = (
                request.args.get('clientId') or 
                request.form.get('clientId') or 
                (request.get_json() or {}).get('clientId')
            )
            
            if client_id:
                try:
                    client_id = int(client_id)
                    if not verify_client_ownership(client_id, user['user_id']):
                        return jsonify({"error": "Access denied to this client"}), 403
                except ValueError:
                    return jsonify({"error": "Invalid client ID"}), 400
        
        request.current_user = user
        return f(*args, **kwargs)
    
    return decorated_function

def require_auth(f):
    """Decorator to require authentication"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        user = get_user_from_token()
        
        if not user:
            return jsonify({"error": "Authentication required"}), 401
        
        # Add user to request context
        request.current_user = user
        return f(*args, **kwargs)
    
    return decorated_function

def require_role(*allowed_roles):
    """Decorator to require specific roles"""
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            user = get_user_from_token()
            
            if not user:
                return jsonify({"error": "Authentication required"}), 401
            
            user_roles = user.get('roles', [])
            
            if not any(role in user_roles for role in allowed_roles):
                return jsonify({
                    "error": "Insufficient permissions",
                    "required_roles": list(allowed_roles),
                    "your_roles": user_roles
                }), 403
            
            request.current_user = user
            return f(*args, **kwargs)
        
        return decorated_function
    return decorator

def check_document_quota():
    """Check if user has remaining document quota"""
    user = get_user_from_token()
    
    if not user:
        return False, "Authentication required"
    
    # Get real-time quota from database
    quota, processed = get_current_user_quota()
    
    if quota is None or processed is None:
        # Fallback to token values if database query fails
        quota = user.get('document_quota', 50)
        processed = user.get('documents_processed', 0)
    
    # -1 means unlimited (Admin)
    if quota == -1:
        return True, None
    
    if processed >= quota:
        return False, {
            "message": f"Document quota exceeded. You have processed {processed}/{quota} documents this month.",
            "quota": quota,
            "processed": processed,
            "quota_exceeded": True
        }
    
    return True, None

def log_user_action(action_type, details=None):
    """Log user actions for audit trail"""
    user = get_user_from_token()
    if not user:
        return
    
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Create audit log table if it doesn't exist
        cursor.execute("""
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'AuditLogs')
            CREATE TABLE AuditLogs (
                Id INT IDENTITY(1,1) PRIMARY KEY,
                UserId NVARCHAR(450) NOT NULL,
                UserEmail NVARCHAR(256),
                Action NVARCHAR(100) NOT NULL,
                Details NVARCHAR(MAX),
                IpAddress NVARCHAR(50),
                Timestamp DATETIME DEFAULT GETDATE()
            )
        """)
        
        ip_address = request.remote_addr
        details_json = str(details) if details else None
        
        cursor.execute("""
            INSERT INTO AuditLogs (UserId, UserEmail, Action, Details, IpAddress)
            VALUES (?, ?, ?, ?, ?)
        """, (user['user_id'], user['email'], action_type, details_json, ip_address))
        
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Error logging user action: {e}")

def get_user_clients(user_id=None):
    """
    Get all clients belonging to a user.
    If user_id is None, uses the current authenticated user.
    
    Args:
        user_id (str, optional): The user ID. Defaults to current user.
    
    Returns:
        list: List of client dictionaries
    """
    if user_id is None:
        user = get_user_from_token()
        if not user:
            return []
        user_id = user['user_id']
    
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT Id, Name, Email, Phone, Address, CreatedAt, UpdatedAt
            FROM Clients
            WHERE UserId = ?
            ORDER BY Name
        """, (user_id,))
        
        rows = cursor.fetchall()
        conn.close()
        
        clients = []
        for row in rows:
            clients.append({
                'id': row[0],
                'name': row[1],
                'email': row[2],
                'phone': row[3],
                'address': row[4],
                'createdAt': row[5].isoformat() if row[5] else None,
                'updatedAt': row[6].isoformat() if row[6] else None
            })
        
        return clients
    except Exception as e:
        print(f"Error getting user clients: {e}")
        return []

def is_admin_or_manager():
    """
    Check if the current user has Admin or Manager role.
    
    Returns:
        bool: True if user is Admin or Manager, False otherwise
    """
    user = get_user_from_token()
    if not user:
        return False
    
    user_roles = user.get('roles', [])
    return 'Admin' in user_roles or 'Manager' in user_roles