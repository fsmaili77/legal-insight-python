try:
    from config import SQLSERVER_CONN_STRING, GEMINI_API_KEY
    print(f"✓ Config loaded successfully")
    print(f"✓ Connection string: {SQLSERVER_CONN_STRING[:50]}...")
    print(f"✓ Gemini API: {'Configured' if GEMINI_API_KEY else 'Not configured'}")
except SystemExit:
    print("✗ Configuration validation failed - check the errors above")
except Exception as e:
    print(f"✗ Error importing config: {e}")