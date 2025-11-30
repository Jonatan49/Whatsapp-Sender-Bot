#!/usr/bin/env python3
"""Test script to verify imports without Qt or Selenium."""

import sys
import os

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

print("Testing imports...")
print("=" * 60)

# Test 1: Database models
try:
    from src.models.user import User
    from src.models.contact import Contact
    from src.models.campaign import Campaign
    from src.models.message import Message
    from src.models.template import MessageTemplate
    print("✓ All models imported successfully")
except Exception as e:
    print(f"✗ Models import failed: {e}")
    import traceback
    traceback.print_exc()

# Test 2: Core database
try:
    from src.core.database import Database
    print("✓ Database class imported successfully")
except Exception as e:
    print(f"✗ Database import failed: {e}")

# Test 3: Config
try:
    from src.core.config import Config
    print("✓ Config class imported successfully")
except Exception as e:
    print(f"✗ Config import failed: {e}")

# Test 4: Logger
try:
    from src.utils.logger import get_logger
    logger = get_logger('test')
    print("✓ Logger imported and instantiated successfully")
except Exception as e:
    print(f"✗ Logger import failed: {e}")

# Test 5: Services
try:
    from src.services.auth_service import AuthService
    print("✓ AuthService imported successfully")
except Exception as e:
    print(f"✗ AuthService import failed: {e}")

# Test 6: Check AuthService signature
try:
    import inspect
    from src.services.auth_service import AuthService
    sig = inspect.signature(AuthService.__init__)
    params = list(sig.parameters.keys())
    print(f"✓ AuthService.__init__ parameters: {params}")
    if len(params) == 1 and params[0] == 'self':
        print("✓ AuthService.__init__ signature is correct (only self)")
    else:
        print(f"✗ AuthService.__init__ has unexpected parameters: {params}")
except Exception as e:
    print(f"✗ AuthService signature check failed: {e}")

# Test 7: UI modules (without importing PyQt5)
try:
    # Just check if files exist
    ui_files = [
        'src/ui/__init__.py',
        'src/ui/login_window.py',
        'src/ui/main_window.py'
    ]
    for f in ui_files:
        if os.path.exists(f):
            print(f"✓ {f} exists")
        else:
            print(f"✗ {f} missing")
except Exception as e:
    print(f"✗ UI file check failed: {e}")

# Test 8: Verify login_window uses AuthService correctly
try:
    with open('src/ui/login_window.py', 'r', encoding='utf-8') as f:
        content = f.read()
        if 'self.auth_service = AuthService()' in content:
            print("✓ login_window.py calls AuthService() correctly (no db parameter)")
        elif 'AuthService(self.db)' in content:
            print("✗ login_window.py still calls AuthService(self.db) - NEEDS FIX")
        else:
            print("? Could not verify AuthService initialization in login_window.py")
except Exception as e:
    print(f"✗ login_window.py check failed: {e}")

print("=" * 60)
print("Import tests complete!")
