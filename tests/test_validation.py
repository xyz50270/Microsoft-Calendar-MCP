from src.utils.validation import validate_iso_datetime, validate_email, validate_enum

def test_validate_iso_datetime():
    """测试 ISO 日期验证逻辑"""
    print("Testing ISO datetime validation...")
    
    # Valid
    validate_iso_datetime("2025-12-23T09:00:00", "test")
    validate_iso_datetime("2025-12-23", "test")
    validate_iso_datetime(None, "test")
    
    # Invalid format
    try:
        validate_iso_datetime("2025/12/23", "date_field")
        return False, "Failed to catch slash format"
    except ValueError as e:
        assert "must be an ISO format date string" in str(e)
        
    # Invalid logic (e.g., month 13)
    try:
        validate_iso_datetime("2025-13-01", "date_field")
        return False, "Failed to catch invalid month"
    except ValueError as e:
        assert "Invalid ISO datetime" in str(e)
        
    return True, ""

def test_validate_email():
    """测试邮箱验证逻辑"""
    print("Testing email validation...")
    
    # Valid
    validate_email("test@example.com", "email")
    validate_email("user.name+tag@domain.co.uk", "email")
    
    # Invalid
    invalid_emails = ["no-at-sign", "too@many@ats.com", "@no-user.com", "no-domain@"]
    for email in invalid_emails:
        try:
            validate_email(email, "email_field")
            return False, f"Failed to catch invalid email: {email}"
        except ValueError as e:
            assert "Invalid email address" in str(e)
            
    return True, ""

def test_validate_enum():
    """测试枚举验证逻辑"""
    print("Testing enum validation...")
    
    valid_options = ["low", "normal", "high"]
    
    # Valid
    validate_enum("low", valid_options, "importance")
    validate_enum("NORMAL", valid_options, "importance") # Case insensitive check
    validate_enum(None, valid_options, "importance")
    
    # Invalid
    try:
        validate_enum("Urgent", valid_options, "importance")
        return False, "Failed to catch invalid enum value"
    except ValueError as e:
        assert "Must be one of ['low', 'normal', 'high']" in str(e)
        
    return True, ""

if __name__ == "__main__":
    print("Running Unit Tests for Validation Logic...")
    
    tests = [
        test_validate_iso_datetime,
        test_validate_email,
        test_validate_enum
    ]
    
    passed = 0
    for test in tests:
        try:
            success, msg = test()
            if success:
                print(f"✅ PASS: {test.__name__}")
                passed += 1
            else:
                print(f"❌ FAIL: {test.__name__} - {msg}")
        except Exception as e:
            print(f"💥 ERROR: {test.__name__} - {str(e)}")
            
    print(f"\nFinal Result: {passed}/{len(tests)} passed.")
    exit(0 if passed == len(tests) else 1)
