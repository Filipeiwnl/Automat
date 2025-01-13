def safe_float_conversion(value):
    try:
        return float(value)
    except ValueError:
        return 0.0

def is_numeric(value):
    try:
        float(value)
        return True
    except ValueError:
        return False
