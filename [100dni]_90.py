class NegativeNumberException(Exception):
    pass

def divide(a, b):
    if not (isinstance(a, (int, float)) and isinstance(b, (int, float))):
        raise TypeError("Both a and b must be numbers")
    if b == 0:
        raise ZeroDivisionError("Cannot divide by zero")
    if a < 0 or b < 0:
        raise NegativeNumberException("Negative numbers are not allowed")

    return a / b


# Przykład użycia bloku try...except do obsługi wyjątków
try:
    result = divide(10, 2)
    print("Wynik dzielenia:", result)

    # Spróbuj podać argumenty, które spowodują wyjątki
    #divide("abc", 2)
    #divide(10, 0)
    #divide(-10, 2)

except TypeError as e:
    print(f"TypeError: {e}")
except ZeroDivisionError as e:
    print(f"ZeroDivisionError: {e}")
except NegativeNumberException as e:
    print(f"NegativeNumberException: {e}")
