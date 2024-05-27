import unittest

# Import functions here
from index import openpyxl_operations

# unittest uses class type structure for tests
class TestFormatData(unittest.TestCase):
    # Use individual tests for checking different parts of the function
    def test_for_scenario(self):
        self.assertEqual(openpyxl_operations("""Function arguments here"""), """ Output message here """)

    def test_for_other_scenario(self):
        self.assertEqual(openpyxl_operations("""Function arguments here"""), """ Output message here """)


# Different equality statements to use
# assertEqual(a, b)             a == b
# assertNotEqual(a, b)          a != b
# assertTrue(x)                 bool(x) is True
# assertFalse(x)                bool(x) is False
# assertIs(a, b)                a is b
# assertIsNot(a, b)             a is not b
# assertIsNone(x)               x is None
# assertIsNotNone(x)            x is not None
# assertIn(a, b)                a in b
# assertNotIn(a, b)             a not in b
# assertIsInstance(a, b)        isinstance(a, b)
# assertNotIsInstance(a, b)     not isinstance(a, b)