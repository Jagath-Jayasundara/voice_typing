import sys
from unittest.mock import MagicMock

# Mock windows/speech recognition specific imports if not installed
sys.modules['speech_recognition'] = MagicMock()
sys.modules['win32com'] = MagicMock()
sys.modules['win32com.client'] = MagicMock()

import unittest
from voice_number_entry import words_to_number, convert_number

class TestVoiceNumberEntry(unittest.TestCase):

    def test_words_to_number_english(self):
        self.assertEqual(words_to_number("one hundred twenty five"), 125)
        self.assertEqual(words_to_number("two thousand five hundred"), 2500)
        self.assertEqual(words_to_number("twenty five thousand"), 25000)
        self.assertEqual(words_to_number("zero"), 0)

    def test_words_to_number_sinhala(self):
        self.assertEqual(words_to_number("තුන"), 3)
        self.assertEqual(words_to_number("දහය"), 10)
        self.assertEqual(words_to_number("සියය"), 100)
        self.assertEqual(words_to_number("එක දහස් පන්සියය"), 1500)
        self.assertEqual(words_to_number("෧෨෩"), 123)

    def test_convert_number(self):
        self.assertEqual(convert_number("125"), 125)
        self.assertEqual(convert_number("please enter 500"), 500)
        self.assertEqual(convert_number("one hundred twenty five"), 125)
        self.assertEqual(convert_number("3.14"), 3.14)
        self.assertEqual(convert_number("කරුණාකර 50"), 50)
        self.assertEqual(convert_number("තුන"), 3)

if __name__ == "__main__":
    unittest.main()
