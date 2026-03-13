from django.test import TestCase
from core.classifier import classify_element_simple
from core.analyzer import check_gost_compliance

class GostCoreTest(TestCase):
    def test_heading_classification(self):
        """Проверка, что заголовки определяются верно"""
        p = {"text": "ВВЕДЕНИЕ", "bold": True}
        self.assertEqual(classify_element_simple(p), "heading")
        
    def test_paragraph_classification(self):
        """Проверка обычного текста"""
        p = {"text": "Это обычный текст абзаца.", "bold": False, "has_numbering": False}
        self.assertEqual(classify_element_simple(p), "text")

    def test_gost_compliance_valid(self):
        """Проверка идеального параграфа"""
        p = {
            "font_name": "Times New Roman",
            "font_size_pt": 14,
            "line_spacing": 1.5,
            "first_indent_cm": 1.25
        }
        status, errors, warnings = check_gost_compliance(p, "text")
        self.assertEqual(status, "correct")

    def test_gost_compliance_error(self):
        """Проверка фиксации ошибки шрифта"""
        p = {"font_name": "Arial", "font_size_pt": 12}
        status, errors, warnings = check_gost_compliance(p, "text")
        self.assertEqual(status, "error")
        self.assertTrue(any("Шрифт" in e for e in errors))