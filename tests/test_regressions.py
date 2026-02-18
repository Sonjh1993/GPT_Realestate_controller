import ast
from pathlib import Path
import unittest

from app.matching import match_properties
from app.sheet_sync import _anonymize_rows


class DesktopAppStructureTests(unittest.TestCase):
    def test_critical_desktop_handlers_are_class_methods(self):
        source = Path("app/desktop_app.py").read_text(encoding="utf-8")
        module = ast.parse(source)

        ledger_class = next(
            node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "LedgerDesktopApp"
        )
        class_method_names = {node.name for node in ledger_class.body if isinstance(node, ast.FunctionDef)}

        required = {
            "refresh_tasks",
            "_build_property_ui",
            "_build_customer_ui",
            "_build_settings_ui",
            "export_sync",
            "open_export_folder",
            "refresh_all",
        }
        missing = required - class_method_names
        self.assertFalse(missing, f"Missing class methods: {sorted(missing)}")


class APIServiceRegressionTests(unittest.TestCase):
    def test_api_service_uses_supported_list_tasks_signature(self):
        source = Path("app/api_service.py").read_text(encoding="utf-8")
        self.assertIn("storage.list_tasks(include_done=False)", source)
        self.assertNotIn("storage.list_tasks(status=", source)


class AnonymizeRegressionTests(unittest.TestCase):
    def test_complex_name_not_masked_but_personal_name_and_phone_are(self):
        rows = [
            {
                "complex_name": "봉담자이",
                "customer_name": "홍길동",
                "owner_name": "김집주인",
                "phone": "010-1234-5678",
                "owner_phone": "01098765432",
            }
        ]
        masked = _anonymize_rows(rows)[0]
        self.assertEqual(masked["complex_name"], "봉담자이")
        self.assertEqual(masked["customer_name"], "")
        self.assertEqual(masked["owner_name"], "")
        self.assertEqual(masked["phone"], "5678")
        self.assertEqual(masked["owner_phone"], "5432")


class DealTypeMatchingRegressionTests(unittest.TestCase):
    def test_multi_deal_type_matches_intersection(self):
        customer = {"deal_type": "전세,월세"}
        props = [
            {"id": 1, "tab": "A", "deal_sale": 1, "deal_jeonse": 0, "deal_wolse": 0},
            {"id": 2, "tab": "A", "deal_sale": 0, "deal_jeonse": 1, "deal_wolse": 0},
            {"id": 3, "tab": "A", "deal_sale": 0, "deal_jeonse": 0, "deal_wolse": 1},
        ]
        got = {r.property_id for r in match_properties(customer, props, limit=10)}
        self.assertEqual(got, {2, 3})

    def test_empty_deal_type_matches_all(self):
        customer = {"deal_type": ""}
        props = [
            {"id": 1, "tab": "A", "deal_sale": 1, "deal_jeonse": 0, "deal_wolse": 0},
            {"id": 2, "tab": "A", "deal_sale": 0, "deal_jeonse": 1, "deal_wolse": 0},
            {"id": 3, "tab": "A", "deal_sale": 0, "deal_jeonse": 0, "deal_wolse": 1},
        ]
        got = {r.property_id for r in match_properties(customer, props, limit=10)}
        self.assertEqual(got, {1, 2, 3})

    def test_multi_deal_type_with_spaces(self):
        customer = {"deal_type": "전세, 월세"}
        props = [
            {"id": 1, "tab": "A", "deal_sale": 1, "deal_jeonse": 0, "deal_wolse": 0},
            {"id": 2, "tab": "A", "deal_sale": 0, "deal_jeonse": 1, "deal_wolse": 0},
            {"id": 3, "tab": "A", "deal_sale": 0, "deal_jeonse": 0, "deal_wolse": 1},
        ]
        got = {r.property_id for r in match_properties(customer, props, limit=10)}
        self.assertEqual(got, {2, 3})


if __name__ == "__main__":
    unittest.main()
