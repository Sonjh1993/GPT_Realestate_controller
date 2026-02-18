import ast
from pathlib import Path
import unittest


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


if __name__ == "__main__":
    unittest.main()
