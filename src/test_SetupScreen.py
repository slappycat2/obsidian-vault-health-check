import unittest
from PyQt5.QtWidgets import QApplication
from PyQt5.QtTest import QTest
from PyQt5.QtCore import Qt
from src.v_chk_setup import SetupScreen
from unittest.mock import MagicMock

class TestSetupScreen(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication([])  # Create a QApplication instance

    def setUp(self):
        # Mock the config object
        self.mock_config = MagicMock()
        self.mock_config.dir_vault = ""
        self.mock_config.pn_wb_exec = ""
        self.mock_config.dirs_skip_rel_str = ""
        self.mock_config.bool_shw_notes = True
        self.mock_config.bool_rel_paths = True
        self.mock_config.bool_summ_rows = True
        self.mock_config.bool_unused_1 = False
        self.mock_config.bool_unused_2 = False
        self.mock_config.bool_unused_3 = False
        self.mock_config.link_lim_vals = 0
        self.mock_config.link_lim_tags = 0
        self.mock_config.validate_dir_vault.return_value = (True, "")
        self.mock_config.validate_pn_wb_exec.return_value = (True, "")
        self.mock_config.validate_dirs_skip_rel_str.return_value = (True, "")

        self.screen = SetupScreen(self.mock_config)

    def test_initial_state(self):
        # Test initial state of widgets
        self.assertEqual(self.screen.dir_vault_combo.currentText(), "")
        self.assertEqual(self.screen.pn_wb_exec_edit.text(), "")
        self.assertEqual(self.screen.dirs_skip_rel_str_edit.text(), "")
        self.assertTrue(self.screen.bool_shw_notes_cb.isChecked())
        self.assertTrue(self.screen.bool_rel_paths_cb.isChecked())
        self.assertTrue(self.screen.bool_summ_rows_cb.isChecked())

    def test_browse_exec_path(self):
        # Simulate clicking the "Browse" button for the executable path
        QTest.mouseClick(self.screen.findChild(QPushButton, "Browse"), Qt.LeftButton)
        # Mock behavior ensures no crash; actual file dialog interaction is skipped

    def test_dropdown_selection(self):
        # Simulate selecting an item in the dropdown
        self.screen.dir_vault_combo.addItems(["Vault1", "Vault2", "Vault3"])
        self.screen.dir_vault_combo.setCurrentIndex(1)
        self.assertEqual(self.screen.dir_vault_combo.currentText(), "Vault2")

    def test_save_and_run(self):
        # Simulate clicking the "Save & Run" button
        save_button = self.screen.findChild(QPushButton, "Save & Run")
        QTest.mouseClick(save_button, Qt.LeftButton)
        self.mock_config.save_config.assert_called_once()  # Ensure save_config was called

    def tearDown(self):
        self.screen.close()

    @classmethod
    def tearDownClass(cls):
        cls.app.quit()

if __name__ == "__main__":
    unittest.main()