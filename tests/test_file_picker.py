#!/usr/bin/env python3
"""
Unit tests for file_picker.py - Native Host OS File and Folder Selection
"""

import os
import unittest
from unittest.mock import MagicMock, patch

import file_picker


class TestFilePicker(unittest.TestCase):
    def test_filetypes_defined(self):
        self.assertGreater(len(file_picker.MAIL_FILETYPES), 0)
        self.assertGreater(len(file_picker.PST_SAVE_FILETYPES), 0)
        self.assertGreater(len(file_picker.MBOX_SAVE_FILETYPES), 0)
        self.assertGreater(len(file_picker.CSV_SAVE_FILETYPES), 0)

    @patch("file_picker._tkinter_available", return_value=True)
    def test_pick_open_file_mocked(self, mock_tk_avail):
        with patch("tkinter.filedialog.askopenfilename", return_value="C:\\Mail\\test.mbox"):
            res = file_picker.pick_open_file(title="Test")
            self.assertEqual(res, os.path.normpath("C:\\Mail\\test.mbox"))

    @patch("file_picker._tkinter_available", return_value=True)
    def test_pick_open_file_cancelled(self, mock_tk_avail):
        with patch("tkinter.filedialog.askopenfilename", return_value=""):
            res = file_picker.pick_open_file(title="Test")
            self.assertIsNone(res)

    @patch("file_picker._tkinter_available", return_value=True)
    def test_pick_open_files_multiple(self, mock_tk_avail):
        with patch("tkinter.filedialog.askopenfilenames", return_value=["C:\\Mail\\1.mbox", "C:\\Mail\\2.mbox"]):
            res = file_picker.pick_open_file(title="Test", multiple=True)
            self.assertEqual(len(res), 2)
            self.assertEqual(res[0], os.path.normpath("C:\\Mail\\1.mbox"))

    @patch("file_picker._tkinter_available", return_value=True)
    def test_pick_folder_mocked(self, mock_tk_avail):
        with patch("tkinter.filedialog.askdirectory", return_value="C:\\Mail\\Thunderbird"):
            res = file_picker.pick_folder(title="Test")
            self.assertEqual(res, os.path.normpath("C:\\Mail\\Thunderbird"))

    @patch("file_picker._tkinter_available", return_value=True)
    def test_pick_save_file_mocked(self, mock_tk_avail):
        with patch("tkinter.filedialog.asksaveasfilename", return_value="C:\\Output\\archive.pst"):
            res = file_picker.pick_save_file(title="Test", default_ext=".pst")
            self.assertEqual(res, os.path.normpath("C:\\Output\\archive.pst"))

    @patch("sys.stdin.isatty", return_value=True)
    @patch("builtins.input", return_value="2")
    @patch("file_picker.pick_folder", return_value="C:\\Mail\\ImapMail")
    def test_pick_source_interactive_directory(self, mock_folder, mock_input, mock_isatty):
        res = file_picker.pick_source(title="Select Mail Source")
        self.assertEqual(res, "C:\\Mail\\ImapMail")
        mock_folder.assert_called_once()

    @patch("sys.stdin.isatty", return_value=True)
    @patch("builtins.input", return_value="1")
    @patch("file_picker.pick_open_file", return_value="C:\\Mail\\inbox.pst")
    def test_pick_source_interactive_file(self, mock_file, mock_input, mock_isatty):
        res = file_picker.pick_source(title="Select Mail Source")
        self.assertEqual(res, "C:\\Mail\\inbox.pst")
        mock_file.assert_called_once()


if __name__ == "__main__":
    unittest.main()
