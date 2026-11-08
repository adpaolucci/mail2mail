#!/usr/bin/env python3
"""
Unit tests for file_picker.py - Native Host OS File and Folder Selection
"""

import os
import sys
import unittest
from unittest.mock import MagicMock, patch

import file_picker


class TestFilePicker(unittest.TestCase):
    def test_filetypes_defined(self):
        self.assertGreater(len(file_picker.MAIL_FILETYPES), 0)
        self.assertGreater(len(file_picker.PST_SAVE_FILETYPES), 0)
        self.assertGreater(len(file_picker.MBOX_SAVE_FILETYPES), 0)
        self.assertGreater(len(file_picker.CSV_SAVE_FILETYPES), 0)

    def test_is_headless(self):
        with patch.dict(os.environ, {"CI": "true"}, clear=False):
            self.assertTrue(file_picker._is_headless())

        with patch.dict(os.environ, {"HEADLESS": "1"}, clear=False):
            self.assertTrue(file_picker._is_headless())

        with patch.dict(os.environ, {}, clear=True):
            with patch("sys.platform", "linux"):
                self.assertTrue(file_picker._is_headless())

            with patch("sys.platform", "linux"):
                with patch.dict(os.environ, {"DISPLAY": ":0"}):
                    self.assertFalse(file_picker._is_headless())

    def test_tkinter_available_headless(self):
        with patch("file_picker._is_headless", return_value=True):
            self.assertFalse(file_picker._tkinter_available())

    @patch("tkinter.Tk")
    @patch("tkinter.filedialog.askopenfilename", return_value="C:\\Mail\\test.mbox")
    def test_tk_pick_open_file_single(self, mock_ask, mock_tk_cls):
        mock_root = MagicMock()
        mock_tk_cls.return_value = mock_root

        res = file_picker._tk_pick_open_file(
            title="Open Mail",
            filetypes=file_picker.MAIL_FILETYPES,
            multiple=False,
            initialdir=None,
        )
        self.assertEqual(res, os.path.normpath("C:\\Mail\\test.mbox"))
        mock_root.withdraw.assert_called_once()
        mock_root.destroy.assert_called_once()

    @patch("tkinter.Tk")
    @patch("tkinter.filedialog.askopenfilenames", return_value=["C:\\Mail\\1.mbox", "C:\\Mail\\2.mbox"])
    def test_tk_pick_open_file_multiple(self, mock_ask, mock_tk_cls):
        mock_root = MagicMock()
        mock_tk_cls.return_value = mock_root

        res = file_picker._tk_pick_open_file(
            title="Open Mail",
            filetypes=file_picker.MAIL_FILETYPES,
            multiple=True,
            initialdir=None,
        )
        self.assertEqual(res, [os.path.normpath("C:\\Mail\\1.mbox"), os.path.normpath("C:\\Mail\\2.mbox")])
        mock_root.destroy.assert_called_once()

    @patch("tkinter.Tk")
    @patch("tkinter.filedialog.askdirectory", return_value="C:\\Mail\\Thunderbird")
    def test_tk_pick_folder(self, mock_ask, mock_tk_cls):
        mock_root = MagicMock()
        mock_tk_cls.return_value = mock_root

        res = file_picker._tk_pick_folder(title="Folder", initialdir=None)
        self.assertEqual(res, os.path.normpath("C:\\Mail\\Thunderbird"))
        mock_root.destroy.assert_called_once()

    @patch("tkinter.Tk")
    @patch("tkinter.filedialog.asksaveasfilename", return_value="C:\\Output\\archive.pst")
    def test_tk_pick_save_file(self, mock_ask, mock_tk_cls):
        mock_root = MagicMock()
        mock_tk_cls.return_value = mock_root

        res = file_picker._tk_pick_save_file(
            title="Save",
            default_ext=".pst",
            filetypes=file_picker.PST_SAVE_FILETYPES,
            initialfile="archive.pst",
            initialdir=None,
        )
        self.assertEqual(res, os.path.normpath("C:\\Output\\archive.pst"))
        mock_root.destroy.assert_called_once()

    @patch("file_picker._tkinter_available", return_value=True)
    @patch("file_picker._tk_pick_open_file", return_value=os.path.normpath("C:\\Mail\\test.mbox"))
    def test_pick_open_file_success(self, mock_tk_pick, mock_tk_avail):
        res = file_picker.pick_open_file(title="Test")
        self.assertEqual(res, os.path.normpath("C:\\Mail\\test.mbox"))

    @patch("file_picker._tkinter_available", return_value=True)
    @patch("file_picker._tk_pick_open_file", return_value=None)
    def test_pick_open_file_cancelled(self, mock_tk_pick, mock_tk_avail):
        res = file_picker.pick_open_file(title="Test")
        self.assertIsNone(res)

    @patch("file_picker._tkinter_available", return_value=True)
    @patch("file_picker._tk_pick_folder", return_value=os.path.normpath("C:\\Mail\\Profile"))
    def test_pick_folder_success(self, mock_tk_folder, mock_tk_avail):
        res = file_picker.pick_folder(title="Test")
        self.assertEqual(res, os.path.normpath("C:\\Mail\\Profile"))

    @patch("file_picker._tkinter_available", return_value=True)
    @patch("file_picker._tk_pick_folder", return_value=None)
    def test_pick_folder_cancelled(self, mock_tk_folder, mock_tk_avail):
        res = file_picker.pick_folder(title="Test")
        self.assertIsNone(res)

    @patch("file_picker._tkinter_available", return_value=True)
    @patch("file_picker._tk_pick_save_file", return_value=os.path.normpath("C:\\Mail\\out.pst"))
    def test_pick_save_file_success(self, mock_tk_save, mock_tk_avail):
        res = file_picker.pick_save_file(title="Test", default_ext=".pst")
        self.assertEqual(res, os.path.normpath("C:\\Mail\\out.pst"))

    @patch("file_picker._tkinter_available", return_value=True)
    @patch("file_picker._tk_pick_save_file", return_value=None)
    def test_pick_save_file_cancelled(self, mock_tk_save, mock_tk_avail):
        res = file_picker.pick_save_file(title="Test", default_ext=".pst")
        self.assertIsNone(res)

    @patch("file_picker._is_headless", return_value=True)
    @patch("file_picker._tkinter_available", return_value=False)
    def test_headless_mode_returns_none(self, mock_tk_avail, mock_headless):
        self.assertIsNone(file_picker.pick_open_file(title="Test"))
        self.assertIsNone(file_picker.pick_folder(title="Test"))
        self.assertIsNone(file_picker.pick_save_file(title="Test"))

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
