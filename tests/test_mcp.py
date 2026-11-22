#!/usr/bin/env python3
"""
tests/test_mcp.py — Unit tests for mail2mail Model Context Protocol (MCP) server
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

import mail2mcp


class TestMail2Mcp(unittest.TestCase):
    def setUp(self):
        self.samples_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "samples"))
        self.sample_mbox = os.path.join(self.samples_dir, "mbox", "sample_mailbox.mbox")

    def test_tools_manifest_structure(self):
        tools = mail2mcp.get_tools_manifest()
        self.assertIsInstance(tools, list)
        self.assertGreaterEqual(len(tools), 5)

        tool_names = [t["name"] for t in tools]
        self.assertIn("mail2mail_info", tool_names)
        self.assertIn("mail2mail_sniff", tool_names)
        self.assertIn("mail2mail_convert", tool_names)
        self.assertIn("mail2mail_audit", tool_names)
        self.assertIn("mail2mail_split", tool_names)

        for t in tools:
            self.assertIn("inputSchema", t)
            self.assertEqual(t["inputSchema"]["type"], "object")

    def test_system_info(self):
        info = mail2mcp.get_system_info()
        self.assertEqual(info["server"], "mail2mail-mcp")
        self.assertEqual(info["version"], "2.3.0")
        self.assertIn("python_version", info)
        self.assertIn("supported_formats", info)
        self.assertIsInstance(info["supported_formats"], list)

    def test_classify_mailbox_format_mbox(self):
        if os.path.exists(self.sample_mbox):
            res = mail2mcp.classify_mailbox_format(self.sample_mbox)
            self.assertEqual(res["format"], "MBOX")
            self.assertFalse(res["is_directory"])
            self.assertGreater(res["size_bytes"], 0)

    def test_classify_mailbox_format_directory(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            eml_file = os.path.join(tmpdir, "test.eml")
            with open(eml_file, "wb") as f:
                f.write(b"From: a@b.com\r\nSubject: Test\r\n\r\nBody")
            res = mail2mcp.classify_mailbox_format(tmpdir)
            self.assertTrue(res["is_directory"])
            self.assertEqual(res["format"], "EML_TREE")

    def test_handle_jsonrpc_initialize(self):
        req = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {"protocolVersion": "2024-11-05"},
        }
        resp = mail2mcp.handle_jsonrpc(req)
        self.assertIsNotNone(resp)
        self.assertEqual(resp["id"], 1)
        self.assertEqual(resp["result"]["protocolVersion"], "2024-11-05")
        self.assertEqual(resp["result"]["serverInfo"]["name"], "mail2mail-mcp")

    def test_handle_jsonrpc_tools_list(self):
        req = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/list",
            "params": {},
        }
        resp = mail2mcp.handle_jsonrpc(req)
        self.assertIsNotNone(resp)
        self.assertEqual(resp["id"], 2)
        tools = resp["result"]["tools"]
        self.assertGreaterEqual(len(tools), 5)

    def test_handle_jsonrpc_tools_call_sniff(self):
        if os.path.exists(self.sample_mbox):
            req = {
                "jsonrpc": "2.0",
                "id": 3,
                "method": "tools/call",
                "params": {
                    "name": "mail2mail_sniff",
                    "arguments": {"path": self.sample_mbox},
                },
            }
            resp = mail2mcp.handle_jsonrpc(req)
            self.assertIsNotNone(resp)
            self.assertEqual(resp["id"], 3)
            self.assertFalse(resp["result"]["isError"])
            content_text = resp["result"]["content"][0]["text"]
            data = json.loads(content_text)
            self.assertEqual(data["format"], "MBOX")

    def test_handle_jsonrpc_tools_call_audit(self):
        if os.path.exists(self.sample_mbox):
            with tempfile.TemporaryDirectory() as tmpdir:
                csv_path = os.path.join(tmpdir, "audit_test.csv")
                req = {
                    "jsonrpc": "2.0",
                    "id": 4,
                    "method": "tools/call",
                    "params": {
                        "name": "mail2mail_audit",
                        "arguments": {
                            "path": self.sample_mbox,
                            "output_csv": csv_path,
                            "bates_prefix": "CASE-TEST",
                        },
                    },
                }
                resp = mail2mcp.handle_jsonrpc(req)
                self.assertIsNotNone(resp)
                self.assertEqual(resp["id"], 4)
                self.assertFalse(resp["result"]["isError"])
                data = json.loads(resp["result"]["content"][0]["text"])
                self.assertEqual(data["status"], "success")
                self.assertGreater(data["total_items"], 0)
                self.assertTrue(data["bates_start"].startswith("CASE-TEST"))
                self.assertTrue(os.path.exists(csv_path))

    def test_handle_jsonrpc_unknown_method(self):
        req = {
            "jsonrpc": "2.0",
            "id": 99,
            "method": "invalid_method",
        }
        resp = mail2mcp.handle_jsonrpc(req)
        self.assertIsNotNone(resp)
        self.assertEqual(resp["error"]["code"], -32601)

    def test_cli_subcommand_mcp_tools(self):
        script = os.path.join(os.path.dirname(__file__), "..", "mail2mail.py")
        proc = subprocess.run(
            [sys.executable, script, "mcp", "--tools"],
            capture_output=True,
            text=True,
            check=True,
        )
        tools = json.loads(proc.stdout)
        self.assertIsInstance(tools, list)
        self.assertGreaterEqual(len(tools), 5)
        convert_tool = next(t for t in tools if t["name"] == "mail2mail_convert")
        self.assertNotIn("bates_prefix", convert_tool["inputSchema"]["properties"])

    def test_run_split_cli_flags(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            src = os.path.join(tmpdir, "box.mbox")
            with open(src, "wb") as f:
                f.write(b"From a@b.com Fri Nov 6 10:00:00 2026\r\nFrom: a@b.com\r\nSubject: S\r\n\r\nBody\r\n")
            out_dir = os.path.join(tmpdir, "out")

            ok_proc = MagicMock(returncode=0, stdout="ok", stderr="")
            with patch("mail2mcp.subprocess.run", return_value=ok_proc) as mock_run:
                mail2mcp.run_split(src, out_dir, split_by="year", target_format="eml")
                eml_cmd = mock_run.call_args[0][0]
                self.assertIn("--layout", eml_cmd)
                self.assertIn("year", eml_cmd)
                self.assertNotIn("--split-by", eml_cmd)

                mail2mcp.run_split(src, out_dir, split_by="year", target_format="mbox")
                mbox_cmd = mock_run.call_args[0][0]
                self.assertIn("--layout", mbox_cmd)
                self.assertIn("year", mbox_cmd)
                self.assertNotIn("--split-by", mbox_cmd)

            with patch("mail2mcp.get_system_info", return_value={"pst_export_supported_locally": True}):
                with patch("mail2mcp.subprocess.run", return_value=ok_proc) as mock_run:
                    mail2mcp.run_split(src, out_dir, max_volume_mb=512, split_by="size", target_format="pst")
                    pst_cmd = mock_run.call_args[0][0]
                    self.assertIn("--src", pst_cmd)
                    self.assertIn("--out-dir", pst_cmd)
                    self.assertIn("--max-pst-size", pst_cmd)
                    self.assertIn("512MB", pst_cmd)

            with patch("mail2mcp.get_system_info", return_value={"pst_export_supported_locally": False}):
                err_res = mail2mcp.run_split(src, out_dir, target_format="pst")
                self.assertEqual(err_res["status"], "error")
                self.assertEqual(err_res["error_type"], "PST_REQUIRES_OUTLOOK")

    def test_run_audit_neutralizes_csv_formulas(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            src_root = os.path.join(tmpdir, "src")
            evil_dir = os.path.join(src_root, "=CMD('calc')")
            os.makedirs(evil_dir, exist_ok=True)
            eml_path = os.path.join(evil_dir, "+evil.eml")
            with open(eml_path, "wb") as f:
                f.write(b"From: a@b.com\r\nSubject: Test\r\n\r\nBody")
            csv_path = os.path.join(tmpdir, "audit.csv")
            res = mail2mcp.run_audit(src_root, output_csv=csv_path, bates_prefix="@BATES")
            self.assertEqual(res["status"], "success")
            import csv
            with open(csv_path, "r", encoding="utf-8", newline="") as f:
                rows = list(csv.DictReader(f))
            self.assertEqual(len(rows), 1)
            self.assertTrue(rows[0]["bates_number"].startswith("'@BATES"))
            self.assertTrue(rows[0]["folder"].startswith("'="))
            self.assertTrue(rows[0]["source_name"].startswith("'+"))


if __name__ == "__main__":
    unittest.main()
