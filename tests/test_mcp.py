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
        self.assertEqual(info["version"], "2.2.0")
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


if __name__ == "__main__":
    unittest.main()
