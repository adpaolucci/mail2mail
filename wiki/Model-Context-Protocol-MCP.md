# Model Context Protocol (MCP) Guide for `mail2mail`

The **Model Context Protocol (MCP)** enables AI agents (Claude Desktop, Cursor, Google Antigravity, Windsurf, AutoGPT) to use `mail2mail` as a native local tool server over standard JSON-RPC 2.0 stdio.

---

## 1. Exposed Tools (5 Total)

| Tool Name | Description | Key Inputs | Notes |
| :--- | :--- | :--- | :--- |
| **`mail2mail_info`** | Inspects Python version, OS platform, Outlook COM availability, and supported formats. | *(None)* | Returns whether local Outlook MAPI PST creation is supported. |
| **`mail2mail_sniff`** | Classifies mailbox formats and magic bytes across 15+ formats. | `path` | Fast in-process detection of files and directory hierarchies. |
| **`mail2mail_convert`** | Converts between EML, MBOX, and PST with optional deduplication, date filtering, and audit log generation. | `source`, `destination`, `target_format`, `dedup`, `since`, `until`, `audit_log` | Gracefully handles non-Windows PST requests. |
| **`mail2mail_audit`** | Calculates SHA-256 Merkle root, applies Bates numbering, exports CSV manifest. | `path`, `output_csv`, `case_number`, `examiner`, `bates_prefix` | Tamper-evident forensic chain of custody. |
| **`mail2mail_split`** | Splits large mailboxes into volumes by size or year. | `source`, `destination_dir`, `max_volume_mb`, `split_by`, `target_format` | Preserves folder structures across split volumes. |

---

## 2. Server Execution

```bash
# Launch stdio JSON-RPC server
mail2mail mcp
# Or
python -m mail2mcp

# Print JSON schemas
mail2mail mcp --tools
```

---

## 3. Configuration for Agents

### Claude Desktop (`claude_desktop_config.json`)
```json
{
  "mcpServers": {
    "mail2mail": {
      "command": "python",
      "args": ["-m", "mail2mcp"]
    }
  }
}
```

### Cursor (`.cursor/mcp.json`)
```json
{
  "mcpServers": {
    "mail2mail": {
      "command": "mail2mcp"
    }
  }
}
```

---

## 4. Platform Note: PST Generation
* In `mail2mail`, PST generation requires Windows Microsoft Outlook (`pywin32` COM).
* For headless Linux/macOS servers or environments without Microsoft Outlook, use our companion pure-Rust tool [**`mail2rust`**](https://github.com/adpaolucci/mail2rust), which parses and builds 64-bit Unicode MS-PST binary containers natively across all operating systems.
