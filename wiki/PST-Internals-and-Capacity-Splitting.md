# PST Internals & Capacity Auto-Splitting (Overcoming the 50GB Limit)

Microsoft Outlook's **Personal Storage Table (`.pst`)** and **Offline Storage Table (`.ost`)** formats are notoriously complex, proprietary database structures.

`mail2rust` includes a pure native Rust implementation (`mail2rust-pst`) that builds valid Unicode PST files **without requiring Microsoft Outlook or Windows MAPI libraries**, enabling automated PST generation on **Linux, macOS, and Windows**.

---

## 1. PST Architecture: ANSI vs Unicode

| Specification | ANSI PST (Legacy Outlook 97–2002) | Unicode PST (Modern Outlook 2003+) |
| :--- | :--- | :--- |
| **Header Magic** | `!BDN` (`0x21 0x42 0x44 0x4E`) | `!BDN` (`0x21 0x42 0x44 0x4E`) |
| **Format Version** | `14` or `15` | `23` (`0x17`) |
| **Pointer Size** | 32-bit file offsets | **64-bit file offsets** |
| **Max Capacity Limit** | **2 GB** (Hard ceiling) | **50 GB** (Default Outlook ceiling) |
| **Overflow Behavior** | **Irreversible data corruption** | Outlook freezes, index lockups, `0x80040119` |

### Key Internal Structures
1. **Header (512 bytes):** Stores root references to the Node B-Tree (`rNBT`) and Block B-Tree (`rBBT`), CRC32 checksums, encryption flags (`0x00 = None`, `0x01 = Permutation/Compressible`, `0x02 = Cyclic`), and allocation counters.
2. **NBT (Node B-Tree):** Indexes all logical MAPI objects (folders, messages, attachments, recipient tables) mapped to unique 32-bit Node IDs (`NID`).
3. **BBT (Block B-Tree):** Indexes physical disk blocks on disk mapped to 64-bit Block IDs (`BID`), managing deduplicated block allocation and sector alignment.
4. **Property Heaps (pc) & Tables (tc):** Variable-length B-Trees inside individual blocks holding MAPI property tags (`PidTag...`) and folder hierarchy lists.

---

## 2. The 50GB PST Limit & Why Auto-Splitting is Essential

While Unicode PST 64-bit architecture theoretically supports up to 4 Petabytes, **Microsoft Outlook enforces a default 50 GB administrative cap**.

### Consequences of Exceeding 45–50 GB in Outlook:
* **`0x80040119` Error:** Outlook refuses to write new emails or move folders.
* **Search Index Freezes:** Windows Desktop Search slows down exponentially or crashes.
* **Scanpst.exe Failure:** Microsoft's Inbox Repair Tool (`scanpst.exe`) runs out of memory or aborts on files $> 50\text{ GB}$.
* **High Corruption Risk:** System crashes or power loss during writes to a 50+ GB PST frequently corrupt the root B-Tree nodes.

---

## 3. Dynamic Capacity Splitting in `mail2rust`

To resolve this limitation, `mail2rust` implements **dynamic volume splitting** via the `--split-size` flag:

```bash
# Split large archive into consecutive 20 GB volume files
mail2rust convert large_export.mbox archive.pst --split-size 20GB
```

### 3.1 Naming Convention & Sequence
When splitting is active, the engine monitors accumulated physical file size and closes volumes at safe message boundaries:
1. `archive.pst` (Initial volume, holds first 20 GB)
2. `archive_001.pst` (Second volume, holds next 20 GB)
3. `archive_002.pst` (Third volume, holds next 20 GB)

### 3.2 Folder Hierarchy Replication
When a volume reaches capacity:
1. The active PST is cleanly finalized (flushing the BBT, NBT, and updating the 512-byte root header checksums).
2. A new empty Unicode PST is initialized.
3. **The full folder hierarchy is replicated** into the new volume (e.g. `Inbox/Subfolder/ProjectA`), and subsequent messages are inserted into their corresponding folders in the new volume.
4. Messages and attachments are never split midway through an item; volume cuts occur cleanly between whole message boundaries.

---

## 4. Multi-PST Auto-Merge

Conversely, when multiple segmented PST files exist (e.g. legacy annual archives or multi-part exports), `mail2rust` and `mail2mail` can merge them back into a single consolidated PST:

```bash
mail2rust merge \
  --input "archive_2021.pst" \
  --input "archive_2022.pst" \
  --input "archive_2023.pst" \
  --dedup message-id \
  --out "master_archive.pst"
```
