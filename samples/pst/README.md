# PST fixtures — read this before using them

The three `.pst` files in this directory are **synthetic fixtures**, not genuine Outlook PST files.

| File | Size | Status |
| :--- | ---: | :--- |
| `sample_unicode.pst` | 23,040 B | Synthetic. Header only; not readable by Outlook or libpff. |
| `sample_standard.pst` | 23,040 B | Synthetic. Header only; not readable by Outlook or libpff. |
| `sample_encrypted_password.pst` | 20,992 B | Synthetic. Header only; not readable by Outlook or libpff. |

They were produced by an early prototype of the sibling `mail2rust` PST writer and are retained
only so that format-sniffing code has a `!BDN`-signed file to identify. Independent verification
with libpff reports:

```
libpff_file_header_read_data: unsupported content type: 0x534d.
```

## Why this matters

`mail2mail` does not implement the PST binary format. It delegates PST reading and writing to
Outlook through MAPI (`pywin32`). Correctness therefore comes from Outlook, not from these
fixtures. Do not treat a successful parse of these files as evidence of PST compatibility.

## Creating a real PST for testing

On Windows with classic Outlook installed:

```powershell
mail2pst ..\eml out.pst
```

Verify the result in Outlook by opening it with **File → Open & Export → Open Outlook Data File**,
or independently with libpff:

```python
import pypff
f = pypff.file()
f.open("out.pst")
print(f.get_root_folder().number_of_sub_folders)
f.close()
```
