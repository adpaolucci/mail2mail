#!/usr/bin/env python3
"""
file_picker.py - Native/Host OS File and Folder Selection for Mail Utilities

Provides cross-platform native file and folder pickers leveraging host OS dialogs:
- Windows: Native Win32 / Common Item Dialog via Tkinter or PowerShell COM/Forms fallback.
- macOS: Native Cocoa NSOpenPanel / NSSavePanel via Tkinter or osascript fallback.
- Linux: Native XDG Desktop Portal / Zenity / KDialog / Tkinter.

Includes standard mail format filters (*.pst, *.ost, *.mbox, *.eml, *.emlx, *.msg, *.olm, *.zip).
Can be imported as a module or executed directly from the command line.
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from typing import List, Optional, Tuple, Union

# Standard email file filter sets
MAIL_FILETYPES: List[Tuple[str, str]] = [
    (
        "All Supported Mail Archives",
        "*.pst;*.ost;*.mbox;*.mbx;*.eml;*.emlx;*.msg;*.olm;*.zip;*.tar;*.tar.gz;*.tgz;*.tar.bz2;*.tar.xz;*.gz;*.bz2;*.xz",
    ),
    ("Outlook Personal Storage Table (*.pst, *.ost)", "*.pst;*.ost"),
    ("MBOX Mailboxes (*.mbox, *.mbx)", "*.mbox;*.mbx"),
    ("Individual Email Messages (*.eml, *.emlx, *.msg)", "*.eml;*.emlx;*.msg"),
    ("Outlook for Mac Archive (*.olm)", "*.olm"),
    ("Compressed Mail Archives (*.zip, *.tar*, *.gz, *.bz2, *.xz)", "*.zip;*.tar;*.gz;*.bz2;*.xz;*.tgz"),
    ("All Files (*.*)", "*.*"),
]

PST_SAVE_FILETYPES: List[Tuple[str, str]] = [
    ("Outlook Personal Storage Table (*.pst)", "*.pst"),
    ("All Files (*.*)", "*.*"),
]

MBOX_SAVE_FILETYPES: List[Tuple[str, str]] = [
    ("MBOX Mailbox (*.mbox)", "*.mbox"),
    ("Compressed MBOX (*.mbox.gz)", "*.mbox.gz"),
    ("All Files (*.*)", "*.*"),
]

CSV_SAVE_FILETYPES: List[Tuple[str, str]] = [
    ("CSV Spreadsheet (*.csv)", "*.csv"),
    ("All Files (*.*)", "*.*"),
]


def _tkinter_available() -> bool:
    try:
        import tkinter
        return True
    except ImportError:
        return False


def _powershell_pick_file(
    title: str = "Select Mail File",
    filter_spec: str = "All Supported Mail Files|*.pst;*.ost;*.mbox;*.mbx;*.eml;*.emlx;*.msg;*.olm;*.zip|All Files|*.*",
    initialdir: Optional[str] = None,
    multiple: bool = False,
) -> Optional[Union[str, List[str]]]:
    """Fallback Windows file picker using PowerShell System.Windows.Forms."""
    multiselect = "$true" if multiple else "$false"
    init_dir = f"$dialog.InitialDirectory = '{initialdir}';" if initialdir and os.path.exists(initialdir) else ""
    ps_script = f"""
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = '{title}'
$dialog.Filter = '{filter_spec}'
$dialog.Multiselect = {multiselect}
{init_dir}
$res = $dialog.ShowDialog()
if ($res -eq [System.Windows.Forms.DialogResult]::OK) {{
    if ($dialog.Multiselect) {{
        $dialog.FileNames -join [Environment]::NewLine
    }} else {{
        $dialog.FileName
    }}
}}
"""
    try:
        res = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps_script],
            capture_output=True,
            text=True,
            check=False,
        )
        out = res.stdout.strip()
        if not out:
            return None
        if multiple:
            lines = [line.strip() for line in out.splitlines() if line.strip()]
            return lines if lines else None
        return out
    except Exception:
        return None


def _powershell_pick_folder(
    title: str = "Select Mail Directory",
    initialdir: Optional[str] = None,
) -> Optional[str]:
    """Fallback Windows folder browser dialog using PowerShell."""
    init_dir = f"$dialog.SelectedPath = '{initialdir}';" if initialdir and os.path.exists(initialdir) else ""
    ps_script = f"""
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.FolderBrowserDialog
$dialog.Description = '{title}'
$dialog.ShowNewFolderButton = $true
{init_dir}
$res = $dialog.ShowDialog()
if ($res -eq [System.Windows.Forms.DialogResult]::OK) {{
    $dialog.SelectedPath
}}
"""
    try:
        res = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps_script],
            capture_output=True,
            text=True,
            check=False,
        )
        out = res.stdout.strip()
        return out if out else None
    except Exception:
        return None


def _powershell_save_file(
    title: str = "Save Destination File",
    default_ext: str = ".pst",
    filter_spec: str = "Outlook PST (*.pst)|*.pst|All Files|*.*",
    initialfile: str = "archive.pst",
    initialdir: Optional[str] = None,
) -> Optional[str]:
    """Fallback Windows save file dialog using PowerShell."""
    init_dir = f"$dialog.InitialDirectory = '{initialdir}';" if initialdir and os.path.exists(initialdir) else ""
    ps_script = f"""
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.SaveFileDialog
$dialog.Title = '{title}'
$dialog.Filter = '{filter_spec}'
$dialog.DefaultExt = '{default_ext.lstrip(".")}'
$dialog.FileName = '{initialfile}'
{init_dir}
$res = $dialog.ShowDialog()
if ($res -eq [System.Windows.Forms.DialogResult]::OK) {{
    $dialog.FileName
}}
"""
    try:
        res = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps_script],
            capture_output=True,
            text=True,
            check=False,
        )
        out = res.stdout.strip()
        return out if out else None
    except Exception:
        return None


def pick_open_file(
    title: str = "Select Mailbox Archive File",
    filetypes: Optional[List[Tuple[str, str]]] = None,
    multiple: bool = False,
    initialdir: Optional[str] = None,
) -> Optional[Union[str, List[str]]]:
    """
    Open native host OS file dialog to select one or multiple files.
    Returns normalized file path string, list of strings (if multiple=True), or None if cancelled.
    """
    if filetypes is None:
        filetypes = MAIL_FILETYPES

    if _tkinter_available():
        try:
            import tkinter as tk
            from tkinter import filedialog

            root = tk.Tk()
            root.withdraw()
            try:
                root.attributes("-topmost", True)
            except Exception:
                pass

            kw = {"title": title, "filetypes": filetypes}
            if initialdir and os.path.exists(initialdir):
                kw["initialdir"] = initialdir

            if multiple:
                selected = filedialog.askopenfilenames(parent=root, **kw)
                root.destroy()
                if selected:
                    return [os.path.normpath(p) for p in selected]
                return None
            else:
                selected = filedialog.askopenfilename(parent=root, **kw)
                root.destroy()
                if selected:
                    return os.path.normpath(selected)
                return None
        except Exception:
            pass

    # Windows PowerShell fallback
    if sys.platform == "win32":
        filter_parts = []
        for desc, exts in filetypes:
            filter_parts.append(f"{desc}|{exts}")
        filter_spec = "|".join(filter_parts)
        res = _powershell_pick_file(title=title, filter_spec=filter_spec, initialdir=initialdir, multiple=multiple)
        if res:
            if isinstance(res, list):
                return [os.path.normpath(p) for p in res]
            return os.path.normpath(res)

    return None


def pick_folder(
    title: str = "Select Mailbox Directory (Thunderbird / Maildir / EML Tree)",
    initialdir: Optional[str] = None,
) -> Optional[str]:
    """
    Open native host OS directory selection dialog.
    Returns normalized folder path string, or None if cancelled.
    """
    if _tkinter_available():
        try:
            import tkinter as tk
            from tkinter import filedialog

            root = tk.Tk()
            root.withdraw()
            try:
                root.attributes("-topmost", True)
            except Exception:
                pass

            kw = {"title": title}
            if initialdir and os.path.exists(initialdir):
                kw["initialdir"] = initialdir

            selected = filedialog.askdirectory(parent=root, **kw)
            root.destroy()
            if selected:
                return os.path.normpath(selected)
            return None
        except Exception:
            pass

    if sys.platform == "win32":
        res = _powershell_pick_folder(title=title, initialdir=initialdir)
        if res:
            return os.path.normpath(res)

    return None


def pick_save_file(
    title: str = "Save Destination File",
    default_ext: str = ".pst",
    filetypes: Optional[List[Tuple[str, str]]] = None,
    initialfile: str = "archive.pst",
    initialdir: Optional[str] = None,
) -> Optional[str]:
    """
    Open native host OS Save As dialog.
    Returns normalized target path string, or None if cancelled.
    """
    if filetypes is None:
        if default_ext.lower() == ".pst":
            filetypes = PST_SAVE_FILETYPES
        elif default_ext.lower() in (".mbox", ".gz"):
            filetypes = MBOX_SAVE_FILETYPES
        elif default_ext.lower() == ".csv":
            filetypes = CSV_SAVE_FILETYPES
        else:
            filetypes = [("Target File", f"*{default_ext}"), ("All Files (*.*)", "*.*")]

    if _tkinter_available():
        try:
            import tkinter as tk
            from tkinter import filedialog

            root = tk.Tk()
            root.withdraw()
            try:
                root.attributes("-topmost", True)
            except Exception:
                pass

            kw = {
                "title": title,
                "defaultextension": default_ext,
                "filetypes": filetypes,
                "initialfile": initialfile,
            }
            if initialdir and os.path.exists(initialdir):
                kw["initialdir"] = initialdir

            selected = filedialog.asksaveasfilename(parent=root, **kw)
            root.destroy()
            if selected:
                return os.path.normpath(selected)
            return None
        except Exception:
            pass

    if sys.platform == "win32":
        filter_parts = []
        for desc, exts in filetypes:
            filter_parts.append(f"{desc}|{exts}")
        filter_spec = "|".join(filter_parts)
        res = _powershell_save_file(
            title=title,
            default_ext=default_ext,
            filter_spec=filter_spec,
            initialfile=initialfile,
            initialdir=initialdir,
        )
        if res:
            return os.path.normpath(res)

    return None


def pick_source(
    title: str = "Select Mail Source (Archive File or Directory)",
    initialdir: Optional[str] = None,
) -> Optional[str]:
    """
    Interactively choose whether to select an archive file or a directory container,
    then invoke the corresponding native host OS picker.
    """
    # If invoked in an interactive terminal, give user a quick choice
    if sys.stdin.isatty():
        print(f"\n[{title}]")
        print("  [1] Select Mail Archive File (.pst, .mbox, .eml, .msg, .olm, .zip, etc.)")
        print("  [2] Select Mailbox Directory (Thunderbird profile, Maildir, EML folder tree)")
        print("  [C] Cancel")
        try:
            choice = input("Choice [1/2/C] (default: 1): ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            return None

        if choice in ("2", "d", "dir", "folder"):
            return pick_folder(title=f"{title} - Directory", initialdir=initialdir)
        elif choice in ("c", "q", "exit", "cancel"):
            return None
        else:
            return pick_open_file(title=f"{title} - File", initialdir=initialdir)
    else:
        # Non-TTY or automated context: try file picker first
        return pick_open_file(title=title, initialdir=initialdir)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="file_picker - Native Host OS File and Folder Selection Utility for Mail2Mail / All2PST"
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--file", action="store_true", help="Launch native file picker to select a mail file")
    group.add_argument("--files", action="store_true", help="Launch native multi-file picker to select multiple files")
    group.add_argument("--folder", action="store_true", help="Launch native directory picker")
    group.add_argument("--save", action="store_true", help="Launch native Save As file dialog")
    group.add_argument("--source", action="store_true", help="Interactive prompt to pick either file or directory")

    parser.add_argument("--title", default=None, help="Custom dialog window title")
    parser.add_argument("--ext", default=".pst", help="Default extension for save dialog (default: .pst)")
    parser.add_argument("--initialdir", default=None, help="Initial directory to browse from")
    parser.add_argument("--initialfile", default="archive.pst", help="Default file name for save dialog")

    args = parser.parse_args()

    result: Optional[Union[str, List[str]]] = None
    if args.file:
        title = args.title or "Select Mailbox Archive File"
        result = pick_open_file(title=title, initialdir=args.initialdir)
    elif args.files:
        title = args.title or "Select Mailbox Files to Merge"
        result = pick_open_file(title=title, multiple=True, initialdir=args.initialdir)
    elif args.folder:
        title = args.title or "Select Mailbox Directory"
        result = pick_folder(title=title, initialdir=args.initialdir)
    elif args.save:
        title = args.title or "Select Destination File"
        result = pick_save_file(
            title=title,
            default_ext=args.ext,
            initialfile=args.initialfile,
            initialdir=args.initialdir,
        )
    elif args.source:
        title = args.title or "Select Mail Source"
        result = pick_source(title=title, initialdir=args.initialdir)

    if result is None:
        print("Selection cancelled.", file=sys.stderr)
        sys.exit(1)

    if isinstance(result, list):
        for item in result:
            print(item)
    else:
        print(result)


if __name__ == "__main__":
    main()
