import os
import sys
import gc
import shutil
import mailbox
import tempfile
import email
from email import policy
from email.utils import parsedate_to_datetime, getaddresses
import win32com.client

# Outlook MAPI invalid characters: \ / : * ? " < > |
FORBIDDEN_CHARS = r'\/:*?"<>|'

# MAPI Property Tags
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACH_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x37140003"
PR_SENDER_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F"
PR_SENDER_EMAIL = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F"
PR_SENT_REPRESENTING_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0042001F"
PR_SENT_REPRESENTING_EMAIL = "http://schemas.microsoft.com/mapi/proptag/0x0044001F"

def sanitize_folder_name(name):
    """Sanitizes folder names to comply with Outlook MAPI naming rules."""
    clean = "".join('_' if c in FORBIDDEN_CHARS else c for c in name).strip()
    return clean or "Unnamed_Folder"

def get_or_create_subfolder(parent_folder, folder_name):
    """Retrieves an existing MAPI subfolder or creates it if not found."""
    safe_name = sanitize_folder_name(folder_name)
    try:
        return parent_folder.Folders[safe_name]
    except Exception:
        try:
            return parent_folder.Folders.Add(safe_name)
        except Exception as e:
            print(f"    [!] Could not create subfolder '{safe_name}': {e}")
            return parent_folder

def is_probable_mbox(file_path):
    """Checks if a file is an mbox file based on extension and magic bytes."""
    basename = os.path.basename(file_path)
    ignored_exts = ('.msf', '.dat', '.json', '.html', '.txt', '.ini', '.bak', '.tmp')
    if basename.startswith('.') or basename.lower().endswith(ignored_exts):
        return False

    if os.path.getsize(file_path) == 0:
        return True

    try:
        with open(file_path, 'rb') as f:
            first_bytes = f.read(5)
            return first_bytes.startswith(b"From ") or first_bytes.startswith(b"\nFrom")
    except Exception:
        return False

def add_recipients(mail_item, parsed_msg):
    """Adds To, CC, and BCC recipients to the Outlook MailItem."""
    recipient_types = {
        'to': 1,   # olTo
        'cc': 2,   # olCC
        'bcc': 3   # olBCC
    }

    for header_name, ol_type in recipient_types.items():
        headers = parsed_msg.get_all(header_name, [])
        if not headers:
            continue
        
        # Parse name and email tuples
        addresses = getaddresses(headers)
        for realname, addr in addresses:
            if not addr:
                continue
            try:
                recip = mail_item.Recipients.Add(f"{realname} <{addr}>" if realname else addr)
                recip.Type = ol_type
            except Exception:
                pass

def handle_sender_info(mail_item, parsed_msg):
    """Maps From and Reply-To headers using MAPI PropertyAccessor."""
    from_headers = parsed_msg.get_all('from', [])
    if from_headers:
        senders = getaddresses(from_headers)
        if senders:
            realname, addr = senders[0]
            try:
                pa = mail_item.PropertyAccessor
                if realname:
                    pa.SetProperty(PR_SENDER_NAME, realname)
                    pa.SetProperty(PR_SENT_REPRESENTING_NAME, realname)
                if addr:
                    pa.SetProperty(PR_SENDER_EMAIL, addr)
                    pa.SetProperty(PR_SENT_REPRESENTING_EMAIL, addr)
                mail_item.SentOnBehalfOfName = realname or addr
            except Exception:
                mail_item.SentOnBehalfOfName = addr or realname

def populate_mailitem_fallback(mail_item, msg_bytes, temp_dir):
    """
    Parses RFC822 MIME bytes directly into a native Outlook MailItem.
    Handles HTML/Text body, inline CID images, regular attachments, and S/MIME.
    """
    parsed = email.message_from_bytes(msg_bytes, policy=policy.default)
    
    # 1. Subject
    mail_item.Subject = str(parsed.get('subject', ''))
    
    # 2. Sent Date
    try:
        date_hdr = parsed.get('date')
        if date_hdr:
            dt = parsedate_to_datetime(date_hdr)
            mail_item.SentOn = dt
    except Exception:
        pass

    # 3. Sender and Recipients
    handle_sender_info(mail_item, parsed)
    add_recipients(mail_item, parsed)

    # 4. S/MIME Detection
    content_type = parsed.get_content_type().lower()
    is_smime = False
    if "pkcs7-mime" in content_type or "signed-data" in content_type:
        mail_item.MessageClass = "IPM.Note.SMIME"
        is_smime = True
    elif "multipart/signed" in content_type:
        mail_item.MessageClass = "IPM.Note.SMIME.MultipartSigned"
        is_smime = True

    # 5. Body Extraction
    if not is_smime:
        body_part = parsed.get_body(preferencelist=('html', 'plain'))
        if body_part:
            try:
                content = body_part.get_content()
                if body_part.get_content_type() == 'text/html':
                    mail_item.HTMLBody = content
                else:
                    mail_item.Body = content
            except Exception:
                mail_item.Body = str(body_part.get_payload())
        else:
            payload = parsed.get_payload()
            if isinstance(payload, str):
                mail_item.Body = payload

    # 6. Extract Attachments and Inline CID Assets
    for att_idx, part in enumerate(parsed.walk()):
        # Skip top-level multipart containers
        if part.is_multipart():
            continue

        cdisp = str(part.get("Content-Disposition", "")).lower()
        cid = part.get("Content-ID", None)
        filename = part.get_filename()

        # Identify whether this part is an attachment or inline asset
        is_attachment = ("attachment" in cdisp) or (filename is not None)
        is_inline_cid = (cid is not None) or ("inline" in cdisp and not is_smime)

        if not (is_attachment or is_inline_cid or is_smime):
            # Normal text/html body parts are already handled
            continue

        if not filename:
            if is_smime and ("pkcs7" in content_type or "pkcs7" in str(part.get_content_type()).lower()):
                filename = "smime.p7m"
            else:
                ext = part.get_content_subtype() or "bin"
                filename = f"attachment_{att_idx}.{ext}"

        safe_filename = "".join(c for c in filename if c not in FORBIDDEN_CHARS).strip() or f"file_{att_idx}.bin"
        temp_att_path = os.path.realpath(os.path.join(temp_dir, f"att_{att_idx}_{safe_filename}"))

        try:
            att_data = part.get_content()
            att_bytes = att_data.encode('utf-8', errors='replace') if isinstance(att_data, str) else bytes(att_data)

            with open(temp_att_path, "wb") as f:
                f.write(att_bytes)
                f.flush()
                os.fsync(f.fileno())

            # Type 1 = olByValue (embed copy directly in PST)
            att = mail_item.Attachments.Add(temp_att_path, 1, 1, safe_filename)

            # Map Content-ID for inline images
            if cid:
                clean_cid = cid.strip("<>")
                try:
                    att_pa = att.PropertyAccessor
                    att_pa.SetProperty(PR_ATTACH_CONTENT_ID, clean_cid)
                    # 0x00000004 = ATT_MIME_INLINE (hides attachment strip in Outlook and renders inline)
                    att_pa.SetProperty(PR_ATTACH_FLAGS, 4)
                except Exception:
                    pass

            del att

        except Exception as att_err:
            print(f"        [!] Warning: Failed to process attachment {filename}: {att_err}")
        finally:
            try:
                if os.path.exists(temp_att_path):
                    os.remove(temp_att_path)
            except OSError:
                pass

    # 7. Commit changes to PST store
    mail_item.Save()

def import_mbox_to_outlook(mbox_path, outlook_folder, temp_dir):
    """Processes an mbox file and adds messages directly into the destination Outlook folder."""
    if os.path.getsize(mbox_path) == 0:
        return

    print(f"[*] Processing: {os.path.basename(mbox_path)} -> [{outlook_folder.Name}]")
    
    try:
        mbox = mailbox.mbox(mbox_path, create=False)
    except Exception as e:
        print(f"    [!] Error opening mbox {mbox_path}: {e}")
        return

    outlook_app = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook_app.GetNamespace("MAPI")
    count = 0

    for idx, msg in enumerate(mbox):
        try:
            raw_bytes = msg.as_bytes()
            if not raw_bytes:
                continue

            temp_eml = os.path.realpath(os.path.join(temp_dir, f"conv_{idx}.eml"))
            
            with open(temp_eml, "wb") as f:
                f.write(raw_bytes)
                f.flush()
                os.fsync(f.fileno())

            item = None
            try:
                # OpenSharedItem natively resolves pre-compiled MAPI items
                item = namespace.OpenSharedItem(temp_eml)
                item.Move(outlook_folder)
            except Exception:
                # Fallback to direct MailItem property injection
                new_item = outlook_folder.Items.Add(0)  # 0 = olMailItem
                populate_mailitem_fallback(new_item, raw_bytes, temp_dir)
                del new_item

            count += 1
            if item:
                del item

            try:
                if os.path.exists(temp_eml):
                    os.remove(temp_eml)
            except PermissionError:
                pass

            if count % 100 == 0:
                print(f"    -> {count} items imported...")
                gc.collect()

        except Exception as err:
            print(f"    [!] Skipped message #{idx}: {err}")
            continue

    print(f"    [+] Finished: {count} messages imported into '{outlook_folder.Name}'.")

def traverse_and_convert(current_dir, current_outlook_folder, temp_dir):
    """Recursively walks through Thunderbird mbox files and .sbd directories."""
    try:
        entries = os.listdir(current_dir)
    except Exception as e:
        print(f"[!] Access error reading {current_dir}: {e}")
        return

    # Process MBOX files
    for entry in entries:
        full_path = os.path.join(current_dir, entry)
        if os.path.isfile(full_path) and is_probable_mbox(full_path):
            target_folder = get_or_create_subfolder(current_outlook_folder, entry)
            import_mbox_to_outlook(full_path, target_folder, temp_dir)

    # Process .sbd container folders
    for entry in entries:
        full_path = os.path.join(current_dir, entry)
        if os.path.isdir(full_path) and entry.lower().endswith(".sbd"):
            parent_folder_name = entry[:-4]
            nested_outlook_folder = get_or_create_subfolder(current_outlook_folder, parent_folder_name)
            traverse_and_convert(full_path, nested_outlook_folder, temp_dir)

def main():
    if len(sys.argv) < 3:
        print("Usage: python imapmail_to_pst.py <path_to_ImapMail_account_dir> <output_pst_path>")
        print(r'Example: python imapmail_to_pst.py "C:\Users\<User>\AppData\Roaming\Thunderbird\Profiles\xxx.default\ImapMail\imap.example.com" "C:\Backups\Mailbox.pst"')
        sys.exit(1)

    imap_src = os.path.abspath(sys.argv[1])
    pst_dest = os.path.abspath(sys.argv[2])

    if not os.path.isdir(imap_src):
        print(f"[!] Source path does not exist or is not a directory: {imap_src}")
        sys.exit(1)

    os.makedirs(os.path.dirname(pst_dest), exist_ok=True)

    print("[*] Initializing Outlook MAPI Interface...")
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook_app.GetNamespace("MAPI")
    namespace.Logon("", "", False, False)

    # Identify existing stores to isolate the new target
    existing_store_ids = set()
    for s in namespace.Stores:
        try:
            existing_store_ids.add(s.StoreID)
        except Exception:
            pass

    print(f"[*] Attaching target PST: {pst_dest}")
    try:
        namespace.AddStoreEx(pst_dest, 1)  # 1 = olStoreUnicode
    except Exception:
        namespace.AddStore(pst_dest)

    target_store = None
    for s in namespace.Stores:
        try:
            if s.StoreID not in existing_store_ids:
                target_store = s
                break
        except Exception:
            continue

    if not target_store:
        for s in namespace.Stores:
            try:
                if s.FilePath and os.path.abspath(s.FilePath).lower() == pst_dest.lower():
                    target_store = s
                    break
            except Exception:
                continue

    if not target_store:
        print("[!] Error: Could not locate the newly attached PST store.")
        sys.exit(1)

    pst_root = target_store.GetRootFolder()
    print(f"[+] PST Root opened: {pst_root.Name}")

    temp_staging = tempfile.mkdtemp(prefix="tb_pst_conv_")

    try:
        traverse_and_convert(imap_src, pst_root, temp_staging)
    finally:
        print("[*] Cleaning temporary staging directory...")
        shutil.rmtree(temp_staging, ignore_errors=True)

        print("[*] Detaching PST store to flush buffers...")
        try:
            namespace.RemoveStore(pst_root)
        except Exception as e:
            print(f"[!] Notice during PST unmount: {e}")
        
        print("[+] Conversion completed successfully.")

if __name__ == "__main__":
    main()