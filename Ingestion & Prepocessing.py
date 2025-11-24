import zipfile
import os
import email
from email import policy
from email.parser import BytesParser
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
import pandas as pd

# =========================
# 0. FILE UPLOAD DIALOG
# =========================
def pick_zip_file():
    Tk().withdraw()
    print("Please select your ZIP file containing emails...")
    file_path = askopenfilename(
        title="Select email ZIP file",
        filetypes=[("ZIP files", "*.zip")]
    )
    if not file_path:
        raise ValueError("No file selected.")
    return file_path

# =========================
# 1. UNZIP AND LOAD EMAILS
# =========================
def extract_eml_files(zip_path, extract_dir="unzipped_emails"):
    if not os.path.exists(extract_dir):
        os.makedirs(extract_dir)
    with zipfile.ZipFile(zip_path, 'r') as z:
        z.extractall(extract_dir)
    eml_files = []
    for root, dirs, files in os.walk(extract_dir):
        for file in files:
            if file.lower().endswith(".eml"):
                eml_files.append(os.path.join(root, file))
    return eml_files

# =========================
# 2. PARSE SINGLE EMAIL
# =========================
def parse_eml(path):
    with open(path, "rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)
    message_id = msg.get("Message-ID")
    in_reply_to = msg.get("In-Reply-To")
    references = msg.get_all("References", [])
    if references:
        references = re.split(r"\s+", " ".join(references))
        references = [ref.strip("<>") for ref in references]
    else:
        references = []
    body = extract_email_body(msg)
    body_clean = clean_email_body(body)
    return {
        "message_id": message_id,
        "subject": msg.get("Subject"),
        "from": msg.get("From"),
        "to": msg.get_all("To", []),
        "date": msg.get("Date"),
        "body_raw": body,
        "body_clean": body_clean,
        "in_reply_to": in_reply_to,
        "references": references
    }

# =========================
# 3. BODY EXTRACTION
# =========================
def extract_email_body(msg):
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                return part.get_content().strip()
        return ""
    else:
        return msg.get_content().strip()

# =========================
# 4. CLEAN EMAIL BODY
# =========================
def clean_email_body(body):
    if not body:
        return ""
    # Remove quoted content
    body = "\n".join(line for line in body.split("\n") if not line.strip().startswith(">"))
    # Remove signatures, disclaimers, greetings
    body = re.split(r"(--|Sent from my iPhone|Best regards|Kind regards|Regards|Thanks|Thank you)", body, flags=re.IGNORECASE)[0]
    # Lowercase
    body = body.lower()
    # Remove URLs
    body = re.sub(r"http\S+", "", body)
    # Remove punctuation
    body = re.sub(r"[^a-z0-9\s]", " ", body)
    # Normalize whitespace
    body = re.sub(r"\s+", " ", body).strip()
    return body

# =========================
# 5. LOAD ALL EMAILS
# =========================
def load_all_emails(eml_paths):
    parsed = {}
    for path in eml_paths:
        data = parse_eml(path)
        if data["message_id"]:
            parsed[data["message_id"]] = data
    return parsed

# =========================
# 6. BUILD THREAD STRUCTURE
# =========================
def build_thread_structure(parsed_emails):
    children_map = {}
    for msg_id, data in parsed_emails.items():
        parent = None
        if data['in_reply_to'] and data['in_reply_to'] in parsed_emails:
            parent = data['in_reply_to']
        elif data['references']:
            for ref in reversed(data['references']):
                if ref in parsed_emails:
                    parent = ref
                    break
        if parent:
            children_map.setdefault(parent, []).append(msg_id)
    for parent, kids in children_map.items():
        kids.sort(key=lambda k: parsed_emails[k]['date'] if parsed_emails[k]['date'] else "")
    return children_map

# =========================
# 7. FLATTEN THREADS TO CONVERSATION ORDER
# =========================
def flatten_threads(parsed_emails, children_map):
    all_children = {child for kids in children_map.values() for child in kids}
    roots = [msg_id for msg_id in parsed_emails if msg_id not in all_children]
    flat_list = []
    def add_email_with_replies(msg_id, level=0):
        data = parsed_emails[msg_id].copy()
        data['level'] = level
        flat_list.append(data)
        for child_id in children_map.get(msg_id, []):
            add_email_with_replies(child_id, level + 1)
    for root in roots:
        add_email_with_replies(root)
    return flat_list

# =========================
# 8. EXPORT TO CSV
# =========================
def export_conversation_csv(flat_list):
    Tk().withdraw()
    save_path = asksaveasfilename(
        title="Save threaded conversation CSV",
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv")]
    )
    if not save_path:
        print("[WARNING] CSV export canceled by user.")
        return
    df = pd.DataFrame(flat_list)
    df['subject'] = df.apply(lambda x: '    ' * x['level'] + str(x['subject']), axis=1)
    df['body_snippet'] = df['body_clean'].apply(lambda x: x[:100] + "..." if len(x) > 100 else x)
    df.drop(columns=['level'], inplace=True)
    df.to_csv(save_path, index=False)
    print(f"[INFO] Exported threaded conversation CSV to {save_path}")

# =========================
# 9. MAIN PIPELINE
# =========================
def process_email_zip():
    zip_path = pick_zip_file()
    print("[INFO] Extracting emails...")
    eml_files = extract_eml_files(zip_path)
    print(f"[INFO] Found {len(eml_files)} .eml files.")
    print("[INFO] Parsing emails...")
    parsed_emails = load_all_emails(eml_files)
    print(f"[INFO] Successfully parsed {len(parsed_emails)} emails.")
    print("[INFO] Building thread structure...")
    children_map = build_thread_structure(parsed_emails)
    print("[INFO] Flattening threads to conversation order...")
    flat_list = flatten_threads(parsed_emails, children_map)
    print("[INFO] Exporting threaded conversation CSV...")
    export_conversation_csv(flat_list)
    print("[INFO] Pipeline completed successfully.")
    return parsed_emails

# =========================
# 10. RUN SCRIPT
# =========================
if __name__ == "__main__":
    parsed = process_email_zip()
