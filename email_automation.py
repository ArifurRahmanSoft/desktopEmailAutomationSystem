import csv
import html
import functools
import json
import os
import random
import re
import shutil
import smtplib
import subprocess
import sys
import threading
import time
import uuid
from urllib.parse import quote
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from dotenv import dotenv_values
from openpyxl import Workbook, load_workbook
from account_service import AccountService
from attachment_library_service import AttachmentLibraryService
from bounce_tracking_service import BounceTrackingService
from build_info import BUILD_TIMESTAMP_UTC
from placeholder_service import PlaceholderService
from tracking_sync_service import TrackingSynchronizationService, synchronization_is_due


APP_NAME = "Email Automation"
TASK_NAME = "The Power People Daily Email Automation"
DEFAULT_DATA_DIR = r"F:\CODEX\Email_automation"
DEFAULT_CONFIG = {"daily_limit": 5, "schedule_time": "09:00", "data_dir": DEFAULT_DATA_DIR, "default_sender_name": "The Power People", "random_delay_min": 5, "random_delay_max": 15, "retry_count": 1, "backup_enabled": True, "log_retention_days": 90, "theme": "Light", "tracking_sync_enabled": False, "tracking_sync_interval_hours": 5, "tracking_last_sync_time": "", "FullSynchronizationDebug": True, "bounce_check_enabled": False, "bounce_check_interval_minutes": 30}
TRACKING_BASE_URL = "https://emailtrackingserver.onrender.com"


def exe_dir():
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent


def install_dir():
    return Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "ThePowerPeople" / "EmailAutomation"


def config_path():
    return install_dir() / "config" / "settings.json"


def load_config():
    p = config_path()
    if not p.exists():
        return dict(DEFAULT_CONFIG)
    try:
        return {**DEFAULT_CONFIG, **json.loads(p.read_text(encoding="utf-8"))}
    except Exception:
        return dict(DEFAULT_CONFIG)


def save_config(cfg):
    config_path().parent.mkdir(parents=True, exist_ok=True)
    config_path().write_text(json.dumps(cfg, indent=2), encoding="utf-8")


def app_paths(cfg=None):
    cfg = cfg or load_config()
    data = Path(cfg["data_dir"])
    root = install_dir()
    return {
        "root": root,
        "data": data,
        "list": data / "mail_list.xlsx",
        "env": root / ".env",
        "logs": root / "logs",
        "backup": root / "backup",
        "reports": root / "reports",
        "accounts": root / "config" / "accounts.json",
        "schedules": root / "config" / "schedules.json",
    }


def ensure_dirs():
    p = app_paths()
    for key in ("root", "logs", "backup", "reports"):
        p[key].mkdir(parents=True, exist_ok=True)


def daily_log(email="", subject="", status="", error="", sender_email="", smtp_response="", execution_time="", tracking_id=""):
    ensure_dirs()
    now = datetime.now()
    path = app_paths()["logs"] / f"{now:%Y-%m-%d}.log"
    new_file = not path.exists()
    with path.open("a", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        if new_file:
            writer.writerow(["Date", "Time", "TrackingId", "Sender Email", "Recipient Email", "Subject", "Status", "SMTP Response", "Error Message", "Execution Time Seconds"])
        writer.writerow([f"{now:%Y-%m-%d}", f"{now:%H:%M:%S}", tracking_id, sender_email, email, subject, status, smtp_response, error, execution_time])


def startup_log():
    ensure_dirs()
    path = app_paths()["logs"] / "startup.log"
    with path.open("a", encoding="utf-8") as handle:
        handle.write(f"{datetime.now(timezone.utc).isoformat()} | Build: {BUILD_TIMESTAMP_UTC} UTC\n")


def account_service():
    return AccountService(app_paths()["accounts"])


def tracking_sync_service():
    return TrackingSynchronizationService(TRACKING_BASE_URL, app_paths()["list"], app_paths()["logs"])


def bounce_tracking_service():
    return BounceTrackingService(app_paths()["list"], app_paths()["logs"])


def attachment_library_service():
    return AttachmentLibraryService(TRACKING_BASE_URL, app_paths()["logs"])


def resolve_selected_attachments(attachment_ids):
    selected_ids = [str(value) for value in (attachment_ids or [])]
    if not selected_ids:
        return []
    available = {item["id"]: item for item in attachment_library_service().list_attachments()}
    missing = [value for value in selected_ids if value not in available]
    if missing:
        raise ValueError("One or more selected attachments have been deleted from the server. Refresh the attachment list and try again.")
    return [available[value] for value in selected_ids]


def build_attachment_links_html(tracking_id, attachments):
    if not attachments:
        return "", []
    rows = ["<p><strong>Attachments</strong></p>"]
    urls = []
    for attachment in attachments:
        url = f"{TRACKING_BASE_URL}/download/{tracking_id}/{quote(str(attachment['id']), safe='')}"
        urls.append(url)
        rows.append(f'<p>📄 <a href="{html.escape(url, quote=True)}">{html.escape(attachment["file_name"])}</a></p>')
    return "".join(rows), urls


def register_attachment_mapping(tracking_id, attachments):
    if not attachments:
        return
    attachment_library_service().register_tracking_attachments(tracking_id, [item["id"] for item in attachments])


def attachment_send_log(tracking_id, recipient, attachments, urls):
    if not attachments:
        return
    ensure_dirs(); path = app_paths()["logs"] / f"attachment-send-{datetime.now():%Y-%m-%d}.log"
    entry = {"timestamp": datetime.now().isoformat(), "tracking_id": tracking_id, "recipient": recipient, "attachment_ids": [item["id"] for item in attachments], "original_file_names": [item["file_name"] for item in attachments], "download_urls": urls}
    with path.open("a", encoding="utf-8") as handle: handle.write(json.dumps(entry, ensure_ascii=False) + "\n")


def migrate_legacy_account():
    service = account_service()
    env_path = app_paths()["env"]
    if service.list_accounts():
        if env_path.exists():
            try: env_path.unlink()
            except Exception: pass
        return
    try:
        address, password = credentials()
        service.save_account("Primary Gmail", address, password, True)
        env_path.unlink(missing_ok=True)
    except Exception:
        pass


def credentials():
    values = dotenv_values(app_paths()["env"])
    address = (values.get("EMAIL_ADDRESS") or "").strip()
    password = (values.get("EMAIL_PASSWORD") or "").strip().replace(" ", "")
    if not address or not password:
        raise ValueError("EMAIL_ADDRESS or EMAIL_PASSWORD is missing from .env")
    return address, password


def header_map(ws):
    return {str(cell.value).strip().lower(): i for i, cell in enumerate(ws[1], 1) if cell.value is not None}


@functools.lru_cache(maxsize=2048)
def domain_has_mx(domain):
    """Return True only when the recipient domain publishes a usable MX record."""
    try:
        result = subprocess.run(
            ["nslookup", "-type=MX", domain],
            capture_output=True,
            text=True,
            timeout=15,
            startupinfo=hidden_startupinfo(),
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    except Exception:
        return False
    output = f"{result.stdout}\n{result.stderr}"
    exchangers = re.findall(r"mail exchanger\s*=\s*([^\s]+)", output, flags=re.IGNORECASE)
    return any(value.rstrip(".") for value in exchangers)


def validate_recipient_email(email):
    """Validate common email syntax, DNS-safe domain syntax, and MX availability."""
    value = str(email or "").strip()
    if len(value) > 254 or value.count("@") != 1:
        return False
    local, domain = value.rsplit("@", 1)
    if not local or len(local) > 64 or not domain:
        return False
    if not re.fullmatch(r"[A-Za-z0-9!#$%&'*+/=?^_`{|}~.-]+", local):
        return False
    if local.startswith(".") or local.endswith(".") or ".." in local:
        return False
    try:
        ascii_domain = domain.rstrip(".").encode("idna").decode("ascii").lower()
    except UnicodeError:
        return False
    if not ascii_domain or len(ascii_domain) > 253 or "." not in ascii_domain:
        return False
    labels = ascii_domain.split(".")
    if any(not label or len(label) > 63 or label.startswith("-") or label.endswith("-") or not re.fullmatch(r"[a-z0-9-]+", label) for label in labels):
        return False
    return domain_has_mx(ascii_domain)


def validate_sheet(ws):
    h = header_map(ws)
    missing = [name for name in ("email", "subject", "body", "status") if name not in h]
    if missing:
        raise ValueError("Missing Excel columns: " + ", ".join(x.title() for x in missing))
    if "sentdate" not in h:
        col = ws.max_column + 1
        ws.cell(1, col, "SentDate")
        h["sentdate"] = col
    if "result" not in h:
        col = ws.max_column + 1
        ws.cell(1, col, "Result")
        h["result"] = col
    return h


def ensure_tracking_column(workbook, worksheet, workbook_path):
    """Return the TrackingId column, creating and immediately saving it if absent."""
    headers = header_map(worksheet)
    if "trackingid" in headers:
        return headers["trackingid"]
    column = worksheet.max_column + 1
    worksheet.cell(1, column, "TrackingId")
    workbook.save(workbook_path)
    return column


def write_tracking_id(workbook, worksheet, workbook_path, row, tracking_column, tracking_id):
    """Persist one row's TrackingId; SMTP must not run unless this returns successfully."""
    worksheet.cell(row, tracking_column, tracking_id)
    try:
        workbook.save(workbook_path)
    except Exception as exc:
        raise RuntimeError(f"Failed to write TrackingId to mail_list.xlsx: {exc}") from exc


def build_click_tracked_html(body, tracking_id):
    """Convert plain-text URLs to tracked HTML anchors while preserving visible text."""
    text = str(body or "")
    pattern = re.compile(r"https?://[^\s<>\"']+", re.IGNORECASE)
    output = []
    position = 0
    trailing_punctuation = ".,;:!?)]}"
    for match in pattern.finditer(text):
        output.append(html.escape(text[position:match.start()]))
        candidate = match.group(0)
        url = candidate.rstrip(trailing_punctuation)
        trailing = candidate[len(url):]
        if url:
            destination = f"https://emailtrackingserver.onrender.com/email/click/{tracking_id}?url={quote(url, safe='')}"
            output.append(f'<a href="{html.escape(destination, quote=True)}">{html.escape(url)}</a>')
        output.append(html.escape(trailing))
        position = match.end()
    output.append(html.escape(text[position:]))
    return "".join(output).replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>\n")


def create_backup(source):
    ensure_dirs()
    target = app_paths()["backup"] / f"mail_list_{datetime.now():%Y%m%d_%H%M%S_%f}.xlsx"
    shutil.copy2(source, target)
    return target


def pending_count():
    p = app_paths()["list"]
    if not p.exists():
        return 0
    wb = load_workbook(p, read_only=True, data_only=True)
    ws = wb.active
    h = header_map(ws)
    if "status" not in h:
        wb.close()
        return 0
    count = sum(1 for row in range(2, ws.max_row + 1) if str(ws.cell(row, h["status"]).value or "").strip().lower() == "pending")
    wb.close()
    return count


def email_grid_data(excel_path=None):
    p = Path(excel_path) if excel_path else app_paths()["list"]
    if not p.exists():
        return [], {"total": 0, "sent": 0, "pending": 0, "failed": 0}
    wb = load_workbook(p, read_only=True, data_only=True)
    ws = wb.active
    h = header_map(ws)
    rows = []
    counts = {"total": 0, "sent": 0, "pending": 0, "failed": 0}
    for row in range(2, ws.max_row + 1):
        status = str(ws.cell(row, h.get("status", 0)).value or "").strip()
        if h.get("name"):
            display_name = str(ws.cell(row, h["name"]).value or "")
        else:
            first = str(ws.cell(row, h["first_name"]).value or "") if h.get("first_name") else ""
            last = str(ws.cell(row, h["last_name"]).value or "") if h.get("last_name") else ""
            display_name = f"{first} {last}".strip()
        item = (
            display_name,
            str(ws.cell(row, h.get("email", 0)).value or "") if h.get("email") else "",
            status,
            str(ws.cell(row, h.get("sentdate", 0)).value or "") if h.get("sentdate") else "",
            str(ws.cell(row, h.get("result", 0)).value or "") if h.get("result") else "",
        )
        rows.append(item)
        counts["total"] += 1
        key = status.lower()
        if key in counts:
            counts[key] += 1
    wb.close()
    return rows, counts


def send_pending(limit, wait_between=True, progress=None, excel_path=None, attachment_ids=None):
    paths = app_paths()
    if excel_path:
        paths["list"] = Path(excel_path)
    if limit < 1:
        raise ValueError("Email count must be at least 1.")
    if not paths["list"].exists():
        raise FileNotFoundError(f"Excel file not found: {paths['list']}")
    selected_attachments = resolve_selected_attachments(attachment_ids)
    wb = load_workbook(paths["list"])
    ws = wb.active
    h = validate_sheet(ws)
    cfg = load_config()
    if cfg.get("backup_enabled", True):
        create_backup(paths["list"])
    wb.save(paths["list"])
    tracking_column = ensure_tracking_column(wb, ws, paths["list"])
    h["trackingid"] = tracking_column
    rows = [r for r in range(2, ws.max_row + 1) if str(ws.cell(r, h["status"]).value or "").strip().lower() == "pending"][:limit]
    sent = failed = 0
    headers = [cell.value for cell in ws[1]]
    default_sender_name = str(load_config().get("default_sender_name") or "The Power People").strip()
    accounts = account_service()
    smtp = None
    active_sender = None
    try:
      for pos, row in enumerate(rows):
        started = time.monotonic()
        email = str(ws.cell(row, h["email"]).value or "").strip()
        subject = ""
        error = ""
        sender_email = ""
        tracking_id = ""
        try:
            values = [ws.cell(row, column).value for column in range(1, ws.max_column + 1)]
            context = PlaceholderService.create_context(headers, values)
            subject = PlaceholderService.render(ws.cell(row, h["subject"]).value, context).strip()
            body = PlaceholderService.render(ws.cell(row, h["body"]).value, context).strip()
            sender_name = str(ws.cell(row, h["sender_name"]).value or "").strip() if h.get("sender_name") else ""
            sender_name = PlaceholderService.render(sender_name, context).strip() or default_sender_name
            sender_column = h.get("sender_email") or h.get("sender_mail")
            sender_email = str(ws.cell(row, sender_column).value or "").strip().lower() if sender_column else ""
            smtp_configuration = accounts.smtp_configuration(sender_email)
            if not smtp_configuration:
                raise ValueError("Sender Account Not Configured")
            if not validate_recipient_email(email):
                raise ValueError("Invalid email address")
            if not subject:
                raise ValueError("Subject is empty")
            if not body:
                raise ValueError("Body is empty")
            tracking_id = str(uuid.uuid4())
            write_tracking_id(wb, ws, paths["list"], row, tracking_column, tracking_id)
            register_attachment_mapping(tracking_id, selected_attachments)
            msg = EmailMessage()
            msg["From"] = f"{sender_name} <{sender_email}>"
            msg["To"] = email
            msg["Subject"] = subject
            msg.set_content(body)
            pixel_url = f"https://emailtrackingserver.onrender.com/email/open/{tracking_id}"
            html_body = build_click_tracked_html(body, tracking_id)
            attachment_html, download_urls = build_attachment_links_html(tracking_id, selected_attachments)
            email_html = f'{attachment_html}<hr><p><strong>Email Body</strong></p>{html_body}' if selected_attachments else html_body
            msg.add_alternative(f'<html><body>{email_html}<img src="{pixel_url}" width="1" height="1" style="display:none;" alt=""></body></html>', subtype="html")
            if active_sender != sender_email or smtp is None:
                if smtp is not None:
                    try: smtp.quit()
                    except Exception: pass
                smtp = accounts.connect_smtp(smtp_configuration, timeout=45)
                active_sender = sender_email
            attempts = max(1, int(cfg.get("retry_count", 1)) + 1)
            last_error = None
            for attempt in range(attempts):
                try:
                    smtp.send_message(msg)
                    last_error = None
                    break
                except Exception as exc:
                    last_error = exc
                    if attempt + 1 < attempts:
                        time.sleep(2)
            if last_error:
                raise last_error
            ws.cell(row, h["status"], "Sent")
            ws.cell(row, h["sentdate"], datetime.now())
            ws.cell(row, h["result"], "Success")
            attachment_send_log(tracking_id, email, selected_attachments, download_urls)
            sent += 1
            status = "Sent"
        except Exception as exc:
            error = str(exc)
            ws.cell(row, h["status"], "Failed")
            ws.cell(row, h["sentdate"], datetime.now())
            ws.cell(row, h["result"], error[:1000])
            failed += 1
            status = "Failed"
        wb.save(paths["list"])
        daily_log(email=email, subject=subject, status=status, error=error, sender_email=sender_email, smtp_response="Accepted" if status == "Sent" else "", execution_time=f"{time.monotonic()-started:.3f}", tracking_id=tracking_id)
        if progress:
            progress(pos + 1, len(rows), status, email)
        if wait_between and pos < len(rows) - 1:
            minimum = max(0, int(cfg.get("random_delay_min", 5)))
            maximum = max(minimum, int(cfg.get("random_delay_max", 15)))
            time.sleep(random.randint(minimum, maximum))
    finally:
      if smtp is not None:
        try: smtp.quit()
        except Exception: pass
    wb.close()
    return {"requested": limit, "sent": sent, "failed": failed, "remaining": pending_count()}


def hidden_startupinfo():
    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = subprocess.SW_HIDE
    return si


def run_cmd(args):
    return subprocess.run(args, capture_output=True, text=True, startupinfo=hidden_startupinfo(), creationflags=subprocess.CREATE_NO_WINDOW)


def task_query():
    result = run_cmd(["schtasks", "/Query", "/TN", TASK_NAME, "/FO", "LIST", "/V"])
    if result.returncode != 0:
        return {"exists": False, "enabled": False, "next": "Not scheduled"}
    text = result.stdout
    enabled = False
    next_run = "Unknown"
    for line in text.splitlines():
        if line.lower().startswith("scheduled task state"):
            enabled = line.split(":", 1)[1].strip().lower() == "enabled"
        if line.lower().startswith("next run time"):
            next_run = line.split(":", 1)[1].strip()
    return {"exists": True, "enabled": enabled, "next": next_run}


def scheduled_exe():
    return install_dir() / "SendPendingEmails.exe"


def update_task(cfg, enable=True):
    exe = scheduled_exe()
    if not exe.exists():
        raise FileNotFoundError(f"Installed sender not found: {exe}")
    command = f'"{exe}" --scheduled'
    result = run_cmd(["schtasks", "/Create", "/TN", TASK_NAME, "/TR", command, "/SC", "DAILY", "/ST", cfg["schedule_time"], "/F", "/RL", "LIMITED"])
    if result.returncode != 0:
        raise RuntimeError((result.stderr or result.stdout).strip())
    ps = (
        "$s=New-Object -ComObject 'Schedule.Service';$s.Connect();"
        f"$t=$s.GetFolder('\\').GetTask('{TASK_NAME.replace("'", "''")}');"
        "$d=$t.Definition;$d.Settings.StartWhenAvailable=$true;"
        "$s.GetFolder('\\').RegisterTaskDefinition($t.Name,$d,6,$null,$null,3,$null)|Out-Null"
    )
    run_cmd(["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps])
    set_task_enabled(enable)


def set_task_enabled(enabled):
    action = "/Enable" if enabled else "/Disable"
    result = run_cmd(["schtasks", "/Change", "/TN", TASK_NAME, action])
    if result.returncode != 0:
        raise RuntimeError((result.stderr or result.stdout).strip())


def delete_task():
    result = run_cmd(["schtasks", "/Delete", "/TN", TASK_NAME, "/F"])
    if result.returncode != 0:
        raise RuntimeError((result.stderr or result.stdout).strip())


def create_shortcut(name, target):
    desktop = Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Desktop"
    desktop.mkdir(parents=True, exist_ok=True)
    shortcut = desktop / f"{name}.lnk"
    ps = (
        "$w=New-Object -ComObject WScript.Shell;"
        f"$s=$w.CreateShortcut('{str(shortcut).replace("'", "''")}');"
        f"$s.TargetPath='{str(target).replace("'", "''")}';"
        f"$s.WorkingDirectory='{str(target.parent).replace("'", "''")}';"
        "$s.Save()"
    )
    result = run_cmd(["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps])
    if result.returncode != 0 or not shortcut.exists():
        raise RuntimeError(f"Could not create shortcut: {shortcut}")
    return shortcut


def setup_app():
    root = tk.Tk(); root.withdraw()
    try:
        dest = install_dir(); dest.mkdir(parents=True, exist_ok=True)
        for folder in ("config", "logs", "backup", "reports"):
            (dest / folder).mkdir(exist_ok=True)
        source = exe_dir()
        for name in ("SendPendingEmails.exe", "ConfigureSchedule.exe", "VerifyInstallation.exe"):
            src = source / name
            if not src.exists():
                raise FileNotFoundError(f"Deployment file is missing: {name}")
            shutil.copy2(src, dest / name)
        env_source = source / ".env"
        if not env_source.exists():
            env_source = source / ".env.template"
        if not env_source.exists():
            raise FileNotFoundError(".env or .env.template is missing from deployment package")
        shutil.copy2(env_source, dest / ".env")
        cfg = load_config(); save_config(cfg)
        create_shortcut("Send Pending Emails Now", dest / "SendPendingEmails.exe")
        create_shortcut("Configure Daily Email Schedule", dest / "ConfigureSchedule.exe")
        create_shortcut("Verify Installation", dest / "VerifyInstallation.exe")
        update_task(cfg, True)
        messagebox.showinfo(APP_NAME, f"Installation completed successfully.\n\nInstalled to:\n{dest}\n\nThree Desktop shortcuts were created.")
    except Exception as exc:
        messagebox.showerror(APP_NAME, f"Installation failed:\n\n{exc}")
    finally:
        root.destroy()


class AttachmentMultiSelect:
    def __init__(self, parent, selected_ids=None):
        self.frame = ttk.LabelFrame(parent, text="Attachments (Optional)", padding=8)
        self.frame.pack(fill="x", pady=(8, 0))
        self.saved_ids = [str(value) for value in (selected_ids or [])]
        self.attachments = []
        self.variables = {}
        self.button_text = tk.StringVar(value="Select Attachment(s)")
        self.button = ttk.Menubutton(self.frame, textvariable=self.button_text, width=30)
        self.menu = tk.Menu(self.button, tearoff=False)
        self.button.configure(menu=self.menu); self.button.pack(side="left")
        ttk.Button(self.frame, text="Refresh", command=self.refresh).pack(side="left", padx=8)
        self.status = tk.StringVar(value=""); ttk.Label(self.frame, textvariable=self.status).pack(side="left", padx=8)
        self.refresh()
    def refresh(self):
        self.status.set("Loading attachments…"); self.saved_ids = self.selected_ids()
        def worker():
            try:
                attachments = attachment_library_service().list_attachments()
                self.frame.after(0, self.loaded, attachments)
            except Exception as exc: self.frame.after(0, self.failed, str(exc))
        threading.Thread(target=worker, daemon=True).start()
    def loaded(self, attachments):
        self.attachments = attachments; self.variables = {}; self.menu.delete(0, "end")
        for attachment in attachments:
            variable = tk.BooleanVar(value=attachment["id"] in self.saved_ids); self.variables[attachment["id"]] = variable
            self.menu.add_checkbutton(label=attachment["file_name"], variable=variable, command=self.update_text)
        self.status.set(f"{len(attachments)} available"); self.update_text()
    def failed(self, error): self.status.set(f"Attachment server unavailable: {error}")
    def selected_ids(self):
        if not self.variables: return list(self.saved_ids)
        return [attachment_id for attachment_id, variable in self.variables.items() if variable.get()]
    def update_text(self):
        count = len(self.selected_ids()); self.button_text.set("Select Attachment(s)" if count == 0 else f"{count} attachment(s) selected")


class SenderWindow:
    def __init__(self, parent=None):
        self.root = tk.Toplevel(parent) if parent else tk.Tk()
        self.root.title("Send Pending Emails")
        self.root.geometry("920x590")
        self.root.minsize(760, 480)
        self.sending = False
        self.batch_sent = 0
        self.batch_failed = 0
        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="Email Sending Dashboard", font=("Segoe UI", 17, "bold")).pack(anchor="w")
        picker = ttk.Frame(outer); picker.pack(fill="x", pady=(10, 0))
        ttk.Label(picker, text="Excel File:").pack(side="left")
        self.excel_file = tk.StringVar(value=str(app_paths()["list"]))
        ttk.Entry(picker, textvariable=self.excel_file).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(picker, text="Browse", command=self.browse).pack(side="left")
        self.attachment_selector = AttachmentMultiSelect(outer)
        counts = ttk.Frame(outer)
        counts.pack(fill="x", pady=(14, 12))
        self.count_vars = {key: tk.StringVar(value=f"{label}: 0") for key, label in (("total", "Total Emails"), ("sent", "Sent Emails"), ("pending", "Pending Emails"), ("failed", "Failed Emails"))}
        for key in ("total", "sent", "pending", "failed"):
            ttk.Label(counts, textvariable=self.count_vars[key], font=("Segoe UI", 10, "bold"), padding=(0, 0, 28, 0)).pack(side="left")
        columns = ("name", "email", "status", "sentdate", "result")
        self.grid = ttk.Treeview(outer, columns=columns, show="headings", height=17)
        for key, title, width in (("name", "Name", 150), ("email", "Email", 220), ("status", "Status", 90), ("sentdate", "Sent Date", 150), ("result", "Result", 240)):
            self.grid.heading(key, text=title)
            self.grid.column(key, width=width, minwidth=70)
        scroll = ttk.Scrollbar(outer, orient="vertical", command=self.grid.yview)
        self.grid.configure(yscrollcommand=scroll.set)
        self.grid.pack(side="left", fill="both", expand=True, pady=(0, 75))
        scroll.pack(side="left", fill="y", pady=(0, 75))
        controls = ttk.Frame(self.root, padding=(16, 8, 16, 14))
        controls.place(relx=0, rely=1, relwidth=1, anchor="sw")
        self.send_button = ttk.Button(controls, text="Send Pending Emails", command=self.start_send)
        self.send_button.pack(side="left")
        ttk.Button(controls, text="Refresh", command=self.refresh).pack(side="left", padx=8)
        ttk.Button(controls, text="Close", command=self.root.destroy).pack(side="right")
        self.progress = ttk.Progressbar(controls, mode="determinate", length=230)
        self.progress.pack(side="right", padx=14)
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(controls, textvariable=self.status_var).pack(side="left", padx=14)
        self.refresh()

    def toast(self, message, error=False, duration=4500):
        toast = tk.Toplevel(self.root)
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        color = "#a61b1b" if error else "#176b36"
        frame = tk.Frame(toast, bg=color, padx=18, pady=13)
        frame.pack(fill="both", expand=True)
        tk.Label(frame, text=message, bg=color, fg="white", font=("Segoe UI", 10, "bold"), wraplength=420, justify="left").pack()
        toast.update_idletasks()
        x = self.root.winfo_screenwidth() - toast.winfo_reqwidth() - 28
        y = self.root.winfo_screenheight() - toast.winfo_reqheight() - 70
        toast.geometry(f"+{x}+{y}")
        toast.after(duration, toast.destroy)

    def refresh(self):
        try:
            rows, counts = email_grid_data(self.excel_file.get())
            for item in self.grid.get_children():
                self.grid.delete(item)
            for row in rows:
                self.grid.insert("", "end", values=row)
            labels = {"total": "Total Emails", "sent": "Sent Emails", "pending": "Pending Emails", "failed": "Failed Emails"}
            for key, label in labels.items():
                self.count_vars[key].set(f"{label}: {counts[key]}")
            return counts
        except Exception as exc:
            self.status_var.set(f"Refresh failed: {exc}")
            return {"total": 0, "sent": 0, "pending": 0, "failed": 0}

    def start_send(self):
        if self.sending:
            return
        counts = self.refresh()
        if counts["pending"] == 0:
            self.toast("There are no pending emails.")
            return
        value = simpledialog.askinteger(APP_NAME, "How many emails do you want to send now?", parent=self.root, minvalue=1)
        if value is None:
            return
        self.sending = True
        self.batch_sent = self.batch_failed = 0
        self.send_button.configure(state="disabled")
        self.progress.configure(maximum=min(value, counts["pending"]), value=0)
        self.status_var.set("Preparing to send…")
        excel_file = self.excel_file.get()
        attachment_ids = self.attachment_selector.selected_ids()

        def progress(number, total, result, email):
            if result == "Sent":
                self.batch_sent += 1
            else:
                self.batch_failed += 1
            self.root.after(0, self.on_progress, number, total, result, email)

        def worker():
            try:
                result = send_pending(value, True, progress, excel_file, attachment_ids)
                self.root.after(0, self.on_complete, result)
            except PermissionError:
                message = "mail_list.xlsx is open or locked. Close the Excel file, then try again."
                daily_log(status="System Error", error=message)
                self.root.after(0, self.on_error, message)
            except Exception as exc:
                daily_log(status="System Error", error=str(exc))
                self.root.after(0, self.on_error, str(exc))
        threading.Thread(target=worker, daemon=True).start()

    def on_progress(self, number, total, result, email):
        self.progress.configure(value=number, maximum=max(total, 1))
        counts = self.refresh()
        if result == "Sent":
            if counts["pending"] == 0:
                text = f"All {self.batch_sent} emails have been sent successfully."
            else:
                sent_word = "email" if self.batch_sent == 1 else "emails"
                pending_word = "email is" if counts["pending"] == 1 else "emails are"
                text = f"{self.batch_sent} {sent_word} sent successfully. {counts['pending']} {pending_word} pending."
            self.toast(text)
        else:
            self.toast(f"Failed to send email to {email}.", error=True)
        self.status_var.set(f"Processed {number} of {total}: {email}")

    def on_complete(self, result):
        self.sending = False
        self.send_button.configure(state="normal")
        counts = self.refresh()
        if result["failed"]:
            self.status_var.set(f"Finished: {result['sent']} sent, {result['failed']} failed")
        else:
            self.status_var.set(f"Finished: {result['sent']} sent successfully")
        if result["sent"] == 0 and result["failed"] == 0:
            self.toast("No pending emails were selected.")

    def on_error(self, message):
        self.sending = False
        self.send_button.configure(state="normal")
        self.refresh()
        self.status_var.set("Sending stopped")
        self.toast(message, error=True, duration=7000)

    def run(self):
        self.root.mainloop()

    def browse(self):
        value = filedialog.askopenfilename(parent=self.root, title="Select mail_list.xlsx", filetypes=[("Excel files", "*.xlsx")])
        if value: self.excel_file.set(value); self.refresh()


def manual_send():
    SenderWindow().run()


class ScheduleWindow:
    def __init__(self):
        self.root = tk.Tk(); self.root.title("Configure Daily Email Schedule"); self.root.geometry("610x470"); self.root.resizable(False, False)
        self.cfg = load_config()
        frame = ttk.Frame(self.root, padding=20); frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Daily Email Schedule", font=("Segoe UI", 17, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 18))
        self.info = tk.StringVar(); ttk.Label(frame, textvariable=self.info, justify="left", font=("Segoe UI", 10)).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 20))
        ttk.Separator(frame).grid(row=2, column=0, columnspan=3, sticky="ew", pady=8)
        ttk.Label(frame, text="Emails per day:").grid(row=3, column=0, sticky="w", pady=8)
        self.limit = tk.StringVar(value=str(self.cfg["daily_limit"])); ttk.Entry(frame, textvariable=self.limit, width=16).grid(row=3, column=1, sticky="w")
        ttk.Label(frame, text="Daily time (HH:MM):").grid(row=4, column=0, sticky="w", pady=8)
        self.when = tk.StringVar(value=self.cfg["schedule_time"]); ttk.Entry(frame, textvariable=self.when, width=16).grid(row=4, column=1, sticky="w")
        ttk.Label(frame, text="Default sender name:").grid(row=5, column=0, sticky="w", pady=8)
        self.sender_name = tk.StringVar(value=self.cfg.get("default_sender_name", "The Power People")); ttk.Entry(frame, textvariable=self.sender_name, width=30).grid(row=5, column=1, sticky="w")
        buttons = ttk.Frame(frame); buttons.grid(row=6, column=0, columnspan=3, sticky="w", pady=22)
        ttk.Button(buttons, text="Save Changes", command=self.save).pack(side="left", padx=(0, 7))
        ttk.Button(buttons, text="Enable Schedule", command=lambda: self.enable(True)).pack(side="left", padx=7)
        ttk.Button(buttons, text="Disable Schedule", command=lambda: self.enable(False)).pack(side="left", padx=7)
        ttk.Button(buttons, text="Delete Schedule", command=self.delete).pack(side="left", padx=7)
        ttk.Button(frame, text="Close", command=self.root.destroy).grid(row=7, column=0, sticky="w")
        ttk.Label(frame, text="Note: missed runs start when Windows next allows the task. Running while logged off may require Windows credentials and administrator policy.", wraplength=555, foreground="#555").grid(row=8, column=0, columnspan=3, sticky="w", pady=(25, 0))
        self.refresh()
    def refresh(self):
        state = task_query()
        status = "Enabled" if state["enabled"] else ("Disabled" if state["exists"] else "Not Created")
        self.info.set(f"Current Daily Email Limit:  {self.cfg['daily_limit']}\nCurrent Scheduled Time:  {self.cfg['schedule_time']}\nScheduled Task Status:  {status}\nNext Scheduled Run Time:  {state['next']}")
    def values(self):
        limit = int(self.limit.get())
        if limit < 1: raise ValueError("Emails per day must be at least 1.")
        datetime.strptime(self.when.get().strip(), "%H:%M")
        sender_name = self.sender_name.get().strip()
        if not sender_name: raise ValueError("Default sender name cannot be empty.")
        return limit, self.when.get().strip(), sender_name
    def save(self):
        try:
            limit, when, sender_name = self.values(); self.cfg.update({"daily_limit": limit, "schedule_time": when, "default_sender_name": sender_name}); save_config(self.cfg); update_task(self.cfg, True); self.refresh(); messagebox.showinfo(APP_NAME, "Configuration saved and schedule enabled.")
        except Exception as exc: messagebox.showerror(APP_NAME, str(exc))
    def enable(self, enabled):
        try:
            if enabled and not task_query()["exists"]: update_task(self.cfg, True)
            else: set_task_enabled(enabled)
            self.refresh()
        except Exception as exc: messagebox.showerror(APP_NAME, str(exc))
    def delete(self):
        if not messagebox.askyesno(APP_NAME, "Delete the daily scheduled task?"): return
        try: delete_task(); self.refresh()
        except Exception as exc: messagebox.showerror(APP_NAME, str(exc))
    def run(self): self.root.mainloop()


class MailAccountEditor:
    def __init__(self, parent, account=None):
        self.account = account or {}
        self.result = None
        self.window = tk.Toplevel(parent); self.window.title("Edit Mail Account" if account else "Add Mail Account"); self.window.geometry("620x650"); self.window.resizable(False, False); self.window.transient(parent); self.window.grab_set()
        frame = ttk.Frame(self.window, padding=20); frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Mail Provider Account", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 16))
        self.provider = tk.StringVar(value=self.account.get("provider", "Gmail"))
        self.display_name = tk.StringVar(value=self.account.get("display_name", self.account.get("name", "")))
        self.email = tk.StringVar(value=self.account.get("email", ""))
        self.host = tk.StringVar(value=self.account.get("smtp_host", "smtp.gmail.com"))
        self.port = tk.StringVar(value=str(self.account.get("smtp_port", 587)))
        self.encryption = tk.StringVar(value=self.account.get("encryption", "STARTTLS"))
        self.imap_host = tk.StringVar(value=self.account.get("imap_host", "imap.gmail.com"))
        self.imap_port = tk.StringVar(value=str(self.account.get("imap_port", 993)))
        self.imap_encryption = tk.StringVar(value=self.account.get("imap_encryption", "SSL/TLS"))
        self.password = tk.StringVar()
        self.enabled = tk.BooleanVar(value=self.account.get("enabled", True))
        labels = (("Mail Provider:", 1), ("Display Name:", 2), ("Email Address:", 3), ("SMTP Host:", 4), ("SMTP Port:", 5), ("SMTP Encryption:", 6), ("IMAP Host:", 7), ("IMAP Port:", 8), ("IMAP Encryption:", 9), ("Password / App Password:", 10))
        for text, row in labels: ttk.Label(frame, text=text).grid(row=row, column=0, sticky="w", pady=7)
        self.provider_box = ttk.Combobox(frame, textvariable=self.provider, values=AccountService.PROVIDERS, state="readonly", width=32); self.provider_box.grid(row=1, column=1, sticky="w"); self.provider_box.bind("<<ComboboxSelected>>", self.provider_changed)
        ttk.Entry(frame, textvariable=self.display_name, width=38).grid(row=2, column=1, sticky="w")
        ttk.Entry(frame, textvariable=self.email, width=38).grid(row=3, column=1, sticky="w")
        self.host_entry = ttk.Entry(frame, textvariable=self.host, width=38); self.host_entry.grid(row=4, column=1, sticky="w")
        self.port_entry = ttk.Entry(frame, textvariable=self.port, width=15); self.port_entry.grid(row=5, column=1, sticky="w")
        self.encryption_box = ttk.Combobox(frame, textvariable=self.encryption, values=AccountService.ENCRYPTION_TYPES, state="readonly", width=32); self.encryption_box.grid(row=6, column=1, sticky="w")
        self.imap_host_entry = ttk.Entry(frame, textvariable=self.imap_host, width=38); self.imap_host_entry.grid(row=7, column=1, sticky="w")
        self.imap_port_entry = ttk.Entry(frame, textvariable=self.imap_port, width=15); self.imap_port_entry.grid(row=8, column=1, sticky="w")
        self.imap_encryption_box = ttk.Combobox(frame, textvariable=self.imap_encryption, values=AccountService.ENCRYPTION_TYPES, state="readonly", width=32); self.imap_encryption_box.grid(row=9, column=1, sticky="w")
        ttk.Entry(frame, textvariable=self.password, show="*", width=38).grid(row=10, column=1, sticky="w")
        password_note = "Leave empty while editing to keep the stored password." if account else "Gmail requires a Google App Password."
        ttk.Label(frame, text=password_note, foreground="#555").grid(row=11, column=1, sticky="w", pady=(0, 8))
        ttk.Checkbutton(frame, text="Enabled", variable=self.enabled).grid(row=12, column=1, sticky="w", pady=8)
        buttons = ttk.Frame(frame); buttons.grid(row=13, column=0, columnspan=2, sticky="e", pady=(18, 0))
        ttk.Button(buttons, text="Save", command=self.save).pack(side="left", padx=5); ttk.Button(buttons, text="Cancel", command=self.window.destroy).pack(side="left", padx=5)
        self.apply_provider_state(reset_defaults=False)
        self.window.wait_window()
    def provider_changed(self, _event=None): self.apply_provider_state(reset_defaults=True)
    def apply_provider_state(self, reset_defaults=False):
        gmail = self.provider.get() == "Gmail"
        if gmail:
            self.host.set("smtp.gmail.com"); self.port.set("587"); self.encryption.set("STARTTLS"); self.imap_host.set("imap.gmail.com"); self.imap_port.set("993"); self.imap_encryption.set("SSL/TLS")
        elif reset_defaults:
            self.host.set(""); self.port.set("587"); self.encryption.set("STARTTLS"); self.imap_host.set(""); self.imap_port.set("993"); self.imap_encryption.set("SSL/TLS")
        self.host_entry.configure(state="disabled" if gmail else "normal")
        self.port_entry.configure(state="disabled" if gmail else "normal")
        self.encryption_box.configure(state="disabled" if gmail else "readonly")
        self.imap_host_entry.configure(state="disabled" if gmail else "normal")
        self.imap_port_entry.configure(state="disabled" if gmail else "normal")
        self.imap_encryption_box.configure(state="disabled" if gmail else "readonly")
    def save(self):
        try:
            provider = self.provider.get()
            host = "smtp.gmail.com" if provider == "Gmail" else self.host.get().strip()
            port = 587 if provider == "Gmail" else self.port.get().strip()
            encryption = "STARTTLS" if provider == "Gmail" else self.encryption.get()
            imap_host = "imap.gmail.com" if provider == "Gmail" else self.imap_host.get().strip()
            imap_port = 993 if provider == "Gmail" else self.imap_port.get().strip()
            imap_encryption = "SSL/TLS" if provider == "Gmail" else self.imap_encryption.get()
            if not host: raise ValueError("SMTP Host cannot be empty.")
            if not str(port).isdigit(): raise ValueError("SMTP Port must be numeric.")
            if not imap_host: raise ValueError("IMAP Host cannot be empty.")
            if not str(imap_port).isdigit(): raise ValueError("IMAP Port must be numeric.")
            if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", self.email.get().strip()): raise ValueError("A valid email address is required.")
            self.result = {"name": self.display_name.get().strip(), "email": self.email.get().strip(), "password": self.password.get(), "enabled": self.enabled.get(), "provider": provider, "smtp_host": host, "smtp_port": port, "encryption": encryption, "imap_host": imap_host, "imap_port": imap_port, "imap_encryption": imap_encryption}
            self.window.destroy()
        except ValueError as exc: messagebox.showerror(APP_NAME, str(exc), parent=self.window)


class AccountWindow:
    def __init__(self, parent=None):
        self.root = tk.Toplevel(parent) if parent else tk.Tk()
        self.root.title("Mail Account Setup")
        self.root.geometry("1050x460")
        self.service = account_service()
        frame = ttk.Frame(self.root, padding=16); frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Mail Account Management", font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(0, 12))
        self.tree = ttk.Treeview(frame, columns=("provider", "name", "email", "host", "port", "encryption", "status"), show="headings")
        for key, title, width in (("provider", "Provider", 110), ("name", "Display Name", 150), ("email", "Email Address", 210), ("host", "SMTP Host", 180), ("port", "Port", 60), ("encryption", "Encryption", 90), ("status", "Status", 80)):
            self.tree.heading(key, text=title); self.tree.column(key, width=width)
        self.tree.pack(fill="both", expand=True)
        bar = ttk.Frame(frame); bar.pack(fill="x", pady=(12, 0))
        for text, command in (("Add", self.add), ("Edit", self.edit), ("Delete", self.delete), ("Enable", lambda: self.toggle(True)), ("Disable", lambda: self.toggle(False)), ("Test SMTP Connection", self.test)):
            ttk.Button(bar, text=text, command=command).pack(side="left", padx=(0, 7))
        ttk.Button(bar, text="Close", command=self.root.destroy).pack(side="right")
        self.refresh()
    def selected(self):
        selected = self.tree.selection()
        return self.service.find(self.tree.item(selected[0], "values")[2]) if selected else None
    def refresh(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for account in self.service.list_accounts():
            self.tree.insert("", "end", values=(account["provider"], account["display_name"], account["email"], account["smtp_host"], account["smtp_port"], account["encryption"], "Enabled" if account.get("enabled", True) else "Disabled"))
    def account_form(self, existing=None):
        values = MailAccountEditor(self.root, existing).result
        if values:
            self.service.save_account(original_email=existing["email"] if existing else None, **values); self.refresh()
    def add(self):
        try: self.account_form()
        except Exception as exc: messagebox.showerror(APP_NAME, str(exc), parent=self.root)
    def edit(self):
        value = self.selected()
        if not value: return
        try: self.account_form(value)
        except Exception as exc: messagebox.showerror(APP_NAME, str(exc), parent=self.root)
    def delete(self):
        value = self.selected()
        if value and messagebox.askyesno(APP_NAME, f"Delete {value['email']} permanently?", parent=self.root):
            try: self.service.delete_account(value["email"]); self.refresh()
            except Exception as exc: messagebox.showerror(APP_NAME, str(exc), parent=self.root)
    def toggle(self, enabled):
        value = self.selected()
        if value: self.service.set_enabled(value["email"], enabled); self.refresh()
    def test(self):
        value = self.selected()
        if not value: return
        try: self.service.test_smtp(value["email"]); messagebox.showinfo(APP_NAME, "SMTP connection successful.", parent=self.root)
        except Exception as exc: messagebox.showerror(APP_NAME, str(exc), parent=self.root)


def load_schedules():
    path = app_paths()["schedules"]
    if not path.exists(): return []
    try: return json.loads(path.read_text(encoding="utf-8"))
    except Exception: return []


def save_schedules(items):
    path = app_paths()["schedules"]; path.parent.mkdir(parents=True, exist_ok=True); path.write_text(json.dumps(items, indent=2), encoding="utf-8")


def schedule_task_name(item):
    return f"Email Automation - {item['id']}"


def sync_schedule_task(item):
    name = schedule_task_name(item)
    if not item.get("enabled", True):
        run_cmd(["schtasks", "/Delete", "/TN", name, "/F"]); return
    days = ",".join(day[:3].upper() for day in item["days"])
    if not days: raise ValueError("Select at least one day.")
    exe = install_dir() / "Email Automation.exe"
    action = f'"{exe}" --run-schedule {item["id"]}'
    result = run_cmd(["schtasks", "/Create", "/TN", name, "/TR", action, "/SC", "WEEKLY", "/D", days, "/ST", item["time"], "/F", "/RL", "LIMITED"])
    if result.returncode != 0: raise RuntimeError((result.stderr or result.stdout).strip())
    ps = "$s=New-Object -ComObject 'Schedule.Service';$s.Connect();$t=$s.GetFolder('\\').GetTask('" + name.replace("'", "''") + "');$d=$t.Definition;$d.Settings.StartWhenAvailable=$true;$s.GetFolder('\\').RegisterTaskDefinition($t.Name,$d,6,$null,$null,3,$null)|Out-Null"
    run_cmd(["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps])


class ScheduleEditor:
    DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

    def __init__(self, parent, current=None):
        self.current = current or {}
        self.result = None
        self.window = tk.Toplevel(parent)
        self.window.title("Edit Schedule" if current else "Add Schedule")
        self.window.geometry("620x680")
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()
        frame = ttk.Frame(self.window, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Schedule Details", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 16))
        self.name = tk.StringVar(value=self.current.get("name", ""))
        self.excel = tk.StringVar(value=self.current.get("excel_file", ""))
        self.when = tk.StringVar(value=self.current.get("time", "09:00"))
        self.maximum = tk.StringVar(value=str(self.current.get("max_emails", 10)))
        self.enabled = tk.BooleanVar(value=self.current.get("enabled", True))
        ttk.Label(frame, text="Schedule Name:").grid(row=1, column=0, sticky="w", pady=7)
        ttk.Entry(frame, textvariable=self.name, width=48).grid(row=1, column=1, columnspan=2, sticky="ew", pady=7)
        ttk.Label(frame, text="Excel File:").grid(row=2, column=0, sticky="w", pady=7)
        ttk.Entry(frame, textvariable=self.excel, width=40).grid(row=2, column=1, sticky="ew", pady=7)
        ttk.Button(frame, text="Browse", command=self.browse).grid(row=2, column=2, padx=(8, 0), pady=7)
        ttk.Label(frame, text="Days of Week:").grid(row=3, column=0, sticky="nw", pady=(12, 7))
        day_frame = ttk.LabelFrame(frame, text="Select one or more days", padding=12)
        day_frame.grid(row=3, column=1, columnspan=2, sticky="ew", pady=(12, 7))
        saved_days = set(self.current.get("days", []))
        self.day_vars = {}
        for index, day in enumerate(self.DAYS):
            variable = tk.BooleanVar(value=day in saved_days)
            self.day_vars[day] = variable
            ttk.Checkbutton(day_frame, text=day, variable=variable).grid(row=index // 2, column=index % 2, sticky="w", padx=(0, 36), pady=4)
        attachment_host = ttk.Frame(frame); attachment_host.grid(row=4, column=0, columnspan=3, sticky="ew")
        self.attachment_selector = AttachmentMultiSelect(attachment_host, self.current.get("attachment_ids", []))
        ttk.Label(frame, text="Time (HH:mm):").grid(row=5, column=0, sticky="w", pady=7)
        ttk.Entry(frame, textvariable=self.when, width=16).grid(row=5, column=1, sticky="w", pady=7)
        ttk.Label(frame, text="Maximum Emails:").grid(row=6, column=0, sticky="w", pady=7)
        ttk.Entry(frame, textvariable=self.maximum, width=16).grid(row=6, column=1, sticky="w", pady=7)
        ttk.Checkbutton(frame, text="Schedule Enabled", variable=self.enabled).grid(row=7, column=1, sticky="w", pady=10)
        buttons = ttk.Frame(frame)
        buttons.grid(row=8, column=0, columnspan=3, sticky="e", pady=(18, 0))
        ttk.Button(buttons, text="Save", command=self.save).pack(side="left", padx=5)
        ttk.Button(buttons, text="Cancel", command=self.window.destroy).pack(side="left", padx=5)
        frame.columnconfigure(1, weight=1)
        self.window.protocol("WM_DELETE_WINDOW", self.window.destroy)
        self.window.wait_window()

    def browse(self):
        value = filedialog.askopenfilename(parent=self.window, title="Select mail_list.xlsx", filetypes=[("Excel files", "*.xlsx")])
        if value:
            self.excel.set(value)

    def save(self):
        selected_days = [day for day in self.DAYS if self.day_vars[day].get()]
        if not selected_days:
            messagebox.showerror(APP_NAME, "Please select at least one day.", parent=self.window)
            return
        try:
            if not self.name.get().strip():
                raise ValueError("Schedule Name is required.")
            if not self.excel.get().strip():
                raise ValueError("Excel File is required.")
            datetime.strptime(self.when.get().strip(), "%H:%M")
            maximum = int(self.maximum.get())
            if maximum < 1:
                raise ValueError("Maximum Emails must be at least 1.")
            self.result = {
                "id": self.current.get("id", str(uuid.uuid4())),
                "name": self.name.get().strip(),
                "excel_file": self.excel.get().strip(),
                "days": selected_days,
                "time": self.when.get().strip(),
                "max_emails": maximum,
                "attachment_ids": self.attachment_selector.selected_ids(),
                "enabled": self.enabled.get(),
            }
            self.window.destroy()
        except ValueError as exc:
            messagebox.showerror(APP_NAME, str(exc), parent=self.window)


class MultiScheduleWindow:
    DAYS = ScheduleEditor.DAYS
    def __init__(self, parent=None):
        self.root = tk.Toplevel(parent) if parent else tk.Tk(); self.root.title("Daily Scheduling"); self.root.geometry("900x480")
        frame = ttk.Frame(self.root, padding=16); frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Recurring Email Schedules", font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(0, 12))
        self.tree = ttk.Treeview(frame, columns=("name", "file", "days", "time", "maximum", "status"), show="headings")
        for key, title, width in (("name", "Schedule Name", 140), ("file", "Excel File", 220), ("days", "Days", 190), ("time", "Time", 70), ("maximum", "Maximum", 70), ("status", "Status", 80)):
            self.tree.heading(key, text=title); self.tree.column(key, width=width)
        self.tree.pack(fill="both", expand=True)
        bar=ttk.Frame(frame);bar.pack(fill="x",pady=12)
        for text,cmd in (("Add",self.add),("Edit",self.edit),("Delete",self.delete),("Enable",lambda:self.toggle(True)),("Disable",lambda:self.toggle(False)),("Duplicate",self.duplicate)):
            ttk.Button(bar,text=text,command=cmd).pack(side="left",padx=(0,7))
        ttk.Button(bar,text="Close",command=self.root.destroy).pack(side="right");self.refresh()
    def items(self): return load_schedules()
    def selected_id(self):
        values=self.tree.selection(); return self.tree.item(values[0],"tags")[0] if values else None
    def refresh(self):
        for row in self.tree.get_children(): self.tree.delete(row)
        for item in self.items(): self.tree.insert("","end",values=(item["name"],item["excel_file"],", ".join(item["days"]),item["time"],item["max_emails"],"Enabled" if item.get("enabled",True) else "Disabled"),tags=(item["id"],))
    def form(self, item=None):
        return ScheduleEditor(self.root, item).result
    def add(self):
        try:
            item=self.form()
            if item: values=self.items();values.append(item);save_schedules(values);sync_schedule_task(item);self.refresh()
        except Exception as exc:messagebox.showerror(APP_NAME,str(exc),parent=self.root)
    def edit(self):
        key=self.selected_id();values=self.items();old=next((x for x in values if x["id"]==key),None)
        if not old:return
        try:
            updated=self.form(old)
            if updated: values=[updated if x["id"]==key else x for x in values];save_schedules(values);sync_schedule_task(updated);self.refresh()
        except Exception as exc:messagebox.showerror(APP_NAME,str(exc),parent=self.root)
    def delete(self):
        key=self.selected_id()
        if not key or not messagebox.askyesno(APP_NAME,"Delete this schedule?",parent=self.root):return
        values=self.items();item=next(x for x in values if x["id"]==key);run_cmd(["schtasks","/Delete","/TN",schedule_task_name(item),"/F"]);save_schedules([x for x in values if x["id"]!=key]);self.refresh()
    def toggle(self,enabled):
        key=self.selected_id();values=self.items()
        for item in values:
            if item["id"]==key:item["enabled"]=enabled;sync_schedule_task(item)
        save_schedules(values);self.refresh()
    def duplicate(self):
        key=self.selected_id();values=self.items();old=next((x for x in values if x["id"]==key),None)
        if old:
            item=dict(old);item["id"]=str(uuid.uuid4());item["name"] += " Copy";values.append(item);save_schedules(values);sync_schedule_task(item);self.refresh()


class GlobalSettingsWindow:
    def __init__(self, parent):
        self.root=tk.Toplevel(parent);self.root.title("Settings");self.root.geometry("520x470");self.cfg=load_config();frame=ttk.Frame(self.root,padding=18);frame.pack(fill="both",expand=True)
        ttk.Label(frame,text="Global Settings",font=("Segoe UI",16,"bold")).grid(row=0,column=0,columnspan=2,sticky="w",pady=(0,15))
        fields=(("random_delay_min","Minimum Delay Seconds"),("random_delay_max","Maximum Delay Seconds"),("retry_count","Retry Count"),("log_retention_days","Log Retention Days"),("default_sender_name","Default Sender Name"))
        self.vars={}
        for row,(key,label) in enumerate(fields,1):
            ttk.Label(frame,text=label+":").grid(row=row,column=0,sticky="w",pady=6);self.vars[key]=tk.StringVar(value=str(self.cfg.get(key,"")));ttk.Entry(frame,textvariable=self.vars[key],width=28).grid(row=row,column=1,sticky="w")
        self.backup=tk.BooleanVar(value=self.cfg.get("backup_enabled",True));ttk.Checkbutton(frame,text="Backup Enabled",variable=self.backup).grid(row=6,column=0,columnspan=2,sticky="w",pady=8)
        ttk.Label(frame,text="Default Theme:").grid(row=7,column=0,sticky="w");self.theme=tk.StringVar(value=self.cfg.get("theme","Light"));ttk.Combobox(frame,textvariable=self.theme,values=("Light","Dark"),state="readonly",width=25).grid(row=7,column=1,sticky="w")
        bar=ttk.Frame(frame);bar.grid(row=8,column=0,columnspan=2,sticky="w",pady=18)
        ttk.Button(bar,text="Save",command=self.save).pack(side="left",padx=(0,7))
        for text,key in (("Open Backup Folder","backup"),("Open Reports Folder","reports"),("Open Logs Folder","logs")):
            ttk.Button(bar,text=text,command=lambda k=key:os.startfile(app_paths()[k])).pack(side="left",padx=(0,7))
        ttk.Button(frame,text="Check for Updates (Future)",command=lambda:messagebox.showinfo(APP_NAME,"Update service is future-ready.",parent=self.root)).grid(row=9,column=0,columnspan=2,sticky="w")
    def save(self):
        try:
            for key in ("random_delay_min","random_delay_max","retry_count","log_retention_days"):self.cfg[key]=int(self.vars[key].get())
            self.cfg["default_sender_name"]=self.vars["default_sender_name"].get().strip();self.cfg["backup_enabled"]=self.backup.get();self.cfg["theme"]=self.theme.get();save_config(self.cfg);messagebox.showinfo(APP_NAME,"Settings saved.",parent=self.root)
        except Exception as exc:messagebox.showerror(APP_NAME,str(exc),parent=self.root)


class SettingsWindow:
    def __init__(self, parent):
        self.root = tk.Toplevel(parent); self.root.title("Settings"); self.root.geometry("520x300"); self.root.resizable(False, False)
        frame = ttk.Frame(self.root, padding=24); frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Settings", font=("Segoe UI", 18, "bold")).pack(anchor="w", pady=(0, 22))
        ttk.Button(frame, text="Global Settings", command=lambda:GlobalSettingsWindow(self.root)).pack(fill="x", ipady=16, pady=(0, 14))
        ttk.Button(frame, text="Attachment Library", command=lambda:AttachmentLibraryWindow(self.root)).pack(fill="x", ipady=16)


class AttachmentLibraryWindow:
    def __init__(self, parent):
        self.root = tk.Toplevel(parent); self.root.title("Attachment Library"); self.root.geometry("790x520")
        self.service = attachment_library_service(); self.attachments_by_item = {}
        frame = ttk.Frame(self.root, padding=18); frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Attachment Library", font=("Segoe UI", 17, "bold")).pack(anchor="w", pady=(0, 14))
        self.tree = ttk.Treeview(frame, columns=("name", "size", "date"), show="headings", height=16)
        for key, title, width in (("name", "File Name", 340), ("size", "File Size", 130), ("date", "Upload Date", 240)):
            self.tree.heading(key, text=title); self.tree.column(key, width=width)
        self.tree.pack(fill="both", expand=True)
        self.progress = ttk.Progressbar(frame, mode="determinate"); self.progress.pack(fill="x", pady=(12, 4))
        self.status = tk.StringVar(value="Ready"); ttk.Label(frame, textvariable=self.status).pack(anchor="w", pady=(0, 8))
        bar = ttk.Frame(frame); bar.pack(fill="x")
        self.upload_button = ttk.Button(bar, text="Upload", command=self.upload); self.upload_button.pack(side="left", padx=(0, 7))
        self.delete_button = ttk.Button(bar, text="Delete", command=self.delete); self.delete_button.pack(side="left", padx=7)
        self.refresh_button = ttk.Button(bar, text="Refresh", command=self.refresh); self.refresh_button.pack(side="left", padx=7)
        ttk.Button(bar, text="Close", command=self.root.destroy).pack(side="right")
        self.refresh()
    @staticmethod
    def display_size(size):
        value = float(size)
        for unit in ("B", "KB", "MB", "GB"):
            if value < 1024 or unit == "GB": return f"{value:.0f} {unit}" if unit == "B" else f"{value:.1f} {unit}"
            value /= 1024
    def set_busy(self, busy):
        state = "disabled" if busy else "normal"
        self.upload_button.configure(state=state); self.delete_button.configure(state=state); self.refresh_button.configure(state=state)
    def refresh(self):
        self.set_busy(True); self.status.set("Loading attachments…")
        def worker():
            try: result = self.service.list_attachments(); self.root.after(0, self.refresh_complete, result)
            except Exception as exc: self.root.after(0, self.failed, "Refresh", str(exc))
        threading.Thread(target=worker, daemon=True).start()
    def refresh_complete(self, attachments):
        for item in self.tree.get_children(): self.tree.delete(item)
        self.attachments_by_item.clear()
        for index, attachment in enumerate(attachments):
            item = self.tree.insert("", "end", values=(attachment["file_name"], self.display_size(attachment["file_size"]), attachment["upload_date"]))
            self.attachments_by_item[item] = attachment
        self.status.set(f"Loaded {len(attachments)} attachment(s)."); self.progress.configure(value=0); self.set_busy(False)
    def upload(self):
        files = filedialog.askopenfilenames(parent=self.root, title="Select attachment files")
        if not files: return
        self.set_busy(True); self.status.set("Uploading…"); self.progress.configure(value=0, maximum=100)
        def worker():
            outcomes = []
            for file_path in files:
                name = Path(file_path).name
                try:
                    self.service.upload(file_path, lambda sent,total,n=name:self.root.after(0, self.upload_progress, n, sent, total))
                    outcomes.append((name, True, "Upload Successful"))
                except Exception as exc: outcomes.append((name, False, f"Upload Failed: {exc}"))
            self.root.after(0, self.upload_complete, outcomes)
        threading.Thread(target=worker, daemon=True).start()
    def upload_progress(self, name, sent, total):
        self.status.set(f"Uploading... {name}"); self.progress.configure(value=(sent / total * 100) if total else 100)
    def upload_complete(self, outcomes):
        text = "\n".join(f"{name}: {message}" for name, _ok, message in outcomes); self.status.set(text)
        if all(ok for _name, ok, _message in outcomes): messagebox.showinfo(APP_NAME, text, parent=self.root)
        else: messagebox.showwarning(APP_NAME, text, parent=self.root)
        self.set_busy(False); self.refresh()
    def delete(self):
        selected = self.tree.selection()
        if not selected: messagebox.showwarning(APP_NAME, "Select an attachment to delete.", parent=self.root); return
        attachment = self.attachments_by_item[selected[0]]
        if not messagebox.askyesno(APP_NAME, f"Delete {attachment['file_name']}?", parent=self.root): return
        self.set_busy(True); self.status.set("Deleting...")
        def worker():
            try: self.service.delete(attachment["id"]); self.root.after(0, self.delete_complete)
            except Exception as exc: self.root.after(0, self.failed, "Delete", str(exc))
        threading.Thread(target=worker, daemon=True).start()
    def delete_complete(self):
        self.status.set("Deleted Successfully"); messagebox.showinfo(APP_NAME, "Deleted Successfully", parent=self.root); self.set_busy(False); self.refresh()
    def failed(self, operation, error):
        self.set_busy(False); self.progress.configure(value=0); self.status.set(f"{operation} failed: {error}"); messagebox.showerror(APP_NAME, f"The attachment server is unavailable or returned an error.\n\n{error}", parent=self.root)


class LogWindow:
    def __init__(self,parent):
        self.root=tk.Toplevel(parent);self.root.title("Logs");self.root.geometry("1050x560");frame=ttk.Frame(self.root,padding=14);frame.pack(fill="both",expand=True)
        top=ttk.Frame(frame);top.pack(fill="x");self.search=tk.StringVar();ttk.Entry(top,textvariable=self.search,width=40).pack(side="left");ttk.Button(top,text="Search / Refresh",command=self.refresh).pack(side="left",padx=7);ttk.Button(top,text="Export CSV",command=self.export).pack(side="left");ttk.Button(top,text="Open Log Folder",command=lambda:os.startfile(app_paths()["logs"])).pack(side="right")
        cols=("date","time","tracking","sender","recipient","subject","status","response","error","duration");self.tree=ttk.Treeview(frame,columns=cols,show="headings")
        for key,title,width in zip(cols,("Date","Time","TrackingId","Sender Email","Recipient Email","Subject","Status","SMTP Response","Error","Seconds"),(85,75,230,170,170,190,75,100,220,70)):
            self.tree.heading(key,text=title);self.tree.column(key,width=width)
        self.tree.pack(fill="both",expand=True,pady=10);self.refresh()
    def rows(self):
        result=[]
        for path in sorted(app_paths()["logs"].glob("*.log")):
            with path.open(encoding="utf-8",errors="replace",newline="") as f:
                for row in list(csv.reader(f))[1:]:
                    if len(row)>=10: result.append(row[:10])
                    elif len(row)>=9: result.append(row[:2] + [""] + row[2:9])
        return result
    def refresh(self):
        query=self.search.get().lower()
        for item in self.tree.get_children():self.tree.delete(item)
        for row in self.rows():
            if not query or query in " ".join(row).lower():self.tree.insert("","end",values=row)
    def export(self):
        path=filedialog.asksaveasfilename(parent=self.root,defaultextension=".csv",filetypes=[("CSV","*.csv")])
        if path:
            with open(path,"w",encoding="utf-8",newline="") as f:csv.writer(f).writerows([[self.tree.heading(c,"text") for c in self.tree["columns"]]]+[list(self.tree.item(i,"values")) for i in self.tree.get_children()])


class ReportsWindow:
    def __init__(self,parent):
        self.root=tk.Toplevel(parent);self.root.title("Reports");self.root.geometry("900x540");frame=ttk.Frame(self.root,padding=14);frame.pack(fill="both",expand=True)
        top=ttk.Frame(frame);top.pack(fill="x");self.status=tk.StringVar(value="All");ttk.Label(top,text="Status:").pack(side="left");ttk.Combobox(top,textvariable=self.status,values=("All","Pending","Sent","Failed"),state="readonly",width=12).pack(side="left",padx=7);ttk.Button(top,text="Refresh",command=self.refresh).pack(side="left");ttk.Button(top,text="Export CSV",command=self.export_csv).pack(side="right");ttk.Button(top,text="Export Excel",command=self.export_excel).pack(side="right",padx=7);ttk.Button(top,text="Export PDF",command=self.export_pdf).pack(side="right")
        self.tree=ttk.Treeview(frame,columns=("name","email","company","sender","status","date","result"),show="headings")
        for key,title,width in (("name","Name",130),("email","Email",190),("company","Company",130),("sender","Sender Email",180),("status","Status",75),("date","Sent Date",130),("result","Result",180)):
            self.tree.heading(key,text=title);self.tree.column(key,width=width)
        self.tree.pack(fill="both",expand=True,pady=10);self.refresh()
    def data(self):
        p=app_paths()["list"];wb=load_workbook(p,read_only=True,data_only=True);ws=wb.active;h=header_map(ws);rows=[]
        for r in range(2,ws.max_row+1):
            get=lambda k:str(ws.cell(r,h[k]).value or "") if k in h else "";status=get("status")
            if self.status.get()!="All" and status.lower()!=self.status.get().lower():continue
            rows.append((f"{get('first_name')} {get('last_name')}".strip(),get("email"),get("company"),get("sender_email") or get("sender_mail"),status,get("sentdate"),get("result")))
        wb.close();return rows
    def refresh(self):
        for i in self.tree.get_children():self.tree.delete(i)
        for row in self.data():self.tree.insert("","end",values=row)
    def export_csv(self):
        p=filedialog.asksaveasfilename(parent=self.root,defaultextension=".csv",filetypes=[("CSV","*.csv")]);
        if p:
            with open(p,"w",encoding="utf-8",newline="") as f:csv.writer(f).writerows([[self.tree.heading(c,"text") for c in self.tree["columns"]]]+self.data())
    def export_excel(self):
        p=filedialog.asksaveasfilename(parent=self.root,defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")]);
        if p:
            wb=Workbook();ws=wb.active;ws.append([self.tree.heading(c,"text") for c in self.tree["columns"]]);[ws.append(row) for row in self.data()];wb.save(p)
    def export_pdf(self):
        p=filedialog.asksaveasfilename(parent=self.root,defaultextension=".pdf",filetypes=[("PDF","*.pdf")]);
        if p: write_simple_pdf(p,[" | ".join(map(str,row)) for row in self.data()]);messagebox.showinfo(APP_NAME,"PDF exported.",parent=self.root)


def write_simple_pdf(path,lines):
    safe=lambda s:str(s).replace("\\","\\\\").replace("(","\\(").replace(")","\\)").encode("latin-1","replace").decode("latin-1")
    content="BT /F1 8 Tf 35 800 Td "+" ".join(f"({safe(line[:150])}) Tj 0 -12 Td" for line in (["Email Automation Report"]+lines[:60]))+" ET";objects=["1 0 obj<< /Type /Catalog /Pages 2 0 R >>endobj","2 0 obj<< /Type /Pages /Kids [3 0 R] /Count 1 >>endobj","3 0 obj<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 842] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>endobj","4 0 obj<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>endobj",f"5 0 obj<< /Length {len(content.encode('latin-1'))} >>stream\n{content}\nendstream endobj"]
    data=b"%PDF-1.4\n";offsets=[0]
    for obj in objects:offsets.append(len(data));data+=obj.encode("latin-1")+b"\n"
    xref=len(data);data+=f"xref\n0 {len(objects)+1}\n0000000000 65535 f \n".encode();data+=b"".join(f"{o:010d} 00000 n \n".encode() for o in offsets[1:]);data+=f"trailer<< /Size {len(objects)+1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF".encode();Path(path).write_bytes(data)


class TrackingSynchronizationWindow:
    def __init__(self, parent, on_config_saved=None, on_sync_complete=None):
        self.root = tk.Toplevel(parent)
        self.root.title("Tracking Synchronization")
        self.root.geometry("570x410")
        self.root.resizable(False, False)
        self.on_config_saved = on_config_saved
        self.on_sync_complete_callback = on_sync_complete
        self.cfg = load_config()
        frame = ttk.Frame(self.root, padding=22)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Tracking Synchronization", font=("Segoe UI", 17, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 20))
        self.enabled = tk.BooleanVar(value=self.cfg.get("tracking_sync_enabled", False))
        ttk.Checkbutton(frame, text="Enable Automatic Synchronization", variable=self.enabled).grid(row=1, column=0, columnspan=3, sticky="w", pady=8)
        ttk.Label(frame, text="Every").grid(row=2, column=0, sticky="w", pady=10)
        self.interval = tk.StringVar(value=str(self.cfg.get("tracking_sync_interval_hours", 5)))
        ttk.Spinbox(frame, from_=1, to=24, textvariable=self.interval, width=8).grid(row=2, column=1, sticky="w")
        ttk.Label(frame, text="Hours").grid(row=2, column=2, sticky="w")
        self.last_sync = tk.StringVar(value=self.cfg.get("tracking_last_sync_time") or "Never")
        ttk.Label(frame, text="Last Synchronization:").grid(row=3, column=0, sticky="nw", pady=10)
        ttk.Label(frame, textvariable=self.last_sync, wraplength=330).grid(row=3, column=1, columnspan=2, sticky="w", pady=10)
        self.status = tk.StringVar(value="Ready")
        ttk.Label(frame, textvariable=self.status, wraplength=500).grid(row=4, column=0, columnspan=3, sticky="w", pady=12)
        buttons = ttk.Frame(frame)
        buttons.grid(row=5, column=0, columnspan=3, sticky="w", pady=(18, 0))
        self.sync_button = ttk.Button(buttons, text="Sync Now", command=self.sync_now)
        self.sync_button.pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="Save Configuration", command=self.save_configuration).pack(side="left", padx=8)
        ttk.Button(buttons, text="Close", command=self.root.destroy).pack(side="left", padx=8)

    def save_configuration(self, show_confirmation=True):
        try:
            interval = int(self.interval.get())
            if interval < 1 or interval > 24:
                raise ValueError("Synchronization interval must be between 1 and 24 hours.")
            self.cfg["tracking_sync_enabled"] = self.enabled.get()
            self.cfg["tracking_sync_interval_hours"] = interval
            save_config(self.cfg)
            if self.on_config_saved:
                self.on_config_saved()
            if show_confirmation:
                messagebox.showinfo(APP_NAME, "Synchronization configuration saved.", parent=self.root)
        except ValueError as exc:
            messagebox.showerror(APP_NAME, str(exc), parent=self.root)

    def sync_now(self):
        self.save_configuration(show_confirmation=False)
        self.sync_button.configure(state="disabled")
        self.status.set("Synchronization started…")

        def worker():
            try:
                debug_mode = self.cfg.get("FullSynchronizationDebug", True)
                result = tracking_sync_service().sync(self.cfg.get("tracking_last_sync_time", ""), debug_mode)
                if not debug_mode:
                    self.cfg["tracking_last_sync_time"] = result["last_sync_time"]
                save_config(self.cfg)
                self.root.after(0, self.sync_complete, result)
            except Exception as exc:
                self.root.after(0, self.sync_failed, str(exc))
        threading.Thread(target=worker, daemon=True).start()

    def sync_complete(self, result):
        self.sync_button.configure(state="normal")
        self.last_sync.set(result["last_sync_time"])
        text = ("Synchronization Complete\n\n"
                f"Records Downloaded: {result['records_downloaded']}\n"
                f"Rows Updated: {result['rows_updated']}\n"
                f"Execution Time: {result['execution_time']:.2f} seconds\n"
                f"Last Synchronization Time: {result['last_sync_time']}")
        self.status.set(text)
        if self.on_sync_complete_callback:
            self.on_sync_complete_callback()
        messagebox.showinfo(APP_NAME, text, parent=self.root)

    def sync_failed(self, error):
        self.sync_button.configure(state="normal")
        self.status.set(f"Synchronization failed: {error}")
        messagebox.showerror(APP_NAME, f"Tracking synchronization failed:\n\n{error}", parent=self.root)


class BounceTrackingWindow:
    def __init__(self, parent, on_check_complete=None):
        self.root = tk.Toplevel(parent); self.root.title("Bounce Tracking"); self.root.geometry("760x520")
        self.on_check_complete = on_check_complete
        self.cfg = load_config(); self.accounts = account_service(); self.service = bounce_tracking_service()
        frame = ttk.Frame(self.root, padding=18); frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Bounce Mail Tracking", font=("Segoe UI", 17, "bold")).pack(anchor="w", pady=(0, 14))
        self.tree = ttk.Treeview(frame, columns=("provider", "name", "email", "status"), show="headings", height=9)
        for key, title, width in (("provider", "Provider", 110), ("name", "Display Name", 180), ("email", "Email Address", 300), ("status", "Status", 90)):
            self.tree.heading(key, text=title); self.tree.column(key, width=width)
        self.tree.pack(fill="both", expand=True)
        for account in self.accounts.list_accounts(): self.tree.insert("", "end", values=(account["provider"], account["display_name"], account["email"], "Enabled" if account.get("enabled", True) else "Disabled"))
        options = ttk.LabelFrame(frame, text="Automatic Bounce Check Scheduler", padding=12); options.pack(fill="x", pady=12)
        self.enabled = tk.BooleanVar(value=self.cfg.get("bounce_check_enabled", False)); ttk.Checkbutton(options, text="Enable Automatic Bounce Check", variable=self.enabled).pack(side="left")
        ttk.Label(options, text="Every").pack(side="left", padx=(30, 6)); self.interval = tk.StringVar(value=str(self.cfg.get("bounce_check_interval_minutes", 30))); ttk.Spinbox(options, from_=1, to=1440, textvariable=self.interval, width=8).pack(side="left"); ttk.Label(options, text="minutes").pack(side="left", padx=6)
        self.status = tk.StringVar(value="Ready"); ttk.Label(frame, textvariable=self.status).pack(anchor="w", pady=6)
        bar = ttk.Frame(frame); bar.pack(fill="x", pady=8)
        self.check_button = ttk.Button(bar, text="Check Bounce Now", command=self.check_now); self.check_button.pack(side="left", padx=(0, 7))
        ttk.Button(bar, text="Test IMAP Connection", command=self.test_imap).pack(side="left", padx=7)
        ttk.Button(bar, text="Save Scheduler", command=self.save_scheduler).pack(side="left", padx=7)
        ttk.Button(bar, text="Close", command=self.root.destroy).pack(side="right")
    def selected_email(self):
        selected = self.tree.selection(); return self.tree.item(selected[0], "values")[2] if selected else None
    def save_scheduler(self, show_message=True):
        try:
            interval = int(self.interval.get())
            if interval < 1: raise ValueError("Interval must be at least 1 minute.")
            self.cfg["bounce_check_enabled"] = self.enabled.get(); self.cfg["bounce_check_interval_minutes"] = interval; save_config(self.cfg)
            if show_message: messagebox.showinfo(APP_NAME, "Bounce scheduler saved.", parent=self.root)
            return True
        except ValueError as exc: messagebox.showerror(APP_NAME, str(exc), parent=self.root); return False
    def configurations(self):
        return [value for account in self.accounts.list_accounts() if account.get("enabled", True) for value in [self.accounts.imap_configuration(account["email"])] if value]
    def check_now(self):
        if not self.save_scheduler(show_message=False): return
        self.check_button.configure(state="disabled"); self.status.set("Checking unread bounce emails…")
        def worker():
            detected = matched = 0; errors = []
            for configuration in self.configurations():
                try:
                    result = self.service.check_account(configuration); detected += result["detected"]; matched += result["matched"]
                except Exception as exc: errors.append(f"{configuration['email']}: {exc}")
            self.root.after(0, self.check_complete, detected, matched, errors)
        threading.Thread(target=worker, daemon=True).start()
    def check_complete(self, detected, matched, errors):
        self.check_button.configure(state="normal"); text = f"Bounce check complete. Detected: {detected}. Excel rows updated: {matched}."
        if errors: text += "\n\n" + "\n".join(errors)
        self.status.set(text)
        if self.on_check_complete: self.on_check_complete()
        (messagebox.showwarning if errors else messagebox.showinfo)(APP_NAME, text, parent=self.root)
    def test_imap(self):
        email_address = self.selected_email()
        if not email_address: messagebox.showwarning(APP_NAME, "Select a mail account first.", parent=self.root); return
        configuration = self.accounts.imap_configuration(email_address)
        if not configuration: messagebox.showerror(APP_NAME, "Mail account is disabled or not configured.", parent=self.root); return
        try: self.service.test_connection(configuration); messagebox.showinfo(APP_NAME, "IMAP connection successful.", parent=self.root)
        except Exception as exc: messagebox.showerror(APP_NAME, f"IMAP connection failed:\n\n{exc}", parent=self.root)


class EmailAutomationApp:
    def __init__(self):
        self.root=tk.Tk();self.root.title("Email Automation");self.root.geometry("1120x680");self.root.minsize(900,580)
        sidebar=tk.Frame(self.root,bg="#16324f",width=210);sidebar.pack(side="left",fill="y");sidebar.pack_propagate(False)
        tk.Label(sidebar,text="Email Automation",bg="#16324f",fg="white",font=("Segoe UI",16,"bold"),pady=24).pack(fill="x")
        actions=(("Dashboard",self.refresh),("Send Mail Now",lambda:SenderWindow(self.root)),("Daily Scheduling",lambda:MultiScheduleWindow(self.root)),("Mail Account Setup",lambda:AccountWindow(self.root)),("Tracking Synchronization",lambda:TrackingSynchronizationWindow(self.root,self.schedule_tracking_check,self.record_tracking_sync_run)),("Bounce Tracking",lambda:BounceTrackingWindow(self.root,self.record_bounce_check_run)),("Reports",lambda:ReportsWindow(self.root)),("Settings",lambda:SettingsWindow(self.root)),("Logs",lambda:LogWindow(self.root)),("Exit",self.root.destroy))
        for text,cmd in actions:tk.Button(sidebar,text=text,command=cmd,bg="#16324f",fg="white",activebackground="#285b87",activeforeground="white",relief="flat",anchor="w",padx=22,pady=11,font=("Segoe UI",10)).pack(fill="x")
        tk.Label(sidebar,text=f"Build:\n{BUILD_TIMESTAMP_UTC} UTC",bg="#16324f",fg="#b9cfe3",font=("Segoe UI",8),justify="left",padx=22,pady=12).pack(side="bottom",fill="x",anchor="w")
        self.main=ttk.Frame(self.root,padding=22);self.main.pack(side="left",fill="both",expand=True);self.sync_running=False;self.tracking_check_job=None;self.tracking_debug_last_run="";self.bounce_check_running=False;self.bounce_last_run=0.0;self.refresh();self.tracking_check_job=self.root.after(1500,self.check_tracking_synchronization);self.root.after(2500,self.check_automatic_bounces)

    def record_bounce_check_run(self): self.bounce_last_run = time.monotonic()

    def check_automatic_bounces(self):
        cfg = load_config(); interval = max(1, int(cfg.get("bounce_check_interval_minutes", 30))) * 60
        due = not self.bounce_last_run or time.monotonic() - self.bounce_last_run >= interval
        if cfg.get("bounce_check_enabled", False) and due and not self.bounce_check_running:
            self.bounce_check_running = True
            def worker():
                try:
                    accounts = account_service(); service = bounce_tracking_service()
                    for account in accounts.list_accounts():
                        configuration = accounts.imap_configuration(account["email"]) if account.get("enabled", True) else None
                        if configuration:
                            try: service.check_account(configuration)
                            except Exception: pass
                finally:
                    self.record_bounce_check_run(); self.bounce_check_running = False
            threading.Thread(target=worker, daemon=True).start()
        self.root.after(60000, self.check_automatic_bounces)

    def record_tracking_sync_run(self):
        self.tracking_debug_last_run = datetime.now(timezone.utc).isoformat()

    def schedule_tracking_check(self):
        if self.tracking_check_job:
            self.root.after_cancel(self.tracking_check_job)
        self.tracking_check_job = self.root.after(100, self.check_tracking_synchronization)

    def check_tracking_synchronization(self):
        cfg = load_config()
        debug_mode = cfg.get("FullSynchronizationDebug", True)
        scheduling_cursor = self.tracking_debug_last_run if debug_mode else cfg.get("tracking_last_sync_time", "")
        due = synchronization_is_due(cfg.get("tracking_sync_enabled", False), cfg.get("tracking_sync_interval_hours", 5), scheduling_cursor)
        if due and not self.sync_running:
            self.sync_running = True
            last_sync = cfg.get("tracking_last_sync_time", "")
            def worker():
                try:
                    current = load_config(); debug_mode = current.get("FullSynchronizationDebug", True)
                    result = tracking_sync_service().sync(last_sync, debug_mode)
                    if not debug_mode:
                        updated = load_config(); updated["tracking_last_sync_time"] = result["last_sync_time"]; save_config(updated)
                    else:
                        self.record_tracking_sync_run()
                except Exception:
                    pass
                finally:
                    self.sync_running = False
            threading.Thread(target=worker, daemon=True).start()
        self.tracking_check_job = self.root.after(60000, self.check_tracking_synchronization)
    def refresh(self):
        for child in self.main.winfo_children():child.destroy()
        ttk.Label(self.main,text="Dashboard",font=("Segoe UI",20,"bold")).pack(anchor="w")
        accounts=account_service().list_accounts();rows,counts=email_grid_data();enabled=sum(1 for a in accounts if a.get("enabled",True));today_sent=today_failed=weekly_sent=monthly_sent=0;per={}
        p=app_paths()["list"]
        if p.exists():
            wb=load_workbook(p,read_only=True,data_only=True);ws=wb.active;h=header_map(ws)
            for r in range(2,ws.max_row+1):
                status=str(ws.cell(r,h.get("status",0)).value or "");sender_key=h.get("sender_email") or h.get("sender_mail");sender=str(ws.cell(r,sender_key).value or "") if sender_key else "";entry=per.setdefault(sender,{"sent_today":0,"pending":0})
                if status.lower()=="pending":entry["pending"]+=1
                date=ws.cell(r,h.get("sentdate",0)).value if h.get("sentdate") else None
                now=datetime.now();is_today=isinstance(date,datetime) and date.date()==now.date();is_week=isinstance(date,datetime) and 0 <= (now.date()-date.date()).days < 7;is_month=isinstance(date,datetime) and date.year==now.year and date.month==now.month
                if status.lower()=="sent":entry["sent_today"]+=int(is_today);today_sent+=int(is_today);weekly_sent+=int(is_week);monthly_sent+=int(is_month)
                if status.lower()=="failed" and is_today:today_failed+=1
            wb.close()
        total_done=counts["sent"]+counts["failed"];success=(counts["sent"]/total_done*100) if total_done else 0
        metrics=(("Total Gmail Accounts",len(accounts)),("Enabled Gmail Accounts",enabled),("Disabled Gmail Accounts",len(accounts)-enabled),("Grand Total Emails",counts["total"]),("Pending Emails",counts["pending"]),("Sent Emails",counts["sent"]),("Failed Emails",counts["failed"]),("Today's Sent",today_sent),("Today's Failed",today_failed),("Weekly Sent",weekly_sent),("Monthly Sent",monthly_sent),("Success Rate",f"{success:.1f}%"),("Failure Rate",f"{100-success:.1f}%"))
        cards=ttk.Frame(self.main);cards.pack(fill="x",pady=18)
        for i,(label,value) in enumerate(metrics):
            box=ttk.LabelFrame(cards,text=label,padding=12);box.grid(row=i//4,column=i%4,sticky="nsew",padx=5,pady=5);ttk.Label(box,text=str(value),font=("Segoe UI",16,"bold")).pack();cards.columnconfigure(i%4,weight=1)
        ttk.Label(self.main,text="Per Sender Email Statistics",font=("Segoe UI",13,"bold")).pack(anchor="w",pady=(12,5));tree=ttk.Treeview(self.main,columns=("sender","sent","pending"),show="headings",height=8)
        for k,t in (("sender","Sender Email"),("sent","Sent Today"),("pending","Pending")):tree.heading(k,text=t)
        for sender,value in per.items():tree.insert("","end",values=(sender,value["sent_today"],value["pending"]))
        tree.pack(fill="both",expand=True);ttk.Button(self.main,text="Refresh",command=self.refresh).pack(anchor="e",pady=8)
    def run(self):self.root.mainloop()


def verify_app():
    root = tk.Tk(); root.withdraw(); paths = app_paths(); desktop = Path.home() / "Desktop"
    checks = []
    try: credentials(); checks.append(("SMTP configuration", True, "Credentials found"))
    except Exception as exc: checks.append(("SMTP configuration", False, str(exc)))
    checks.append(("Excel file", paths["list"].exists(), str(paths["list"])))
    shortcuts = [desktop / "Send Pending Emails Now.lnk", desktop / "Configure Daily Email Schedule.lnk", desktop / "Verify Installation.lnk"]
    checks.append(("Desktop shortcuts", all(x.exists() for x in shortcuts), str(desktop)))
    task = task_query(); checks.append(("Scheduled task", task["exists"], "Enabled" if task["enabled"] else "Disabled or missing"))
    checks.append(("Log folder", paths["logs"].is_dir(), str(paths["logs"])))
    checks.append(("Config files", config_path().exists() and paths["env"].exists(), str(config_path())))
    text = "INSTALLATION VERIFICATION\n\n" + "\n".join(f"{'PASS' if ok else 'FAIL'}  {name}\n      {detail}" for name, ok, detail in checks)
    ensure_dirs(); report = paths["reports"] / f"verification_{datetime.now():%Y%m%d_%H%M%S}.txt"; report.write_text(text, encoding="utf-8")
    messagebox.showinfo("Verify Installation", text + f"\n\nReport saved to:\n{report}"); root.destroy()


def sample_workbook(path):
    wb = Workbook(); ws = wb.active; ws.title = "Mail List"
    headers = ["First_Name", "Last_Name", "Email", "Company", "Designation", "Country", "Subject", "Body", "Sender_Name", "Sender_Email", "Status", "Result", "SentDate"]
    ws.append(headers)
    ws.append(["John", "Doe", "recipient@example.com", "Example Company", "Manager", "USA", "Opportunity for {{Company}}", "Hi {{First_Name}} {{Last_Name}},\n\nA message for {{Company}}.", "Power Soft", "sales@gmail.com", "Pending", "", ""])
    ws.freeze_panes = "A2"; ws.auto_filter.ref = "A1:M2"
    widths = (16, 16, 30, 24, 24, 16, 36, 60, 22, 30, 14, 40, 22)
    for column, width in enumerate(widths, 1): ws.column_dimensions[ws.cell(1, column).column_letter].width = width
    wb.save(path)


def main():
    name = Path(sys.executable if getattr(sys, "frozen", False) else sys.argv[0]).stem.lower()
    if "--run-schedule" in sys.argv:
        try:
            schedule_id=sys.argv[sys.argv.index("--run-schedule")+1];item=next(x for x in load_schedules() if x["id"]==schedule_id and x.get("enabled",True));send_pending(int(item["max_emails"]),True,excel_path=item["excel_file"],attachment_ids=item.get("attachment_ids",[]))
        except Exception as exc:daily_log(status="Scheduled Run Error",error=str(exc))
        return
    if name == "email automation" or name == "emailautomation":
        migrate_legacy_account();startup_log();EmailAutomationApp().run();return
    if "setup" in name: setup_app()
    elif "configure" in name: ScheduleWindow().run()
    elif "verify" in name: verify_app()
    elif "--scheduled" in sys.argv:
        try: send_pending(int(load_config()["daily_limit"]), True)
        except Exception as exc: daily_log(status="Scheduled Run Error", error=str(exc))
    else: manual_send()


if __name__ == "__main__":
    main()
