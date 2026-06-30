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
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from dotenv import dotenv_values
from openpyxl import Workbook, load_workbook
from account_service import AccountService
from placeholder_service import PlaceholderService


APP_NAME = "Email Automation"
TASK_NAME = "The Power People Daily Email Automation"
DEFAULT_DATA_DIR = r"F:\CODEX\Email_automation"
DEFAULT_CONFIG = {"daily_limit": 5, "schedule_time": "09:00", "data_dir": DEFAULT_DATA_DIR, "default_sender_name": "The Power People", "random_delay_min": 5, "random_delay_max": 15, "retry_count": 1, "backup_enabled": True, "log_retention_days": 90, "theme": "Light"}


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


def account_service():
    return AccountService(app_paths()["accounts"])


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


def send_pending(limit, wait_between=True, progress=None, excel_path=None):
    paths = app_paths()
    if excel_path:
        paths["list"] = Path(excel_path)
    if limit < 1:
        raise ValueError("Email count must be at least 1.")
    if not paths["list"].exists():
        raise FileNotFoundError(f"Excel file not found: {paths['list']}")
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
            smtp_values = accounts.smtp_credentials(sender_email)
            if not smtp_values:
                raise ValueError("Sender Account Not Configured")
            if not validate_recipient_email(email):
                raise ValueError("Invalid email address")
            if not subject:
                raise ValueError("Subject is empty")
            if not body:
                raise ValueError("Body is empty")
            tracking_id = str(uuid.uuid4())
            write_tracking_id(wb, ws, paths["list"], row, tracking_column, tracking_id)
            msg = EmailMessage()
            msg["From"] = f"{sender_name} <{sender_email}>"
            msg["To"] = email
            msg["Subject"] = subject
            msg.set_content(body)
            pixel_url = f"https://emailtrackingserver.onrender.com/email/open/{tracking_id}"
            html_body = build_click_tracked_html(body, tracking_id)
            msg.add_alternative(f'<html><body>{html_body}<img src="{pixel_url}" width="1" height="1" style="display:none;" alt=""></body></html>', subtype="html")
            if active_sender != sender_email or smtp is None:
                if smtp is not None:
                    try: smtp.quit()
                    except Exception: pass
                smtp = smtplib.SMTP("smtp.gmail.com", 587, timeout=45)
                smtp.ehlo()
                smtp.starttls()
                smtp.ehlo()
                smtp.login(*smtp_values)
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

        def progress(number, total, result, email):
            if result == "Sent":
                self.batch_sent += 1
            else:
                self.batch_failed += 1
            self.root.after(0, self.on_progress, number, total, result, email)

        def worker():
            try:
                result = send_pending(value, True, progress, self.excel_file.get())
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


class AccountWindow:
    def __init__(self, parent=None):
        self.root = tk.Toplevel(parent) if parent else tk.Tk()
        self.root.title("Mail Account Setup")
        self.root.geometry("720x430")
        self.service = account_service()
        frame = ttk.Frame(self.root, padding=16); frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Gmail Account Management", font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(0, 12))
        self.tree = ttk.Treeview(frame, columns=("name", "email", "status"), show="headings")
        for key, title, width in (("name", "Account Name", 200), ("email", "Gmail Address", 300), ("status", "Status", 100)):
            self.tree.heading(key, text=title); self.tree.column(key, width=width)
        self.tree.pack(fill="both", expand=True)
        bar = ttk.Frame(frame); bar.pack(fill="x", pady=(12, 0))
        for text, command in (("Add", self.add), ("Edit", self.edit), ("Delete", self.delete), ("Enable", lambda: self.toggle(True)), ("Disable", lambda: self.toggle(False)), ("Test SMTP", self.test)):
            ttk.Button(bar, text=text, command=command).pack(side="left", padx=(0, 7))
        ttk.Button(bar, text="Close", command=self.root.destroy).pack(side="right")
        self.refresh()
    def selected(self):
        selected = self.tree.selection()
        return self.tree.item(selected[0], "values") if selected else None
    def refresh(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for account in self.service.list_accounts():
            self.tree.insert("", "end", values=(account["name"], account["email"], "Enabled" if account.get("enabled", True) else "Disabled"))
    def account_form(self, existing=None):
        name = simpledialog.askstring(APP_NAME, "Account Name:", initialvalue=existing[0] if existing else "", parent=self.root)
        if name is None: return
        email = simpledialog.askstring(APP_NAME, "Gmail Address:", initialvalue=existing[1] if existing else "", parent=self.root)
        if email is None: return
        prompt = "New Google App Password (leave empty to keep existing):" if existing else "Google App Password:"
        password = simpledialog.askstring(APP_NAME, prompt, show="*", parent=self.root)
        if password is None: return
        self.service.save_account(name, email, password, True, existing[1] if existing else None)
        self.refresh()
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
        if value and messagebox.askyesno(APP_NAME, f"Delete {value[1]} permanently?", parent=self.root):
            try: self.service.delete_account(value[1]); self.refresh()
            except Exception as exc: messagebox.showerror(APP_NAME, str(exc), parent=self.root)
    def toggle(self, enabled):
        value = self.selected()
        if value: self.service.set_enabled(value[1], enabled); self.refresh()
    def test(self):
        value = self.selected()
        if not value: return
        try: self.service.test_smtp(value[1]); messagebox.showinfo(APP_NAME, "SMTP connection successful.", parent=self.root)
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
        self.window.geometry("620x560")
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
        ttk.Label(frame, text="Time (HH:mm):").grid(row=4, column=0, sticky="w", pady=7)
        ttk.Entry(frame, textvariable=self.when, width=16).grid(row=4, column=1, sticky="w", pady=7)
        ttk.Label(frame, text="Maximum Emails:").grid(row=5, column=0, sticky="w", pady=7)
        ttk.Entry(frame, textvariable=self.maximum, width=16).grid(row=5, column=1, sticky="w", pady=7)
        ttk.Checkbutton(frame, text="Schedule Enabled", variable=self.enabled).grid(row=6, column=1, sticky="w", pady=10)
        buttons = ttk.Frame(frame)
        buttons.grid(row=7, column=0, columnspan=3, sticky="e", pady=(18, 0))
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


class SettingsWindow:
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


class EmailAutomationApp:
    def __init__(self):
        self.root=tk.Tk();self.root.title("Email Automation");self.root.geometry("1120x680");self.root.minsize(900,580)
        sidebar=tk.Frame(self.root,bg="#16324f",width=210);sidebar.pack(side="left",fill="y");sidebar.pack_propagate(False)
        tk.Label(sidebar,text="Email Automation",bg="#16324f",fg="white",font=("Segoe UI",16,"bold"),pady=24).pack(fill="x")
        actions=(("Dashboard",self.refresh),("Send Mail Now",lambda:SenderWindow(self.root)),("Daily Scheduling",lambda:MultiScheduleWindow(self.root)),("Mail Account Setup",lambda:AccountWindow(self.root)),("Reports",lambda:ReportsWindow(self.root)),("Settings",lambda:SettingsWindow(self.root)),("Logs",lambda:LogWindow(self.root)),("Exit",self.root.destroy))
        for text,cmd in actions:tk.Button(sidebar,text=text,command=cmd,bg="#16324f",fg="white",activebackground="#285b87",activeforeground="white",relief="flat",anchor="w",padx=22,pady=11,font=("Segoe UI",10)).pack(fill="x")
        self.main=ttk.Frame(self.root,padding=22);self.main.pack(side="left",fill="both",expand=True);self.refresh()
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
            schedule_id=sys.argv[sys.argv.index("--run-schedule")+1];item=next(x for x in load_schedules() if x["id"]==schedule_id and x.get("enabled",True));send_pending(int(item["max_emails"]),True,excel_path=item["excel_file"])
        except Exception as exc:daily_log(status="Scheduled Run Error",error=str(exc))
        return
    if name == "email automation" or name == "emailautomation":
        migrate_legacy_account();EmailAutomationApp().run();return
    if "setup" in name: setup_app()
    elif "configure" in name: ScheduleWindow().run()
    elif "verify" in name: verify_app()
    elif "--scheduled" in sys.argv:
        try: send_pending(int(load_config()["daily_limit"]), True)
        except Exception as exc: daily_log(status="Scheduled Run Error", error=str(exc))
    else: manual_send()


if __name__ == "__main__":
    main()
