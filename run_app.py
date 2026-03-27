# --- SOD/EOD REPORT GENERATOR ---
#
# Author: Allen Paul Olguera (with Gemini)
# Last Updated: 2026-02-06
#
# --- KEY FEATURES ---
# 1.  SOD Tab:
#     - "Load Preset": Loads tasks for today (Daily, Weekday, Monthly).
#     - "Load Unfinished Tasks": Loads tasks from previous EOD. If task was
#       "In Progress" or "Carried Over", that status is added to the Remarks.
#     - Tasklist Input: Includes a helper dropdown below the text field.
#       Populate this dropdown in the "Edit Daily Presets" menu.
#
# 2.  Settings:
#     - "Daily Logic": Independent setting.
#       - Delayed: Uses (Today - 1) for DdddYYYY format.
#       - Current: Uses Today for DdddYYYY format.
#     - "Weekly Logic": Independent setting.
#       - Delayed: Uses (ISO Week - 1) for W{WW}{YYYY} format.
#       - Current: Uses ISO Week for W{WW}{YYYY} format.
#     - "Monthly Logic": Independent setting.
#       - Current (Month)
#       - Delayed (Month-1)
#     - "Monthly Start Day": day of month that flips the effective month.
#
# 3. Work Details Table Updates:
#     - "Remarks" -> "Work Schedule" (START_TIME - END_TIME)
#     - "From" -> "SOD Created" (system time when SOD draft is created)
#     - "To" -> "EOD Created" (system time when EOD draft is generated)
#     - "OT Classification" -> "Work Location"
#     - "OT Criteria" removed
#     - "Reason" -> "Stream"
#
# --- END OF SUMMARY ---

import tkinter as tk
from tkinter import filedialog, messagebox

# --- UI THEME: ttkbootstrap (Corporate Clean) ---
# Alias 'ttk' to ttkbootstrap so existing ttk.* calls continue to work.
try:
    import ttkbootstrap as tb
    from ttkbootstrap.constants import *
    ttk = tb
except Exception:
    tb = None
    from tkinter import ttk  # fallback

# Base window depends on whether ttkbootstrap is available
BaseWindow = tb.Window if tb else tk.Tk
import os
import datetime
import json
import re
import urllib.request
import urllib.error
import webbrowser
import win32com.client as win32
import win32clipboard
import win32con
import calendar
from dateutil.relativedelta import relativedelta

try:
    from packaging.version import parse as parse_version
except Exception:
    parse_version = None

try:
    from PIL import ImageGrab, Image
except ImportError:
    messagebox.showerror(
        "Missing Library",
        "The 'Pillow' library is required.\n\nPlease install it by running:\n\npip install Pillow"
    )
    raise SystemExit


# --- CONFIGURATION ---
APP_NAME = 'Prod_task_generator'
APP_VERSION = '2026.03.27'
HARDCODED_UPDATE_MANIFEST_URL = 'https://api.github.com/repos/mrallen29/iqvia-daily-task-generator/releases/latest'
RESOURCES_DIR = 'resources'
CONFIG_FILE = os.path.join(RESOURCES_DIR, 'config.json')
PRESETS_FILE = os.path.join(RESOURCES_DIR, 'presets.json')
TEMP_SCREENSHOT_PATH = os.path.join(RESOURCES_DIR, 'temp_screenshot.png')

# OT data files
OT_IN_FILE_PREFIX = os.path.join(RESOURCES_DIR, 'ot_in_')
OT_OUT_FILE_PREFIX = os.path.join(RESOURCES_DIR, 'ot_out_')

# Weekly start day options (Mon-Sun)
WEEK_START_DAY_OPTIONS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
WEEKDAY_TO_INDEX = {'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3, 'Friday': 4, 'Saturday': 5, 'Sunday': 6}


DEFAULT_CONFIG = {
    # First-run defaults (used when resources/config.json does not exist)
    'YOUR_NAME': '',
    'SIGNATURE_NAME': '',

    # Recipients (blank by default)
    'RECIPIENTS_TO': '',
    'RECIPIENTS_CC': '',
    'RECIPIENTS_TO_NEW': '',
    'RECIPIENTS_CC_NEW': '',

    # Fixed fields (blank by default)
    'FIXED_COUNTRY': '',
    'FIXED_STREAM': '',

    # Schedule (blank by default)
    'START_TIME': '',
    'END_TIME': '',

    # General / Period Logic defaults
    'DAILY_LOGIC': 'Current (Day)',
    'WEEKLY_LOGIC': 'Current (ISO)',
    'WEEKLY_START_OFFSET_DAYS': 0,  # 0=Mon..6=Sun
    'MONTHLY_LOGIC': 'Current (Month)',
    'MONTHLY_START_DAY': 1,
    'WEEK_START_DAY': 'Monday',

    # Outlook workflow
    'OUTLOOK_VERSION': 'Classic',
    'NEW_OUTLOOK_SIGNATURE_DISABLED': False,

    # Updates
    'UPDATE_MANIFEST_URL': '',
    'AUTO_CHECK_UPDATES': False,
    'SKIPPED_UPDATE_VERSION': '',
    'LAST_UPDATE_CHECK': '',
    'LAST_UPDATE_VERSION_SEEN': '',
}

DAYS_TO_KEEP_DATA = 90


def clean_version_string(value: str) -> str:
    s = str(value or '').strip()
    if s.lower().startswith('v'):
        s = s[1:]
    return s


def compare_version_strings(current_version: str, latest_version: str) -> int:
    cur = clean_version_string(current_version)
    lat = clean_version_string(latest_version)
    if not cur and not lat:
        return 0
    if not cur:
        return -1
    if not lat:
        return 1
    if parse_version is not None:
        try:
            cur_v = parse_version(cur)
            lat_v = parse_version(lat)
            if cur_v < lat_v:
                return -1
            if cur_v > lat_v:
                return 1
            return 0
        except Exception:
            pass

    def _parts(v: str):
        vals = []
        for part in re.split(r'[^0-9]+', v):
            if part:
                vals.append(int(part))
        return tuple(vals or [0])

    cp = _parts(cur)
    lp = _parts(lat)
    if cp < lp:
        return -1
    if cp > lp:
        return 1
    return 0


def normalize_update_payload(payload: dict) -> dict:
    if not isinstance(payload, dict):
        raise ValueError('Update manifest must be a JSON object.')

    version = payload.get('latest_version') or payload.get('version') or payload.get('tag_name') or payload.get('name') or ''
    download_url = payload.get('download_url') or payload.get('download_page_url') or payload.get('release_page_url') or payload.get('html_url') or ''
    assets = payload.get('assets') or []
    if not download_url and isinstance(assets, list):
        preferred = []
        others = []
        for asset in assets:
            if not isinstance(asset, dict):
                continue
            candidate = asset.get('browser_download_url') or asset.get('url') or ''
            if not candidate:
                continue
            name = str(asset.get('name') or '').lower()
            if name.endswith('.exe') or name.endswith('.msi') or name.endswith('.zip'):
                preferred.append(candidate)
            else:
                others.append(candidate)
        all_candidates = preferred + others
        if all_candidates:
            download_url = all_candidates[0]

    release_notes = payload.get('release_notes') or payload.get('notes') or payload.get('body') or ''
    published_at = payload.get('published_at') or payload.get('created_at') or ''
    return {
        'latest_version': clean_version_string(version),
        'download_url': str(download_url or '').strip(),
        'release_notes': str(release_notes or '').strip(),
        'mandatory': bool(payload.get('mandatory', False)),
        'published_at': str(published_at or '').strip(),
        'raw': payload,
    }


# --- HELPER FUNCTIONS ---
def calculate_total_hours(start_str, end_str):
    try:
        time_format = "%I:%M%p"
        start_time = datetime.datetime.strptime(start_str, time_format)
        end_time = datetime.datetime.strptime(end_str, time_format)
        if end_time < start_time:
            end_time += datetime.timedelta(days=1)
        duration = end_time - start_time
        hours = duration.total_seconds() / 3600
        return f"{hours:.2f}HRS"
    except (ValueError, TypeError):
        return "(calc err)"


# --- OT STATUS HELPERS ---
def _ot_status_state_from_any(text_value: str) -> str:
    """Return 'done' or 'in_progress' from any OT status display string (single or dual)."""
    s = (text_value or '').strip()
    if not s:
        return 'in_progress'
    # dual checkbox style
    if '☑' in s:
        chk = s.find('☑')
        tail = s[chk:]
        if 'Done' in tail and ('In Progress' not in tail or tail.find('Done') < tail.find('In Progress')):
            return 'done'
        return 'in_progress'
    # fallback
    if 'Done' in s or '🟩' in s:
        return 'done'
    return 'in_progress'

def ot_status_dual(state: str) -> str:
    """Dual display for UI: shows both options with one ticked."""
    st = (state or '').strip().lower()
    if st == 'done':
        return '☐ 🟨 In Progress ☑ 🟩 Done'
    return '☑ 🟨 In Progress ☐ 🟩 Done'

def ot_status_single(state: str) -> str:
    """Single label for email (Option A): only the selected status."""
    st = (state or '').strip().lower()
    return '🟩 Done' if st == 'done' else '🟨 In Progress'


def format_time_display(time_str):
    """Converts '02:00PM' -> '02:00 PM' for nicer display. If cannot parse, returns original."""
    if not time_str:
        return ""
    try:
        t = datetime.datetime.strptime(time_str.strip(), "%I:%M%p")
        return t.strftime("%I:%M %p")
    except Exception:
        return time_str


def now_display_time():
    """Return current system time in 'HH:MM AM/PM' format."""
    return datetime.datetime.now().strftime("%I:%M %p")


def work_schedule_display(config):
    """Return 'START - END' schedule from config with spaces before AM/PM."""
    start = format_time_display(config.get('START_TIME', ''))
    end = format_time_display(config.get('END_TIME', ''))
    if start and end:
        return f"{start} - {end}"
    return f"{start}{(' - ' + end) if end else ''}".strip()




def work_schedule_display_from_times(start_time_str, end_time_str):
    """Return 'START - END' schedule from raw time strings like '02:00PM'."""
    start = format_time_display(start_time_str)
    end = format_time_display(end_time_str)
    if start and end:
        return f"{start} - {end}"
    return f"{start}{(' - ' + end) if end else ''}".strip()
def clamp_monthly_start_day(year, month, start_day):
    """Clamp start_day to the number of days in the month."""
    try:
        _, max_day = calendar.monthrange(year, month)
        return max(1, min(int(start_day), max_day))
    except Exception:
        return 1




def most_recent_weekday(date_obj: datetime.date, weekday_index: int) -> datetime.date:
    """Return the most recent date <= date_obj that matches weekday_index (Mon=0..Sun=6)."""
    delta = (date_obj.weekday() - weekday_index) % 7
    return date_obj - datetime.timedelta(days=delta)


def yy_from_year(year: int) -> str:
    """Return full year as 4 digits (YYYY) as string."""
    return f"{int(year):04d}"

# Frequency parsing helpers (supports multi-frequency like "Weekly, Monthly")
FREQUENCY_ORDER = ['Daily', 'Weekly', 'Monthly']

def normalize_frequency_string(freq_str: str):
    """Return (freq_list, display_str). Accepts separators: ',', '+', '&', '|', '/'."""
    raw = (freq_str or '').strip()
    if not raw:
        return [], ''
    tmp = raw
    for sep in ['+', '&', '|', '/']:
        tmp = tmp.replace(sep, ',')
    parts = [p.strip() for p in tmp.split(',') if p.strip()]
    mapped = []
    for p in parts:
        key = p.replace(' ', '').upper()
        if key in ('D','DAILY'):
            mapped.append('Daily')
        elif key in ('W','WEEKLY'):
            mapped.append('Weekly')
        elif key in ('M','MONTHLY'):
            mapped.append('Monthly')
        else:
            mapped.append(p.title())
    seen=set(); dedup=[]
    for item in mapped:
        if item not in seen:
            seen.add(item); dedup.append(item)
    ordered = [f for f in FREQUENCY_ORDER if f in dedup] + [f for f in dedup if f not in FREQUENCY_ORDER]
    return ordered, ', '.join(ordered)

def ot_date_display(date_obj: datetime.date) -> str:
    """Return OT date format as dd/mm/yyyy."""
    return date_obj.strftime('%d/%m/%Y')


def infer_shift_date_from_config(now_dt: datetime.datetime, config: dict) -> datetime.date:
    """Infer shift date using the same overnight rule used by EOD."""
    start_time_str = str(config.get('START_TIME', '') or '').strip()
    try:
        start_clock = datetime.datetime.strptime(start_time_str, '%I:%M%p').time() if start_time_str else None
    except Exception:
        start_clock = None
    starts_pm = start_time_str.upper().endswith('PM')
    if starts_pm and start_clock and now_dt.time() < start_clock:
        return now_dt.date() - datetime.timedelta(days=1)
    return now_dt.date()

# --- CORE EMAIL LOGIC ---
def create_sod_html_body(tasks_data, config, sod_created_time_str, actual_start_shift=None):
    today_str = datetime.date.today().strftime('%d/%m/%Y')

    task_table_html = "<p>Please see below tasks list:</p><table><thead><tr>"
    headers = ["Country", "Responsible", "Stream", "Tasklist", "Offsite Schedule", "Frequency", "Period", "Start Time", "End Time", "Issue Encountered", "Status", "Remarks"]
    for header in headers:
        task_table_html += f"<th>{header}</th>"
    task_table_html += "</tr></thead><tbody>"
    for task_row in tasks_data:
        task_table_html += "<tr>"
        for item in task_row:
            task_table_html += f"<td>{item}</td>"
        # pad missing cells so Status and Remarks are separate columns
        if len(task_row) < len(headers):
            for _ in range(len(headers) - len(task_row)):
                task_table_html += "<td></td>"
        task_table_html += "</tr>"
    task_table_html += "</tbody></table><br>"

    # Work Schedule display rule (SOD):
    # - If Actual Start Shift differs from Settings START_TIME, show: 'ACTUAL_START - (tbd)'
    # - Otherwise, show scheduled: 'START_TIME - END_TIME'
    scheduled_start_raw = (config.get('START_TIME') or '').strip()
    actual_start_raw = (actual_start_shift or '').strip()
    def _norm_time_key(t: str) -> str:
        return (t or '').replace(' ', '').upper()
    if actual_start_raw and _norm_time_key(actual_start_raw) != _norm_time_key(scheduled_start_raw):
        schedule_str = f"{format_time_display(actual_start_raw)} - (tbd)"
    else:
        schedule_str = work_schedule_display(config)
    eod_created_str = "(tbd)"

    return f"""
    <html>
      <head>
        <style>
          body{{font-family:Calibri,sans-serif;font-size:11pt}}
          table{{border-collapse:collapse;width:100%}}
          th,td{{border:1px solid #B2B2B2;padding:8px;text-align:left;font-size:10pt}}
          th{{background-color:#DDEBF7}}
        </style>
      </head>
      <body>
<p>Hi Everyone,</p>
 <p>I will now start my shift today.</p>
        <p>Please see my below WFH details.</p>
        <p><b>Work Details:</b></p>

        <table>
          <thead>
            <tr>
              <th>Workday</th>
              <th>Name</th>
              <th>Work Schedule</th>
              <th>SOD Created</th>
              <th>EOD Created</th>
              <th>Total</th>
              <th>Work Location</th>
              <th>Reason</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>{today_str}</td>
              <td>{config['YOUR_NAME']}</td>
              <td>{schedule_str}</td>
              <td>{sod_created_time_str}</td>
              <td>{eod_created_str}</td>
              <td>(tbd)</td>
              <td>WFH</td>
              <td>{config['FIXED_STREAM']}</td>
            </tr>
          </tbody>
        </table>
        <br>

        {task_table_html}

        <br>
        {build_signature_html(config)}
      </body>
    </html>
    """


def create_eod_html_body(tasks_data, config, actual_end_time, shift_date, include_screenshot,
                         sod_created_time_str, eod_created_time_str, actual_start_shift=None):
    shift_date_str = shift_date.strftime('%d/%m/%Y')
    start_time_for_total = actual_start_shift or config.get('START_TIME')
    end_time_for_total = actual_end_time if actual_end_time else config.get('END_TIME')
    total_hours = calculate_total_hours(start_time_for_total, end_time_for_total) if (start_time_for_total and end_time_for_total) else "(tbd)"
    schedule_str = work_schedule_display_from_times(start_time_for_total or '', end_time_for_total or '')

    wfh_details_html = f"""
    <p><b>Work Details:</b></p>
    <table>
      <thead>
        <tr>
          <th>Workday</th>
          <th>Name</th>
          <th>Work Schedule</th>
          <th>SOD Created</th>
          <th>EOD Created</th>
          <th>Total</th>
          <th>Work Location</th>
          <th>Reason</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td>{shift_date_str}</td>
          <td>{config['YOUR_NAME']}</td>
          <td>{schedule_str}</td>
          <td>{sod_created_time_str}</td>
          <td>{eod_created_time_str}</td>
          <td>{total_hours}</td>
          <td>WFH</td>
          <td>{config['FIXED_STREAM']}</td>
        </tr>
      </tbody>
    </table>
    <br>
    """

    task_table_html = ""
    if tasks_data:
        task_table_html = "<p><b>Task Status Report:</b></p><table><thead><tr>"
        headers = ["Country", "Responsible", "Stream", "Tasklist", "Offsite Schedule", "Frequency", "Period", "Start Time", "End Time", "Issue Encountered", "Status", "Remarks"]
        for header in headers:
            task_table_html += f"<th>{header}</th>"
        task_table_html += "</tr></thead><tbody>"
        for task_row in tasks_data:
            task_table_html += "<tr>"
            for item in task_row[:len(headers)]:
                task_table_html += f"<td>{item}</td>"
            # pad missing cells (ensures Remarks column has its own bordered cell)
            if len(task_row) < len(headers):
                for _ in range(len(headers) - len(task_row)):
                    task_table_html += "<td></td>"
            task_table_html += "</tr>"
        task_table_html += "</tbody></table><br>"

    screenshot_html = ""
    if include_screenshot:
        screenshot_html = """
        <p><b>Attendance Log:</b></p>
        <img src="cid:biologs_screenshot" style="width:100%; max-width:100%;">
        <br>
        """

    return f"""
    <html>
      <head>
        <style>
          body{{font-family:Calibri,sans-serif;font-size:11pt}}
          table{{border-collapse:collapse;width:100%}}
          th,td{{border:1px solid #B2B2B2;padding:8px;text-align:left;font-size:10pt}}
          th{{background-color:#DDEBF7}}
        </style>
      </head>
      <body>
        <p>Hi Everyone,</p>
        <p>I will end my WFH for today, please see below details for reference:</p>
        {wfh_details_html}
        {task_table_html}
        {screenshot_html}
        {build_signature_html(config)}
      </body>
    </html>
    """


def create_ot_in_html_body(ot_tasks_data, config, shift_date, ot_from_str, ot_to_str, total_hours_str, justification_text):
    date_str = ot_date_display(shift_date)
    day_str = shift_date.strftime('%A')
    to_display = ot_to_str if ot_to_str else '(tbd)'
    total_display = total_hours_str if total_hours_str else '(tbd)'
    justification = (justification_text or '').strip()

    ot_details_html = f"""
    <p>Hi Everyone,</p>
    <p>I will now start my OT, see below tasklist.</p>
    <table><thead><tr>
      <th>Date<br>(dd/mm/yyyy)</th><th>Day</th><th>Name</th><th>Stream</th><th>From</th><th>To</th><th>Total Hrs</th><th>Actual Hours</th><th>Comment<br>(Filling / Offset)</th><th>Justification</th>
    </tr></thead><tbody><tr>
      <td>{date_str}</td><td>{day_str}</td><td>{config.get('YOUR_NAME','')}</td><td></td><td>{ot_from_str}</td><td>{to_display}</td><td>{total_display}</td><td></td><td></td><td>{justification}</td>
    </tr></tbody></table><br>
    """

    headers = ['Country','Responsible','Stream','Task List','Offsite Schedule','Frequency','Period','Start Time','End Time','Status','Issue Encountered?','Remarks']
    task_table_html = ''
    if ot_tasks_data:
        task_table_html = "<table><thead><tr>" + "".join([f"<th>{h}</th>" for h in headers]) + "</tr></thead><tbody>"
        for row in ot_tasks_data:
            task_table_html += "<tr>" + "".join([f"<td>{cell}</td>" for cell in row[:len(headers)]]) + "</tr>"
        task_table_html += "</tbody></table><br>"

    return f"""
    <html><head><style>
    body{{font-family:Calibri,sans-serif;font-size:11pt}}
    table{{border-collapse:collapse;width:100%}}
    th,td{{border:1px solid #000;padding:8px;text-align:left;font-size:10pt}}
    th{{background-color:#00B0F0}}
    </style></head><body>
    {ot_details_html}
    {task_table_html}
    {build_signature_html(config)}
    </body></html>
    """

def create_ot_out_html_body(ot_tasks_data, config, shift_date, ot_from_str, ot_to_str, total_hours_str, justification_text):
    date_str = ot_date_display(shift_date)
    day_str = shift_date.strftime('%A')
    to_display = ot_to_str if ot_to_str else '(tbd)'
    total_display = total_hours_str if total_hours_str else '(tbd)'
    justification = (justification_text or '').strip()

    ot_details_html = f"""
    <p>Hi Everyone,</p>
    <p>I will end my OT for today, please see below details for reference:</p>
    <table><thead><tr>
      <th>Date<br>(dd/mm/yyyy)</th><th>Day</th><th>Name</th><th>Stream</th><th>From</th><th>To</th><th>Total Hrs</th><th>Actual Hours</th><th>Comment<br>(Filling / Offset)</th><th>Justification</th>
    </tr></thead><tbody><tr>
      <td>{date_str}</td><td>{day_str}</td><td>{config.get('YOUR_NAME','')}</td><td>{config.get('FIXED_STREAM','')}</td><td>{ot_from_str}</td><td>{to_display}</td><td>{total_display}</td><td></td><td></td><td>{justification}</td>
    </tr></tbody></table><br>
    """

    headers = ['Country','Responsible','Stream','Task List','Offsite Schedule','Frequency','Period','Start Time','End Time','Status','Issue Encountered?','Remarks']
    task_table_html = ''
    if ot_tasks_data:
        task_table_html = "<p><b>Task Status Report:</b></p><table><thead><tr>" + "".join([f"<th>{h}</th>" for h in headers]) + "</tr></thead><tbody>"
        for row in ot_tasks_data:
            task_table_html += "<tr>" + "".join([f"<td>{cell}</td>" for cell in row[:len(headers)]]) + "</tr>"
        task_table_html += "</tbody></table><br>"

    return f"""
    <html><head><style>
    body{{font-family:Calibri,sans-serif;font-size:11pt}}
    table{{border-collapse:collapse;width:100%}}
    th,td{{border:1px solid #000;padding:8px;text-align:left;font-size:10pt}}
    th{{background-color:#00B0F0}}
    </style></head><body>
    {ot_details_html}
    {task_table_html}
    {build_signature_html(config)}
    </body></html>
    """



def build_signature_html(config: dict) -> str:
    """Signature block helper for all bodies."""
    try:
        is_new = str(config.get('OUTLOOK_VERSION', 'Classic') or 'Classic').strip().lower().startswith('new')
    except Exception:
        is_new = False
    try:
        disable = bool(config.get('NEW_OUTLOOK_SIGNATURE_DISABLED', False))
    except Exception:
        disable = False

    # When disabled in New Outlook: leave only 'Regards,' with no extra bottom spacing
    if is_new and disable:
        return "<p style='margin:0'>Regards,</p>"

    name = (config.get('SIGNATURE_NAME') or '').strip()
    if name:
        return f"<p style='margin:0'>Regards,<br><b>{name}</b></p>"
    return "<p style='margin:0'>Regards,</p>"
def generate_email_draft(subject, body, config, screenshot_path=None):
    try:
        outlook = win32.Dispatch('outlook.application')
    except Exception as e:
        messagebox.showerror("Outlook Error", f"Could not connect to 'Classic' Outlook.\n\nDetails: {e}")
        return False

    try:
        mail = outlook.CreateItem(0)
        mail.To = config['RECIPIENTS_TO']
        mail.CC = config['RECIPIENTS_CC']
        mail.Subject = subject
        mail.Recipients.ResolveAll()
        mail.HTMLBody = body

        if screenshot_path and os.path.exists(screenshot_path):
            attachment = mail.Attachments.Add(screenshot_path)
            attachment.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                "biologs_screenshot"
            )
        mail.Display()
        return True
    except Exception as e:
        messagebox.showerror("Outlook Error", f"Could not create email draft.\n\nDetails: {e}")
        return False


# --- UI Application Class ---
class App(BaseWindow):
    def __init__(self):
        super().__init__(themename='flatly') if ('tb' in globals() and tb) else super().__init__()
        os.makedirs(RESOURCES_DIR, exist_ok=True)
        self._cleanup_old_data_files()
        self.load_config()
        self.load_presets()

        self.title("Prod_task_generator")
        self.geometry("800x650")

        # Header bar
        self.header_frame = ttk.Frame(self, padding=(12, 10))
        self.header_frame.pack(fill='x')
        ttk.Label(self.header_frame, text='Prod_task_generator', font=('Segoe UI', 14, 'bold')).pack(side='left')


        # Header navigation buttons (always visible)  # DO NOT REMOVE

        nav = ttk.Frame(self.header_frame)

        nav.pack(side='right')

        ttk.Button(nav, text='🏠 Home', command=lambda: self.show_frame(self.main_frame)).pack(side='left', padx=(0, 6))

        ttk.Button(nav, text='⚙️ Management', command=lambda: self.show_frame(self.management_frame)).pack(side='left')

        ttk.Button(nav, text='❓ Help', command=lambda: self.show_frame(self.help_frame)).pack(side='left', padx=(6, 0))
        self.main_frame = ttk.Frame(self)
        self.help_frame = ttk.Frame(self)
        self.main_frame.pack(fill="both", expand=True)

        self.notebook = ttk.Notebook(self.main_frame)
        self._apply_notebook_tab_highlight()
        self.notebook.pack(pady=10, padx=10, fill="both", expand=True)

        self.sod_frame = ttk.Frame(self.notebook)
        self.eod_frame = ttk.Frame(self.notebook)
        self.ot_in_frame = ttk.Frame(self.notebook)
        self.ot_out_frame = ttk.Frame(self.notebook)

        self.notebook.add(self.sod_frame, text="Start of Day (SOD)")
        self.notebook.add(self.eod_frame, text="End of Day (EOD)")
        self.notebook.add(self.ot_in_frame, text="OT In")
        self.notebook.add(self.ot_out_frame, text="OT Out")

        self.sod_full_data_storage = {}
        self.eod_full_data = {}

        self.current_shift_date = None
        self.eod_screenshot_path = tk.StringVar()
        self.loaded_sod_created_time = "N/A"
        self.loaded_actual_start_shift = None

        self.create_sod_widgets()
        self.create_eod_widgets()
        self.create_ot_in_widgets()
        self.create_ot_out_widgets()
        self.create_settings_widgets()
        self.create_presets_widgets()
        self.create_management_widgets()
        self.create_help_widgets()
        try:
            self.after(1200, self.auto_check_for_updates_on_startup)
        except Exception:
            pass
        # Apply Outlook version UI on startup (New Outlook hides Prepare buttons)
        try:
            self._apply_outlook_version_ui()
        except Exception:
            pass



        # Enable Excel-like column reordering for task tables
        try:
            for _t in (getattr(self, 'sod_tree', None), getattr(self, 'eod_tree', None), getattr(self, 'ot_in_tree', None), getattr(self, 'ot_out_tree', None), getattr(self, 'preset_tree', None)):
                if _t is not None:
                    self._enable_column_reorder(_t)
        except Exception:
            pass
    # -----------------------------

    def _apply_notebook_tab_highlight(self):
        """Ensure selected notebook tab is highlighted (selected tab styling)."""
        try:
            # ttkbootstrap Window may provide self.style; fallback to ttk.Style
            style = getattr(self, 'style', None)
            if style is None:
                style = ttk.Style()
        except Exception:
            try:
                style = ttk.Style()
            except Exception:
                return
        try:
            style.configure('TNotebook.Tab', padding=(12, 6))
            style.map('TNotebook.Tab',
                      background=[('selected', '#1F4E79')],
                      foreground=[('selected', 'white')])
        except Exception:
            pass

    # -----------------------------
    # OUTLOOK VERSION + COPY WORKFLOW
    # -----------------------------
    def _is_new_outlook(self) -> bool:
        """Return True if New Outlook mode is saved in config (not just selected)."""
        try:
            return str(self.config.get('OUTLOOK_VERSION', 'Classic') or 'Classic').strip().lower().startswith('new')
        except Exception:
            return False

    def _apply_outlook_version_ui(self):
        """Apply UI changes based on selected Outlook version."""
        is_new = self._is_new_outlook()

        # Screenshot tools (EOD) - hidden in New Outlook
        try:
            if hasattr(self, 'screenshot_frame'):
                if is_new:
                    try:
                        self.screenshot_frame.pack_forget()
                    except Exception:
                        pass
                    try:
                        if hasattr(self, 'screenshot_label'):
                            self.screenshot_label.config(text='Screenshot disabled in New Outlook mode.')
                    except Exception:
                        pass
                else:
                    try:
                        if not self.screenshot_frame.winfo_ismapped():
                            self.screenshot_frame.pack(fill='x', padx=10, pady=5)
                    except Exception:
                        pass
        except Exception:
            pass

        # Prepare buttons (hidden only in New Outlook)
        for attr in ('btn_prepare_sod', 'btn_prepare_eod', 'btn_prepare_ot_in', 'btn_prepare_ot_out'):
            try:
                b = getattr(self, attr, None)
                if not b:
                    continue
                if is_new:
                    try:
                        b.pack_forget()
                    except Exception:
                        pass
                else:
                    try:
                        if not b.winfo_ismapped():
                            opts = getattr(self, f'_{attr}_pack', None) or {}
                            # Keep SOD button order stable when switching back to Classic
                            if attr == 'btn_prepare_sod':
                                try:
                                    parent = b.master
                                    kids = [w for w in parent.winfo_children() if w is not b]
                                    before_widget = kids[0] if kids else None
                                    if before_widget is not None:
                                        b.pack(**opts, before=before_widget)
                                    else:
                                        b.pack(**opts)
                                except Exception:
                                    b.pack(**opts)
                            else:
                                b.pack(**opts)
                    except Exception:
                        pass
            except Exception:
                pass

        # Copy bars (visible only in New Outlook)
        try:
            self._set_copy_bars_visible(is_new)
        except Exception:
            pass

    def _set_copy_bars_visible(self, visible: bool):
        """Show/hide copy bars across tabs."""
        # SOD uses grid; others use pack
        try:
            if hasattr(self, 'sod_copy_bar'):
                if visible:
                    self.sod_copy_bar.grid()
                else:
                    self.sod_copy_bar.grid_remove()
        except Exception:
            pass

        for name in ('eod_copy_bar', 'ot_in_copy_bar', 'ot_out_copy_bar'):
            try:
                fr = getattr(self, name, None)
                if not fr:
                    continue
                if visible:
                    if not fr.winfo_ismapped():
                        opts = getattr(self, f'_{name}_pack', None) or {'side': 'bottom', 'fill': 'x', 'padx': 10, 'pady': 5}
                        fr.pack(**opts)
                else:
                    fr.pack_forget()
            except Exception:
                pass

    def _copy_to_clipboard_text(self, value: str, label: str = 'Text'):
        try:
            self.clipboard_clear()
            self.clipboard_append(value if value is not None else '')
            self.update()
            messagebox.showinfo('Copied', f'{label} copied to clipboard.')
        except Exception as e:
            messagebox.showerror('Clipboard Error', f'Could not copy {label}. Details: {e}')

    def _copy_to_clipboard_html(self, html: str, label: str = 'Body'):
        """Copy HTML to clipboard (Windows CF_HTML) with plain-text fallback."""
        html = html or ''
        try:
            plain = re.sub(r'<[^>]+>', '', html)
        except Exception:
            plain = html

        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            # Always provide plain text
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, plain)

            # Provide CF_HTML
            cf_html = win32clipboard.RegisterClipboardFormat('HTML Format')
            html_full = '<html><body><!--StartFragment-->' + html + '<!--EndFragment--></body></html>'
            html_bytes = html_full.encode('utf-8', errors='replace')
            crlf = chr(13) + chr(10)
            header_template = (
                'Version:0.9' + crlf +
                'StartHTML:{:08d}' + crlf +
                'EndHTML:{:08d}' + crlf +
                'StartFragment:{:08d}' + crlf +
                'EndFragment:{:08d}' + crlf
            )
            header0 = header_template.format(0, 0, 0, 0).encode('ascii', errors='replace')
            start_html = len(header0)
            start_frag = start_html + html_bytes.find(b'<!--StartFragment-->') + len(b'<!--StartFragment-->')
            end_frag = start_html + html_bytes.find(b'<!--EndFragment-->')
            end_html = start_html + len(html_bytes)
            header = header_template.format(start_html, end_html, start_frag, end_frag).encode('ascii', errors='replace')
            payload = header + html_bytes
            win32clipboard.SetClipboardData(cf_html, payload)
        except Exception:
            # Text already in clipboard
            pass
        finally:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass

        try:
            messagebox.showinfo('Copied', f'{label} copied to clipboard.')
        except Exception:
            pass

    def _run_copy_action(self, fn, label: str):
        """Run a copy action and show a friendly error if prerequisites are missing."""
        try:
            fn()
        except Exception as e:
            try:
                msg = 'Could not copy {}.'.format(label) + '\\n\\nDetails: {}'.format(e)
                messagebox.showerror('Copy Error', msg)
            except Exception:
                pass

    def _populate_copy_bar(self, bar_frame, subject_fn, body_fn):
        """Populate a bar frame with Copy buttons in the order: Body - Subject - To - CC."""
        try:
            ttk.Button(bar_frame, text='Copy Body', command=lambda: self._run_copy_action(body_fn, 'Body')).pack(side='left', padx=(0, 6))
            ttk.Button(bar_frame, text='Copy Subject', command=lambda: self._run_copy_action(subject_fn, 'Subject')).pack(side='left', padx=(0, 12))
            ttk.Button(bar_frame, text='Copy To', command=lambda: self._run_copy_action(self.copy_recipients_to, 'To')).pack(side='left', padx=(0, 6))
            ttk.Button(bar_frame, text='Copy CC', command=lambda: self._run_copy_action(self.copy_recipients_cc, 'CC')).pack(side='left')
            ttk.Label(bar_frame, text='(For New Outlook: paste into a new email)', foreground='#1F4E79').pack(side='right')
        except Exception:
            pass

    def _normalize_email_lines(self, raw: str) -> str:
        """Normalize email inputs (one per line or ';') to '; ' for New Outlook paste."""
        s = (raw or '').replace('\r', '')
        parts = []
        for chunk in s.replace(';', '\n').split('\n'):
            item = (chunk or '').strip()
            if item:
                parts.append(item)
        return '; '.join(parts)



    def _validate_iqvia_emails(self, raw: str):

        """Validate IQVIA company emails: firstname.lastname@iqvia.com (lowercase).

        Returns (cleaned_emails_list, invalid_list)."""

        s = (raw or '').replace('\r', '')

        parts = []

        for chunk in s.replace(';', '\n').split('\n'):

            item = (chunk or '').strip()

            if item:

                parts.append(item)

        cleaned = []

        invalid = []

        rx = re.compile(r'^[a-z]+\.[a-z]+@iqvia\.com$')

        for e in parts:

            el = (e or '').strip().lower()

            if rx.match(el):

                cleaned.append(el)

            else:

                invalid.append(e)

        return cleaned, invalid


    def _ensure_recipient_styles(self):

        """Create ttk styles used to highlight which recipient mode/box is active/selected."""

        if getattr(self, '_recipient_styles_ready', False):

            return

        try:

            st = getattr(self, 'style', None)

            if st is None:

                from tkinter import ttk as _ttk

                st = _ttk.Style()


            # Frame styles

            st.configure('Active.TLabelframe', relief='solid', borderwidth=2)

            st.configure('Active.TLabelframe.Label', foreground='#1F4E79', font=('Segoe UI', 10, 'bold'))


            st.configure('Selected.TLabelframe', relief='solid', borderwidth=3)

            st.configure('Selected.TLabelframe.Label', foreground='#0B2F4D', font=('Segoe UI', 10, 'bold'))

            st.configure('Hover.TLabelframe', relief='solid', borderwidth=2)

            st.configure('Hover.TLabelframe.Label', foreground='#1F4E79', font=('Segoe UI', 10, 'bold'))

            st.configure('InactiveHover.TLabelframe', relief='solid', borderwidth=2)

            st.configure('InactiveHover.TLabelframe.Label', foreground='#808080', font=('Segoe UI', 10, 'bold'))


            st.configure('Inactive.TLabelframe', relief='groove', borderwidth=1)

            st.configure('Inactive.TLabelframe.Label', foreground='#808080', font=('Segoe UI', 10))


            # Entry style (Classic TO) – keep bold and grey background when disabled/inactive

            st.configure('Active.TEntry', font=('Segoe UI', 10, 'bold'))

            st.configure('Inactive.TEntry', font=('Segoe UI', 10, 'bold'))

            st.map('Active.TEntry',

                   fieldbackground=[('disabled', 'white'), ('!disabled', 'white')],

                   foreground=[('disabled', '#1F1F1F'), ('!disabled', '#1F1F1F')])

            st.map('Inactive.TEntry',

                   fieldbackground=[('disabled', '#F2F2F2'), ('!disabled', '#F2F2F2')],

                   foreground=[('disabled', '#808080'), ('!disabled', '#808080')])

            # Borderless Entry styles (Classic To) – remove focus/hover border ring

            st.configure('ActiveNoBorder.TEntry', font=('Segoe UI', 10, 'bold'), relief='flat', borderwidth=0)

            st.configure('InactiveNoBorder.TEntry', font=('Segoe UI', 10, 'bold'), relief='flat', borderwidth=0)

            try:

                st.map('ActiveNoBorder.TEntry',

                       fieldbackground=[('disabled', 'white'), ('!disabled', 'white')],

                       foreground=[('disabled', '#1F1F1F'), ('!disabled', '#1F1F1F')],

                       bordercolor=[('focus', 'white'), ('!focus', 'white')],

                       focuscolor=[('focus', 'white'), ('!focus', 'white')])

                st.map('InactiveNoBorder.TEntry',

                       fieldbackground=[('disabled', '#F2F2F2'), ('!disabled', '#F2F2F2')],

                       foreground=[('disabled', '#808080'), ('!disabled', '#808080')],

                       bordercolor=[('focus', '#F2F2F2'), ('!focus', '#F2F2F2')],

                       focuscolor=[('focus', '#F2F2F2'), ('!focus', '#F2F2F2')])

            except Exception:

                pass


        except Exception:

            pass

        self._recipient_styles_ready = True

    def _refresh_recipient_inputs_state(self):

        """Enable/disable Classic vs New Outlook recipient boxes, grey inactive ones,

        and highlight the box the user selected (focus/click) within the active mode.

        """

        try:

            self._ensure_recipient_styles()

        except Exception:

            pass


        try:

            is_new = bool(self._is_new_outlook())

        except Exception:

            is_new = False

        focus_key = getattr(self, '_recipient_focus_key', '') or ''

        hover_key = getattr(self, '_recipient_hover_key', '') or ''

        # Only allow hover/selected highlights within the active Outlook mode
        try:
            if is_new:
                if focus_key in ('to_classic', 'cc_classic'):
                    focus_key = ''
                if hover_key in ('to_classic', 'cc_classic'):
                    hover_key = ''
            else:
                if focus_key in ('to_new', 'cc_new'):
                    focus_key = ''
                if hover_key in ('to_new', 'cc_new'):
                    hover_key = ''
        except Exception:
            pass


        # styles

        active_frame = 'Active.TLabelframe'

        selected_frame = 'Selected.TLabelframe'

        hover_frame = 'Hover.TLabelframe'

        inactive_hover_frame = 'InactiveHover.TLabelframe'

        inactive_frame = 'Inactive.TLabelframe'


        # ---- widget states (greyed out = disabled but still visible) ----

        try:
            if hasattr(self, 'to_entry'):
                dim_bg = '#F2F2F2'
                dim_fg = '#808080'
                normal_bg = 'white'
                normal_fg = '#1F1F1F'
                if is_new:
                    try: self.to_entry.config(state='disabled')
                    except Exception: pass
                    try: self.to_entry.config(bg=dim_bg, fg=dim_fg, insertbackground=dim_fg, cursor='arrow', takefocus=0)
                    except Exception: pass
                    try: self.to_entry.tag_remove('sel', '1.0', 'end')
                    except Exception: pass
                else:
                    try: self.to_entry.config(state='normal')
                    except Exception: pass
                    try: self.to_entry.config(bg=normal_bg, fg=normal_fg, insertbackground=normal_fg, cursor='xterm', takefocus=1)
                    except Exception: pass
        except Exception:
            pass



        try:

            if hasattr(self, 'cc_entry'):

                try: self.cc_entry.config(state=('disabled' if is_new else 'normal'))

                except Exception: pass

        except Exception:

            pass
        except Exception:

            pass
        except Exception:

            pass


        # ---- grey appearance for tk.Text widgets ----

        # NOTE: tk.Text does not support 'disabledbackground', so we explicitly set colors.

        dim_bg = '#F2F2F2'

        dim_fg = '#808080'

        normal_bg = 'white'

        normal_fg = '#1F1F1F'


        def _set_text_widget_appearance(_w, _active: bool):

            try:

                _w.configure(

                    bg=(normal_bg if _active else dim_bg),

                    fg=(normal_fg if _active else dim_fg),

                    insertbackground=(normal_fg if _active else dim_fg),

                    cursor=('xterm' if _active else 'arrow')

                )

            except Exception:

                try:

                    _w.config(bg=(normal_bg if _active else dim_bg), fg=(normal_fg if _active else dim_fg))

                except Exception:

                    pass


            # When inactive, also block selection/focus.

            if not _active:

                try:

                    _w.tag_remove('sel', '1.0', 'end')

                except Exception:

                    pass

                try:

                    _w.configure(takefocus=0)

                except Exception:

                    pass

            else:

                try:

                    _w.configure(takefocus=1)

                except Exception:

                    pass
        try:

            if hasattr(self, 'cc_entry'):

                # Classic CC box is active only in Classic mode.

                _set_text_widget_appearance(self.cc_entry, (not is_new))
                try:
                    if hasattr(self, 'cc_new_entry'):
                        _set_text_widget_appearance(self.cc_new_entry, is_new)
                except Exception:
                    pass
                try:
                    if hasattr(self, 'to_new_entry'):
                        _set_text_widget_appearance(self.to_new_entry, is_new)
                except Exception:
                    pass

        except Exception:

            pass
        except Exception:

            pass



        # ---- frame highlight logic ----

        # Determine which box is "selected" only within active mode.

        to_selected = ('to_new' if is_new else 'to_classic')

        cc_selected = ('cc_new' if is_new else 'cc_classic')

        if focus_key in ('to_classic','to_new'):

            to_selected = focus_key

        if focus_key in ('cc_classic','cc_new'):

            cc_selected = focus_key


        # Apply styles: selected -> Selected, other in same mode -> Active, inactive mode -> Inactive

        try:

            if hasattr(self, 'to_classic_frame'):

                if is_new:

                    self.to_classic_frame.configure(style=(inactive_frame))

                else:

                    self.to_classic_frame.configure(style=(active_frame))  # Classic-To: fixed style (no hover/click highlight)

        except Exception:

            pass

        try:

            if hasattr(self, 'to_new_frame'):

                if not is_new:

                    self.to_new_frame.configure(style=(inactive_frame))

                else:

                    self.to_new_frame.configure(style=(selected_frame if to_selected=='to_new' else (hover_frame if hover_key=='to_new' else active_frame)))

        except Exception:

            pass

        try:

            if hasattr(self, 'cc_classic_frame'):

                if is_new:

                    self.cc_classic_frame.configure(style=(inactive_frame))

                else:

                    self.cc_classic_frame.configure(style=(selected_frame if cc_selected=='cc_classic' else (hover_frame if hover_key=='cc_classic' else active_frame)))

        except Exception:

            pass

        try:

            if hasattr(self, 'cc_new_frame'):

                if not is_new:

                    self.cc_new_frame.configure(style=(inactive_frame))

                else:

                    self.cc_new_frame.configure(style=(selected_frame if cc_selected=='cc_new' else (hover_frame if hover_key=='cc_new' else active_frame)))

        except Exception:

            pass


        # ttkbootstrap: ensure inactive frames are visually greyed-out (matches Classic/New behavior)

        try:

            if ('tb' in globals() and tb):

                # Blue borders for all recipient windows EXCEPT Classic-To (requested)
                for _fr_name in ('to_classic_frame','to_new_frame','cc_classic_frame','cc_new_frame'):
                    try:
                        _fr = getattr(self, _fr_name, None)
                        if _fr is None:
                            continue
                        _boot = ('secondary' if _fr_name == 'to_classic_frame' else 'primary')
                        _fr.configure(bootstyle=_boot)
                        # Clear explicit ttk style so bootstyle border is visible
                        try:
                            _fr.configure(style='')
                        except Exception:
                            pass
                    except Exception:
                        pass

        except Exception:
            pass

        try:

            self._apply_to_tab_state()
            self._apply_cc_tab_state()

        except Exception:

            pass


# Badges

        try:

            if hasattr(self, 'to_classic_badge'):

                self.to_classic_badge.config(text=('ACTIVE' if (not is_new) else ''), foreground=('#1F4E79' if (not is_new) else '#808080'))

            if hasattr(self, 'to_new_badge'):

                self.to_new_badge.config(text=('ACTIVE' if is_new else ''), foreground=('#1F4E79' if is_new else '#808080'))

            if hasattr(self, 'cc_classic_badge'):

                self.cc_classic_badge.config(text=('ACTIVE' if (not is_new) else ''), foreground=('#1F4E79' if (not is_new) else '#808080'))

            if hasattr(self, 'cc_new_badge'):

                self.cc_new_badge.config(text=('ACTIVE' if is_new else ''), foreground=('#1F4E79' if is_new else '#808080'))

        except Exception:

            pass

    def _show_to_mode_hint(self, msg: str, duration_ms: int = 2500):

        """Show a small hint under the Settings > To tab."""

        try:

            if not hasattr(self, 'to_mode_hint_var'):

                return

            self.to_mode_hint_var.set(str(msg or ''))

            if getattr(self, '_to_hint_after_id', None):

                try:

                    self.after_cancel(self._to_hint_after_id)

                except Exception:

                    pass

            def _clear():

                try:

                    self.to_mode_hint_var.set('')

                except Exception:

                    pass

            self._to_hint_after_id = self.after(int(duration_ms), _clear)

        except Exception:

            pass



    def _to_new_block_event(self, event=None):

        """Block all interaction with New Outlook TO box when Classic mode is active."""

        try:

            is_new = bool(self._is_new_outlook())

        except Exception:

            is_new = False


        if is_new:

            return None  # allow normal behavior


        # Classic mode active -> completely dead

        try:

            self._show_to_mode_hint('Hint: Switch Outlook Version to NEW to edit the New Outlook To field.')

        except Exception:

            pass

        try:

            if hasattr(self, 'to_entry'):

                self.to_entry.focus_set()

        except Exception:

            pass

        return 'break'



    def _to_new_hover_enter(self, event=None):



        """Border highlight on hover (only when New Outlook mode is active)."""



        try:



            if not bool(self._is_new_outlook()):



                return



        except Exception:



            return




        try:



            if hasattr(self, 'to_new_frame'):



                # store original bootstyle once



                if not hasattr(self, '_to_new_frame_bootstyle_orig'):



                    try:



                        self._to_new_frame_bootstyle_orig = str(self.to_new_frame.cget('bootstyle') or '')



                    except Exception:



                        self._to_new_frame_bootstyle_orig = ''



                # ttkbootstrap path



                try:



                    self.to_new_frame.configure(bootstyle='secondary')



                    return



                except Exception:



                    pass



                # ttk fallback



                try:



                    self.to_new_frame.configure(style='Selected.TLabelframe')



                except Exception:



                    pass



        except Exception:



            pass



    def _wire_to_tab_behavior(self):



        """Bind To-tab behavior.




        Classic mode: New Outlook To is completely dead (no selection, no focus, no scroll).



        New mode: New Outlook frame shows border highlight on hover.



        """



        if getattr(self, '_to_tab_wired', False):



            return



        self._to_tab_wired = True




        # targets for hover



        targets = []



        for attr in ('to_new_frame', 'to_email_wrap', 'to_email_entry', 'to_email_scroll'):



            try:



                w = getattr(self, attr, None)



                if w is not None:



                    targets.append(w)



            except Exception:



                pass




        for w in targets:



            try:



                w.bind('<Enter>', self._to_new_hover_enter, add='+')



            except Exception:



                pass



            try:



                w.bind('<Leave>', self._to_new_hover_leave, add='+')



            except Exception:



                pass




        # block events when Classic mode is active



        block_events = (



            '<Button-1>', '<B1-Motion>', '<ButtonRelease-1>', '<Double-Button-1>',



            '<Key>', '<FocusIn>', '<MouseWheel>', '<Button-4>', '<Button-5>'



        )




        try:



            if hasattr(self, 'to_new_frame'):



                self.to_new_frame.bind('<Button-1>', self._to_new_block_event, add='+')



        except Exception:



            pass
        try:



            if hasattr(self, 'to_email_wrap'):



                self.to_email_wrap.bind('<Button-1>', self._to_new_block_event, add='+')



        except Exception:



            pass




        try:



            if hasattr(self, 'to_email_scroll'):



                for ev in ('<Button-1>', '<B1-Motion>', '<ButtonRelease-1>', '<MouseWheel>', '<Button-4>', '<Button-5>'):



                    self.to_email_scroll.bind(ev, self._to_new_block_event, add='+')



        except Exception:



            pass



    def _to_new_hover_leave(self, event=None):



        """Restore style after hover on New Outlook To window."""



        try:



            if hasattr(self, 'to_new_frame'):



                try:



                    self.to_new_frame.configure(bootstyle=getattr(self, '_to_new_frame_bootstyle_orig', ''))



                except Exception:



                    pass



                try:



                    orig_style = getattr(self, '_to_new_frame_style_orig', '')



                    if orig_style:



                        self.to_new_frame.configure(style=orig_style)



                except Exception:



                    pass



        except Exception:



            pass



        # ensure active/inactive styles stay correct



        try:



            self._refresh_recipient_inputs_state()



        except Exception:



            pass





    def _is_recipient_key_active(self, key: str, is_new: bool = None) -> bool:
        """Return True if the recipient box key belongs to the currently selected Outlook mode."""
        try:
            if is_new is None:
                is_new = bool(self._is_new_outlook())
        except Exception:
            is_new = False
        k = (key or '').strip()
        if not k:
            return False
        if is_new:
            return k in ('to_new', 'cc_new')
        return k in ('to_classic', 'cc_classic')

    def _set_recipient_focus(self, key: str):
        """Highlight a recipient box as *selected* (clicked/focused) within the active Outlook mode."""
        try:
            if not self._is_recipient_key_active(key):
                return
            self._recipient_focus_key = str(key or '')
        except Exception:
            return
        try:
            self._refresh_recipient_inputs_state()
        except Exception:
            pass

    def _set_recipient_hover(self, key: str):
        """Highlight a recipient box on hover (mode-aware).

        Hover borders are shown only for boxes that belong to the selected Outlook version.
        """
        try:
            if not self._is_recipient_key_active(key):
                return
            self._recipient_hover_key = str(key or '')
        except Exception:
            self._recipient_hover_key = ''
        try:
            self._refresh_recipient_inputs_state()
        except Exception:
            pass

    def _clear_recipient_hover(self, key: str):
        """Clear hover highlight (only if it matches the current hover key)."""
        try:
            if getattr(self, '_recipient_hover_key', '') == str(key or ''):
                self._recipient_hover_key = ''
        except Exception:
            self._recipient_hover_key = ''
        try:
            self._refresh_recipient_inputs_state()
        except Exception:
            pass

    def _wire_recipient_hover_behavior(self):
        """Bind hover + click handlers for recipient boxes (To/CC, Classic/New).

        - Hover border highlight: only for the active Outlook mode (Classic or New).
        - Click selection border highlight: only for the active Outlook mode.
        """
        if getattr(self, '_recipient_hover_wired', False):
            return
        self._recipient_hover_wired = True

        def _bind_group(key, widgets):
            for w in widgets:
                if w is None:
                    continue
                try:
                    w.bind('<Enter>', lambda _e, k=key: self._set_recipient_hover(k), add='+')
                except Exception:
                    pass
                try:
                    w.bind('<Leave>', lambda _e, k=key: self._clear_recipient_hover(k), add='+')
                except Exception:
                    pass
                # Some widgets (scrollbars/wrap frames) won't trigger FocusIn on the Text/Entry,
                # so we also mark selection on mouse click.
                try:
                    w.bind('<Button-1>', lambda _e, k=key: self._set_recipient_focus(k), add='+')
                except Exception:
                    pass

        _bind_group('to_classic', [
            getattr(self, 'to_classic_frame', None),
            getattr(self, 'to_entry', None),
        ])
        _bind_group('cc_classic', [
            getattr(self, 'cc_classic_frame', None),
            getattr(self, 'cc_classic_wrap', None),
            getattr(self, 'cc_entry', None),
            getattr(self, 'cc_classic_scroll', None),
        ])
    def _apply_to_tab_state(self):

        """Apply enable/disable + border color for Settings > To tab Classic window."""

        try:

            is_new = bool(self._is_new_outlook())

        except Exception:

            is_new = False

    

        try:

            if hasattr(self, 'to_classic_border'):

                self.to_classic_border.configure(bg=('#C0C0C0' if is_new else '#1F4E79'))

        except Exception:

            pass

    

        # New Outlook To border/frame (blue when New is active, grey when Classic is active)
        try:
            if hasattr(self, 'to_new_border'):
                self.to_new_border.configure(bg=('#1F4E79' if is_new else '#C0C0C0'))
        except Exception:
            pass
        try:
            if hasattr(self, 'to_new_frame'):
                self.to_new_frame.configure(style=('Active.TLabelframe' if is_new else 'Inactive.TLabelframe'))
        except Exception:
            pass
        try:
            if hasattr(self, 'to_new_save_btn'):
                self.to_new_save_btn.config(state=('normal' if is_new else 'disabled'))
        except Exception:
            pass

        try:

            if hasattr(self, 'to_classic_frame'):

                self.to_classic_frame.configure(style=('Inactive.TLabelframe' if is_new else 'Active.TLabelframe'))

        except Exception:

            pass


    

        # Text widget state + colors (mirror CC style)


    

        try:


    

            if hasattr(self, 'to_entry'):


    

                dim_bg = '#F2F2F2'


    

                dim_fg = '#808080'


    

                normal_bg = 'white'


    

                normal_fg = '#1F1F1F'


    

                if is_new:


    

                    try: self.to_entry.config(state='disabled')


    

                    except Exception: pass


    

                    try: self.to_entry.config(bg=dim_bg, fg=dim_fg, insertbackground=dim_fg, cursor='arrow', takefocus=0)


    

                    except Exception: pass


    

                    try: self.to_entry.tag_remove('sel', '1.0', 'end')


    

                    except Exception: pass


    

                else:


    

                    try: self.to_entry.config(state='normal')


    

                    except Exception: pass


    

                    try: self.to_entry.config(bg=normal_bg, fg=normal_fg, insertbackground=normal_fg, cursor='xterm', takefocus=1)


    

                    except Exception: pass


    

        except Exception:


    

            pass
        # New Outlook To Text widget state + colors
        try:
            if hasattr(self, 'to_new_entry'):
                dim_bg = '#F2F2F2'
                dim_fg = '#808080'
                normal_bg = 'white'
                normal_fg = '#1F1F1F'
                if is_new:
                    try: self.to_new_entry.config(state='normal')
                    except Exception: pass
                    try: self.to_new_entry.config(bg=normal_bg, fg=normal_fg, insertbackground=normal_fg, cursor='xterm', takefocus=1)
                    except Exception: pass
                else:
                    try: self.to_new_entry.config(state='disabled')
                    except Exception: pass
                    try: self.to_new_entry.config(bg=dim_bg, fg=dim_fg, insertbackground=dim_fg, cursor='arrow', takefocus=0)
                    except Exception: pass
                    try: self.to_new_entry.tag_remove('sel', '1.0', 'end')
                    except Exception: pass
        except Exception:
            pass

    

    def _apply_cc_tab_state(self):

    

        """Apply enable/disable + border color for Settings > CC tab Classic window."""

    

        try:

    

            is_new = bool(self._is_new_outlook())

    

        except Exception:

    

            is_new = False

    

    

    

        # Border color: blue when Classic is active, grey when New is active

    

        try:

    

            if hasattr(self, 'cc_classic_border'):

    

                self.cc_classic_border.configure(bg=('#C0C0C0' if is_new else '#1F4E79'))

    

        except Exception:

    

            pass

    

    

    


    

        # New Outlook CC border/frame (blue when New is active, grey when Classic is active)

    

        try:

    

            if hasattr(self, 'cc_new_border'):

    

                self.cc_new_border.configure(bg=('#1F4E79' if is_new else '#C0C0C0'))

    

        except Exception:

    

            pass

    

        try:

    

            if hasattr(self, 'cc_new_frame'):

    

                self.cc_new_frame.configure(style=('Active.TLabelframe' if is_new else 'Inactive.TLabelframe'))

    

        except Exception:

    

            pass

    

        try:

    

            if hasattr(self, 'cc_new_save_btn'):

    

                self.cc_new_save_btn.config(state=('normal' if is_new else 'disabled'))

    

        except Exception:

    

            pass


    

        # New Outlook CC Text widget state + colors

    

        try:

    

            if hasattr(self, 'cc_new_entry'):

    

                dim_bg = '#F2F2F2'

    

                dim_fg = '#808080'

    

                normal_bg = 'white'

    

                normal_fg = '#1F1F1F'

    

                if is_new:

    

                    try: self.cc_new_entry.config(state='normal')

    

                    except Exception: pass

    

                    try: self.cc_new_entry.config(bg=normal_bg, fg=normal_fg, insertbackground=normal_fg, cursor='xterm', takefocus=1)

    

                    except Exception: pass

    

                else:

    

                    try: self.cc_new_entry.config(state='disabled')

    

                    except Exception: pass

    

                    try: self.cc_new_entry.config(bg=dim_bg, fg=dim_fg, insertbackground=dim_fg, cursor='arrow', takefocus=0)

    

                    except Exception: pass

    

                    try: self.cc_new_entry.tag_remove('sel', '1.0', 'end')

    

                    except Exception: pass

    

        except Exception:

    

            pass

        # Frame style

    

        try:

    

            if hasattr(self, 'cc_classic_frame'):

    

                self.cc_classic_frame.configure(style=('Inactive.TLabelframe' if is_new else 'Active.TLabelframe'))

    

        except Exception:

    

            pass

    

    

    

        # Text widget state + colors

    

        try:

    

            if hasattr(self, 'cc_entry'):

    

                dim_bg = '#F2F2F2'

    

                dim_fg = '#808080'

    

                normal_bg = 'white'

    

                normal_fg = '#1F1F1F'

    

                if is_new:

    

                    try:

    

                        self.cc_entry.config(state='disabled')

    

                    except Exception:

    

                        pass

    

                    try:

    

                        self.cc_entry.config(bg=dim_bg, fg=dim_fg, insertbackground=dim_fg, cursor='arrow', takefocus=0)

    

                    except Exception:

    

                        pass

    

                    try:

    

                        self.cc_entry.tag_remove('sel', '1.0', 'end')

    

                    except Exception:

    

                        pass

    

                else:

    

                    try:

    

                        self.cc_entry.config(state='normal')

    

                    except Exception:

    

                        pass

    

                    try:

    

                        self.cc_entry.config(bg=normal_bg, fg=normal_fg, insertbackground=normal_fg, cursor='xterm', takefocus=1)

    

                    except Exception:

    

                        pass

    

        except Exception:

    

            pass


    

    def copy_recipients_to(self):
        try:
            if self._is_new_outlook():
                value = self.config.get('RECIPIENTS_TO_NEW', '')
            else:
                value = self.config.get('RECIPIENTS_TO', '')
        except Exception:
            value = self.config.get('RECIPIENTS_TO', '')
        self._copy_to_clipboard_text(value, label='To')
    def copy_recipients_cc(self):
        try:
            if self._is_new_outlook():
                value = self.config.get('RECIPIENTS_CC_NEW', '')
            else:
                value = self.config.get('RECIPIENTS_CC', '')
        except Exception:
            value = self.config.get('RECIPIENTS_CC', '')
        self._copy_to_clipboard_text(value, label='CC')
    def copy_sod_subject(self):
        subject, _body = self._build_sod_content(save=False)
        self._copy_to_clipboard_text(subject, label='Subject')

    def copy_sod_body(self):
        subject, body = self._build_sod_content(save=True)
        self._copy_to_clipboard_html(body, label='Body')

    def _build_sod_content(self, save: bool = False):
        if not self.sod_tree.get_children():
            raise RuntimeError('Task list is empty.')

        tasks_data = []
        offsite_schedule = datetime.date.today().strftime('%d/%m/%Y')
        for item_id in self.sod_tree.get_children():
            tasklist, stream, frequency, period, start_time, end_time, issue, remarks = self.sod_full_data_storage[item_id]
            tasks_data.append((
                self.config.get('FIXED_COUNTRY', ''),
                self.config.get('YOUR_NAME', ''),
                stream,
                tasklist,
                offsite_schedule,
                frequency,
                period,
                start_time,
                end_time,
                issue,
                remarks
            ))

        sod_created_time = now_display_time()
        ah = getattr(self, 'sod_shift_hour_var', tk.StringVar()).get()
        am = getattr(self, 'sod_shift_minute_var', tk.StringVar()).get()
        ap = getattr(self, 'sod_shift_ampm_var', tk.StringVar()).get()
        actual_start_shift = (f"{ah}:{am}{ap}" if (ah and am and ap) else (self.config.get('START_TIME') or ''))

        subject = f"WFH SOD Notification | {self.config['YOUR_NAME']} | {datetime.date.today().strftime('%d/%m/%Y')}"
        body = create_sod_html_body(tasks_data, self.config, sod_created_time, actual_start_shift=actual_start_shift)

        if save:
            sod_file = os.path.join(RESOURCES_DIR, f"sod_tasks_{datetime.date.today().strftime('%Y-%m-%d')}.json")
            payload = {
                'meta': {
                    'sod_created_time': sod_created_time,
                    'sod_created_iso': datetime.datetime.now().isoformat(timespec='seconds'),
                    'actual_start_shift': actual_start_shift
                },
                'tasks': tasks_data
            }
            try:
                with open(sod_file, 'w') as f:
                    json.dump(payload, f, indent=4)
            except Exception:
                pass

        return subject, body

    def copy_eod_subject(self):
        subject, _body = self._build_eod_content(save=False)
        self._copy_to_clipboard_text(subject, label='Subject')

    def copy_eod_body(self):
        subject, body = self._build_eod_content(save=True)
        self._copy_to_clipboard_html(body, label='Body')

    def _build_eod_content(self, save: bool = False):
        hour, minute, ampm = self.eod_hour_var.get(), self.eod_minute_var.get(), self.eod_ampm_var.get()
        actual_end_time = f"{hour}:{minute}{ampm}" if (hour and minute and ampm) else (self.config.get('END_TIME') or None)

        eod_data = list(self.eod_full_data.values())
        if not self.current_shift_date:
            self.current_shift_date = infer_shift_date_from_config(datetime.datetime.now(), self.config)

        has_valid_tasks = bool(eod_data) and all(len(row) >= 9 for row in eod_data)
        tasks_to_render = eod_data if has_valid_tasks else None

        eod_created_time = now_display_time()
        sod_created_time = self.loaded_sod_created_time or 'N/A'

        # Ensure actual start shift is available
        if not getattr(self, 'loaded_actual_start_shift', None):
            meta_fallback, date_used = self._get_latest_sod_meta()
            if meta_fallback:
                self.loaded_actual_start_shift = meta_fallback.get('actual_start_shift')
                if (self.loaded_sod_created_time in (None, '', 'N/A')) and meta_fallback.get('sod_created_time'):
                    self.loaded_sod_created_time = meta_fallback.get('sod_created_time')
                if (not self.current_shift_date) and date_used:
                    self.current_shift_date = date_used

        actual_start_shift = self.loaded_actual_start_shift or (self.config.get('START_TIME') or None)

        body = create_eod_html_body(
            tasks_data=tasks_to_render,
            config=self.config,
            actual_end_time=actual_end_time,
            shift_date=self.current_shift_date,
            include_screenshot=False,
            sod_created_time_str=sod_created_time,
            eod_created_time_str=eod_created_time,
            actual_start_shift=actual_start_shift
        )

        subject_date = self.current_shift_date.strftime('%d/%m/%Y')
        subject = f"WFH EOD Notification | {self.config.get('YOUR_NAME','')} | {self.current_shift_date.strftime('%d/%m/%Y')}"

        if save and has_valid_tasks:
            try:
                eod_file_path = os.path.join(RESOURCES_DIR, f"eod_report_{self.current_shift_date.strftime('%Y-%m-%d')}.json")
                with open(eod_file_path, 'w') as f:
                    json.dump(eod_data, f, indent=4)
            except Exception:
                pass

        return subject, body

    def copy_ot_in_subject(self):
        subject, _body = self._build_ot_in_content(save=False)
        self._copy_to_clipboard_text(subject, label='Subject')

    def copy_ot_in_body(self):
        subject, body = self._build_ot_in_content(save=True)
        self._copy_to_clipboard_html(body, label='Body')

    def _build_ot_in_content(self, save: bool = False):
        now_dt = datetime.datetime.now()
        shift_date = getattr(self, 'ot_selected_date', datetime.date.today())

        ot_from = self._ot_time_str(self.ot_from_h, self.ot_from_m, self.ot_from_a)
        if not ot_from:
            raise RuntimeError('OT From time is required.')

        ot_to = self._ot_time_str(self.ot_to_h, self.ot_to_m, self.ot_to_a)
        total = calculate_total_hours(ot_from, ot_to) if ot_to else ''
        offsite = shift_date.strftime('%d/%m/%Y')

        tasks = []
        for iid in self.ot_in_tree.get_children():
            data = self.ot_in_full_data.get(iid, {})
            status_disp = data.get('status', '')
            status_single = ot_status_single(_ot_status_state_from_any(status_disp))
            tasks.append((
                self.config.get('FIXED_COUNTRY', ''),
                self.config.get('YOUR_NAME', ''),
                '',
                data.get('task', ''),
                offsite,
                '', '', '', '',
                status_single,
                data.get('issue', ''),
                data.get('remarks', '')
            ))

        justification = self.ot_justification_var.get()
        subject = f"OT Notification | {self.config.get('YOUR_NAME','')} | {''} | ({shift_date.strftime('%d.%m.%Y')})"
        body = create_ot_in_html_body(tasks, self.config, shift_date, ot_from, ot_to, total, justification)

        if save:
            payload = {
                'meta': {
                    'shift_date': shift_date.isoformat(),
                    'ot_from': ot_from,
                    'ot_to': ot_to,
                    'justification': justification,
                    'created_time': now_display_time(),
                    'created_iso': now_dt.isoformat(timespec='seconds')
                },
                'tasks': tasks
            }
            try:
                with open(f"{OT_IN_FILE_PREFIX}{shift_date.strftime('%Y-%m-%d')}.json", 'w') as f:
                    json.dump(payload, f, indent=4)
            except Exception:
                pass

        return subject, body

    def copy_ot_out_subject(self):
        subject, _body = self._build_ot_out_content(save=False)
        self._copy_to_clipboard_text(subject, label='Subject')

    def copy_ot_out_body(self):
        subject, body = self._build_ot_out_content(save=True)
        self._copy_to_clipboard_html(body, label='Body')

    def _build_ot_out_content(self, save: bool = False):
        if not self.loaded_ot_in_tasks:
            raise RuntimeError('No OT In tasks loaded. Click Load OT In Tasks first.')

        ot_to = self._ot_time_str(self.ot_out_h, self.ot_out_m, self.ot_out_a)
        if not ot_to:
            raise RuntimeError('Actual End Time is required for OT Out.')

        ot_from = self.loaded_ot_in_meta.get('ot_from') or ''
        if not ot_from:
            raise RuntimeError('OT From time not found from OT In file.')

        shift_date = self.loaded_ot_shift_date or infer_shift_date_from_config(datetime.datetime.now(), self.config)
        total = calculate_total_hours(ot_from, ot_to)
        justification = self.loaded_ot_in_meta.get('justification', '')

        updated_tasks = []
        for iid in self.ot_out_tree.get_children():
            base = self.ot_out_full_data.get(iid, {}).get('task_row', [])
            status_disp = self.ot_out_tree.item(iid, 'values')[1]
            status_single = ot_status_single(_ot_status_state_from_any(status_disp))
            if isinstance(base, list) and len(base) >= 10:
                base[9] = status_single
            updated_tasks.append(tuple(base))

        subject = f"OT Notification | {self.config.get('YOUR_NAME','')} | {self.config.get('FIXED_STREAM','')} | ({shift_date.strftime('%d.%m.%Y')})"
        body = create_ot_out_html_body(updated_tasks, self.config, shift_date, ot_from, ot_to, total, justification)

        if save:
            payload = {
                'meta': {
                    'shift_date': shift_date.isoformat(),
                    'ot_from': ot_from,
                    'ot_to': ot_to,
                    'justification': justification,
                    'created_time': now_display_time(),
                    'created_iso': datetime.datetime.now().isoformat(timespec='seconds')
                },
                'tasks': updated_tasks
            }
            try:
                with open(f"{OT_OUT_FILE_PREFIX}{shift_date.strftime('%Y-%m-%d')}.json", 'w') as f:
                    json.dump(payload, f, indent=4)
            except Exception:
                pass

        return subject, body


    # OT DATE PICKER (CALENDAR)
    # -----------------------------
    def _open_ot_date_picker(self):
        """Open a simple calendar picker for OT Date (no external dependencies)."""
        try:
            initial = getattr(self, 'ot_selected_date', None)
            if not isinstance(initial, datetime.date):
                initial = datetime.date.today()
        except Exception:
            initial = datetime.date.today()

        top = tk.Toplevel(self)
        top.title('Select OT Date')
        top.transient(self)
        top.grab_set()

        month_var = tk.IntVar(value=initial.month)
        year_var = tk.IntVar(value=initial.year)

        header = ttk.Frame(top)
        header.pack(fill='x', padx=10, pady=8)

        title_lbl = ttk.Label(header, text='', anchor='center')
        title_lbl.pack(side='left', expand=True)

        def _change_month(delta):
            y = year_var.get()
            m = month_var.get() + delta
            if m < 1:
                m = 12
                y -= 1
            elif m > 12:
                m = 1
                y += 1
            month_var.set(m)
            year_var.set(y)
            _refresh()

        ttk.Button(header, text='◀', width=3, command=lambda: _change_month(-1)).pack(side='left')
        ttk.Button(header, text='▶', width=3, command=lambda: _change_month(1)).pack(side='right')

        grid = ttk.Frame(top)
        grid.pack(padx=10, pady=(0,10))

        def _select(d: datetime.date):
            self.ot_selected_date = d
            try:
                if hasattr(self, 'ot_date_var'):
                    self.ot_date_var.set(d.strftime('%d/%m/%Y'))
            except Exception:
                pass
            top.destroy()

        def _refresh():
            for w in grid.winfo_children():
                w.destroy()
            y = year_var.get()
            m = month_var.get()
            title_lbl.config(text=f'{calendar.month_name[m]} {y}')
            # Week headers (Mon-Sun)
            for i, wd in enumerate(['Mon','Tue','Wed','Thu','Fri','Sat','Sun']):
                ttk.Label(grid, text=wd, width=4, anchor='center').grid(row=0, column=i, padx=1, pady=1)
            weeks = calendar.monthcalendar(y, m)
            for r, week in enumerate(weeks, start=1):
                for c, day in enumerate(week):
                    if day == 0:
                        ttk.Label(grid, text='', width=4).grid(row=r, column=c, padx=1, pady=1)
                        continue
                    d = datetime.date(y, m, day)
                    b = ttk.Button(grid, text=str(day), width=4, command=lambda dd=d: _select(dd))
                    b.grid(row=r, column=c, padx=1, pady=1)

        quick = ttk.Frame(top)
        quick.pack(fill='x', padx=10, pady=(0,10))
        ttk.Button(quick, text='Today', command=lambda: _select(datetime.date.today())).pack(side='left')
        ttk.Button(quick, text='Cancel', command=top.destroy).pack(side='right')

        _refresh()

    # -----------------------------
    # TREEVIEW COLUMN REORDER (EXCEL-LIKE)
    # -----------------------------
    def _tree_colname_from_event(self, tree, event):
        """Return the column name under the cursor for a Treeview event (respects display order)."""
        try:
            col_id = tree.identify_column(event.x)  # '#1', '#2', ...
            if not col_id or not col_id.startswith('#'):
                return None
            idx = int(col_id[1:]) - 1
            # displaycolumns can be a tuple of column names
            disp = tree.cget('displaycolumns')
            if disp and disp != ('#all',) and disp != '#all':
                disp_cols = list(disp)
            else:
                disp_cols = list(tree['columns'])
            if 0 <= idx < len(disp_cols):
                return disp_cols[idx]
        except Exception:
            return None
        return None

    def _enable_column_reorder(self, tree):
        """Enable drag-and-drop column reordering by dragging header cells.

        Notes:
        - Only activates when the mouse is on the 'heading' region (not separator),
          so normal column resizing still works.
        - Uses Treeview 'displaycolumns' so underlying data order is unchanged.
        """
        state = {'drag_col': None}

        def on_press(event):
            try:
                if tree.identify_region(event.x, event.y) != 'heading':
                    state['drag_col'] = None
                    return
                state['drag_col'] = self._tree_colname_from_event(tree, event)
            except Exception:
                state['drag_col'] = None

        def on_release(event):
            try:
                if tree.identify_region(event.x, event.y) != 'heading':
                    state['drag_col'] = None
                    return
                src_col = state.get('drag_col')
                tgt_col = self._tree_colname_from_event(tree, event)
                state['drag_col'] = None
                if not src_col or not tgt_col or src_col == tgt_col:
                    return
                disp = tree.cget('displaycolumns')
                if disp and disp != ('#all',) and disp != '#all':
                    disp_cols = list(disp)
                else:
                    disp_cols = list(tree['columns'])
                if src_col not in disp_cols or tgt_col not in disp_cols:
                    return
                # Reorder: move src_col to the index of tgt_col
                disp_cols.remove(src_col)
                tgt_idx = disp_cols.index(tgt_col)
                disp_cols.insert(tgt_idx, src_col)
                tree.configure(displaycolumns=disp_cols)
            except Exception:
                state['drag_col'] = None

        # Bind on heading press/release (add='+': do not break existing bindings)
        tree.bind('<ButtonPress-1>', on_press, add='+')
        tree.bind('<ButtonRelease-1>', on_release, add='+')

    # -----------------------------
    # FILE MAINTENANCE + LOAD/SAVE
    # -----------------------------
    def _cleanup_old_data_files(self):
        try:
            cutoff = datetime.date.today() - datetime.timedelta(days=DAYS_TO_KEEP_DATA)
            for filename in os.listdir(RESOURCES_DIR):
                if (filename.startswith("sod_tasks_")) and filename.endswith(".json"):
                    try:
                        date_str = filename.replace("sod_tasks_", "").replace("eod_report_", "").replace(".json", "")
                        file_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
                        if file_date < cutoff:
                            os.remove(os.path.join(RESOURCES_DIR, filename))
                    except (ValueError, OSError):
                        continue

                if filename == "temp_screenshot.png":
                    try:
                        os.remove(os.path.join(RESOURCES_DIR, filename))
                    except OSError:
                        pass
        except Exception:
            pass

    def load_config(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    self.config = {**DEFAULT_CONFIG, **json.load(f)}
            else:
                self.config = DEFAULT_CONFIG.copy()
        except json.JSONDecodeError:
            self.config = DEFAULT_CONFIG.copy()

        # normalize types
        try:
            self.config['MONTHLY_START_DAY'] = int(self.config.get('MONTHLY_START_DAY', 1))
        except Exception:
            self.config['MONTHLY_START_DAY'] = 1

        # normalize weekly start offset (days from Monday)
        # Backward compatibility: map WEEK_START_DAY (Mon-Sun) -> numeric offset
        if 'WEEKLY_START_OFFSET_DAYS' not in self.config:
            wsd = self.config.get('WEEK_START_DAY', 'Monday')
            if wsd in WEEKDAY_TO_INDEX:
                self.config['WEEKLY_START_OFFSET_DAYS'] = int(WEEKDAY_TO_INDEX[wsd])
            else:
                self.config['WEEKLY_START_OFFSET_DAYS'] = 0
        try:
            self.config['WEEKLY_START_OFFSET_DAYS'] = int(self.config.get('WEEKLY_START_OFFSET_DAYS', 0))
        except Exception:
            self.config['WEEKLY_START_OFFSET_DAYS'] = 0


        # normalize week start day
        wsd = self.config.get('WEEK_START_DAY', 'Monday')
        if wsd not in WEEK_START_DAY_OPTIONS:
            self.config['WEEK_START_DAY'] = 'Monday'
    
        # normalize outlook version (Classic/New)
        if 'OUTLOOK_VERSION' not in self.config:
            self.config['OUTLOOK_VERSION'] = 'Classic'
        self.config['OUTLOOK_VERSION'] = 'New' if str(self.config.get('OUTLOOK_VERSION','Classic')).strip().lower().startswith('new') else 'Classic'
        try:
            self.config['NEW_OUTLOOK_SIGNATURE_DISABLED'] = bool(self.config.get('NEW_OUTLOOK_SIGNATURE_DISABLED', False))
        except Exception:
            self.config['NEW_OUTLOOK_SIGNATURE_DISABLED'] = False
        self.config['UPDATE_MANIFEST_URL'] = HARDCODED_UPDATE_MANIFEST_URL
        try:
            self.config['AUTO_CHECK_UPDATES'] = bool(self.config.get('AUTO_CHECK_UPDATES', False))
        except Exception:
            self.config['AUTO_CHECK_UPDATES'] = False
        self.config['SKIPPED_UPDATE_VERSION'] = str(self.config.get('SKIPPED_UPDATE_VERSION', '') or '').strip()
        self.config['LAST_UPDATE_CHECK'] = str(self.config.get('LAST_UPDATE_CHECK', '') or '').strip()
        self.config['LAST_UPDATE_VERSION_SEEN'] = str(self.config.get('LAST_UPDATE_VERSION_SEEN', '') or '').strip()

    def _fetch_update_manifest(self, manifest_url: str) -> dict:
        manifest_url = str(manifest_url or HARDCODED_UPDATE_MANIFEST_URL).strip()
        if not manifest_url:
            raise ValueError('Hardcoded update URL is blank.')
        req = urllib.request.Request(
            manifest_url,
            headers={
                'User-Agent': f'{APP_NAME}/{APP_VERSION}',
                'Accept': 'application/json, text/plain, */*',
            },
        )
        with urllib.request.urlopen(req, timeout=12) as response:
            payload = json.loads(response.read().decode('utf-8'))
        return normalize_update_payload(payload)


    def auto_check_for_updates_on_startup(self):
        try:
            if bool(self.config.get('AUTO_CHECK_UPDATES', False)):
                self.check_for_updates(show_no_update=False, startup=True)
        except Exception:
            pass


    def open_update_download(self, download_url: str):
        url = str(download_url or '').strip()
        if not url:
            raise ValueError('No download URL was provided by the update manifest.')
        ok = webbrowser.open(url, new=2, autoraise=True)
        if not ok:
            raise RuntimeError('Could not open the download page in your default browser.')


    def _build_update_message(self, manifest_info: dict) -> str:
        lines = [
            f"Current version: {clean_version_string(APP_VERSION) or APP_VERSION}",
            f"Latest version: {manifest_info.get('latest_version', '') or '(unknown)'}",
        ]
        if manifest_info.get('published_at'):
            lines.append(f"Published: {manifest_info['published_at']}")
        notes = str(manifest_info.get('release_notes', '') or '').strip()
        if notes:
            snippet = notes[:700] + ('...' if len(notes) > 700 else '')
            lines.extend(['', 'Release notes:', snippet])
        lines.extend(['', 'Open the download page now?'])
        return "\n".join(lines)


    def check_for_updates(self, show_no_update: bool = True, startup: bool = False):
        manifest_url = HARDCODED_UPDATE_MANIFEST_URL

        try:
            info = self._fetch_update_manifest(manifest_url)
        except Exception as e:
            if show_no_update or not startup:
                messagebox.showerror('Update Check Failed', f'Could not check for updates.\n\nDetails: {e}')
            return False

        latest_version = clean_version_string(info.get('latest_version', ''))
        self.config['LAST_UPDATE_CHECK'] = datetime.datetime.now().isoformat(timespec='seconds')
        self.config['LAST_UPDATE_VERSION_SEEN'] = latest_version
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4)
        except Exception:
            pass
        self.refresh_update_ui()

        if not latest_version:
            if show_no_update:
                messagebox.showwarning('Updates', 'The manifest was reached, but it did not contain a latest version value.')
            return False

        cmp_result = compare_version_strings(APP_VERSION, latest_version)
        skipped = clean_version_string(self.config.get('SKIPPED_UPDATE_VERSION', ''))
        mandatory = bool(info.get('mandatory', False))

        if cmp_result >= 0:
            self.refresh_update_ui()
            if show_no_update:
                messagebox.showinfo('Updates', f'You are already on the latest version ({clean_version_string(APP_VERSION) or APP_VERSION}).')
            return False

        if startup and skipped and skipped == latest_version and not mandatory:
            return False

        choice = messagebox.askyesnocancel('Update Available', self._build_update_message(info))
        if choice is None:
            if not mandatory:
                self.config['SKIPPED_UPDATE_VERSION'] = latest_version
                try:
                    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                        json.dump(self.config, f, indent=4)
                except Exception:
                    pass
            self.refresh_update_ui()
            return True

        if choice:
            self.config['SKIPPED_UPDATE_VERSION'] = ''
            try:
                with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                    json.dump(self.config, f, indent=4)
            except Exception:
                pass
            self.refresh_update_ui()
            try:
                self.open_update_download(info.get('download_url', ''))
            except Exception as e:
                messagebox.showerror('Open Download Failed', f'An update is available, but the download page could not be opened.\n\nDetails: {e}')
            return True

        if not mandatory:
            self.config['SKIPPED_UPDATE_VERSION'] = latest_version
            try:
                with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                    json.dump(self.config, f, indent=4)
            except Exception:
                pass
        self.refresh_update_ui()
        return True


    def save_update_settings(self):
        self.config['UPDATE_MANIFEST_URL'] = HARDCODED_UPDATE_MANIFEST_URL
        try:
            self.config['AUTO_CHECK_UPDATES'] = bool(self.auto_check_updates_var.get())
        except Exception:
            self.config['AUTO_CHECK_UPDATES'] = bool(self.config.get('AUTO_CHECK_UPDATES', False))
        self.save_config()
        self.refresh_update_ui()


    def clear_skipped_update_version(self):
        self.config['SKIPPED_UPDATE_VERSION'] = ''
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4)
        except Exception:
            pass
        self.refresh_update_ui()
        messagebox.showinfo('Updates', 'Skipped update version has been cleared.')

    def refresh_update_ui(self):
        """Refresh Help > Updates controls from current config."""
        try:
            self.config['UPDATE_MANIFEST_URL'] = HARDCODED_UPDATE_MANIFEST_URL
            if hasattr(self, 'update_manifest_entry'):
                try:
                    self.update_manifest_entry.config(state='normal')
                except Exception:
                    pass
                self.update_manifest_entry.delete(0, tk.END)
                self.update_manifest_entry.insert(0, HARDCODED_UPDATE_MANIFEST_URL)
                try:
                    self.update_manifest_entry.config(state='readonly')
                except Exception:
                    pass
            if hasattr(self, 'auto_check_updates_var'):
                self.auto_check_updates_var.set(bool(self.config.get('AUTO_CHECK_UPDATES', False)))
            if hasattr(self, 'update_status_label'):
                skipped = self.config.get('SKIPPED_UPDATE_VERSION', '') or '(none)'
                last_check = self.config.get('LAST_UPDATE_CHECK', '') or '(never)'
                last_seen = self.config.get('LAST_UPDATE_VERSION_SEEN', '') or '(none)'
                self.update_status_label.config(text=f'API: {HARDCODED_UPDATE_MANIFEST_URL}\nLast check: {last_check} | Last seen version: {last_seen} | Skipped: {skipped}')
        except Exception:
            pass
    def save_config(self):
        with open(CONFIG_FILE, 'w') as f:
            json.dump(self.config, f, indent=4)
        messagebox.showinfo("Success", "Settings have been saved!")

    def load_presets(self):
        weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        self.presets = {
            "Daily": [],
            "Weekdays": {day: [] for day in weekdays},
            "Monthly": {str(i): [] for i in range(1, 32)},
            "TaskDropdown": []
        }
        try:
            if os.path.exists(PRESETS_FILE):
                with open(PRESETS_FILE, 'r') as f:
                    data_from_file = json.load(f)
                if "Monday" in data_from_file and isinstance(data_from_file["Monday"], list):
                    self.presets["Weekdays"] = data_from_file
                    with open(PRESETS_FILE, 'w') as f_out:
                        json.dump(self.presets, f_out, indent=4)
                else:
                    self.presets.update(data_from_file)
        except Exception:
            pass
        # Build/refresh TaskDropdown cache (supports task->frequency mapping)
        self._refresh_taskdropdown_cache()


    def _normalize_taskdropdown_items(self):
        # Normalize TaskDropdown presets to a list of [tasklist, frequency] pairs (backward compatible)
        try:
            raw = self.presets.get('TaskDropdown', [])
        except Exception:
            raw = []
        normalized = []
        if isinstance(raw, list):
            for item in raw:
                if isinstance(item, str):
                    normalized.append([item, ''])
                elif isinstance(item, (list, tuple)):
                    if len(item) == 0:
                        continue
                    task = item[0]
                    freq = item[1] if len(item) > 1 else ''
                    normalized.append([task, freq])
                elif isinstance(item, dict):
                    task = item.get('task') or item.get('Tasklist') or item.get('name')
                    freq = item.get('frequency') or item.get('Frequency') or ''
                    if task:
                        normalized.append([task, freq])
        # de-duplicate by task while preserving first-seen order
        seen = set()
        deduped = []
        for task, freq in normalized:
            if not task:
                continue
            if task in seen:
                if freq:
                    for row in deduped:
                        if row[0] == task and not row[1]:
                            row[1] = freq
                            break
                continue
            seen.add(task)
            deduped.append([task, freq])
        self.presets['TaskDropdown'] = deduped

    def _refresh_taskdropdown_cache(self):
        # Build lookup map and refresh SOD helper combobox values (if present)
        self._normalize_taskdropdown_items()
        try:
            self.taskdropdown_freq_map = {t: f for t, f in self.presets.get('TaskDropdown', [])}
        except Exception:
            self.taskdropdown_freq_map = {}
        if hasattr(self, 'task_helper_combo'):
            try:
                self.task_helper_combo['values'] = self.get_task_dropdown_display_list()
            except Exception:
                pass


    def _refresh_sod_task_frequencies(self):

        """Update frequency/period for existing SOD tasks based on TaskDropdown mapping.

        This enables real-time updates when the user edits Task Dropdown Options frequencies.

        Only updates tasks whose Tasklist exists in the mapping and the mapped frequency is non-empty."""

        try:

            tree = getattr(self, 'sod_tree', None)

            store = getattr(self, 'sod_full_data_storage', None)

            if tree is None or store is None:

                return

        except Exception:

            return

        # Ensure cache exists

        try:

            if not hasattr(self, 'taskdropdown_freq_map'):

                self._refresh_taskdropdown_cache()

        except Exception:

            pass

        updated = 0

        for iid in list(tree.get_children()):

            try:

                row = store.get(iid)

                if not row or len(row) < 8:

                    continue

                tasklist, stream, freq, period, start_t, end_t, issue, remarks = row

                mapped = (self.get_task_dropdown_frequency(tasklist) or '').strip()

                if not mapped:

                    continue

                # Normalize and compare

                _fl, mapped_disp = normalize_frequency_string(mapped)

                mapped_final = (mapped_disp or mapped).strip()

                _fl2, freq_disp = normalize_frequency_string(freq)

                freq_final = (freq_disp or (freq or '')).strip()

                if mapped_final and mapped_final != freq_final:

                    new_period = self.calculate_period(mapped_final)

                    # Update tree display (keeps Action column)


                    try:

                        cur_vals = list(tree.item(iid, 'values'))

                        # Expected columns: Tasklist, Stream, Frequency, Period, Start Time, End Time, Issue Encountered, Remarks, Action

                        if len(cur_vals) >= 9:

                            cur_vals[2] = mapped_final

                            cur_vals[3] = new_period

                            tree.item(iid, values=tuple(cur_vals))

                    except Exception:

                        pass

                    store[iid] = (tasklist, stream, mapped_final, new_period, start_t, end_t, issue, remarks)

                    updated += 1

            except Exception:

                continue

        # No messagebox here (silent refresh)



    def _refresh_sod_task_periods(self):
        """Recompute Period column for existing SOD tasks based on current General settings.

        This updates listed tasks immediately after saving changes in Settings > General."""
        try:
            tree = getattr(self, 'sod_tree', None)
            store = getattr(self, 'sod_full_data_storage', None)
            if tree is None or store is None:
                return
        except Exception:
            return
        for iid in list(tree.get_children()):
            try:
                row = store.get(iid)
                if not row or len(row) < 8:
                    continue
                tasklist, stream, freq, period, start_t, end_t, issue, remarks = row
                new_period = self.calculate_period(freq)
                if (new_period or '') != (period or ''):
                    # Update tree display
                    try:
                        cur_vals = list(tree.item(iid, 'values'))
                        if len(cur_vals) >= 9:
                            cur_vals[3] = new_period
                            tree.item(iid, values=tuple(cur_vals))
                    except Exception:
                        pass
                    store[iid] = (tasklist, stream, freq, new_period, start_t, end_t, issue, remarks)
            except Exception:
                continue




    def _refresh_sod_task_periods_preview(self):



        """Live-update Period column for existing SOD tasks using CURRENT (unsaved) General tab values.




        This does NOT persist settings to disk; it only previews how Periods would look."""



        # Gather UI values (fallback to saved config)



        try:



            daily_logic = self.daily_logic_var.get() if hasattr(self, 'daily_logic_var') else self.config.get('DAILY_LOGIC')



            weekly_logic = self.weekly_logic_var.get() if hasattr(self, 'weekly_logic_var') else self.config.get('WEEKLY_LOGIC')



            monthly_logic = self.monthly_logic_var.get() if hasattr(self, 'monthly_logic_var') else self.config.get('MONTHLY_LOGIC')



        except Exception:



            daily_logic = self.config.get('DAILY_LOGIC')



            weekly_logic = self.config.get('WEEKLY_LOGIC')



            monthly_logic = self.config.get('MONTHLY_LOGIC')



        try:



            weekly_offset = int(self.weekly_start_offset_var.get()) if hasattr(self, 'weekly_start_offset_var') else int(self.config.get('WEEKLY_START_OFFSET_DAYS', 0) or 0)



        except Exception:



            weekly_offset = int(self.config.get('WEEKLY_START_OFFSET_DAYS', 0) or 0)



        try:



            monthly_start_day = int(self.monthly_start_day_var.get()) if hasattr(self, 'monthly_start_day_var') else int(self.config.get('MONTHLY_START_DAY', 1) or 1)



        except Exception:



            monthly_start_day = int(self.config.get('MONTHLY_START_DAY', 1) or 1)



        # Temporarily swap config values for calculation (restore after)



        keys = ['DAILY_LOGIC','WEEKLY_LOGIC','WEEKLY_START_OFFSET_DAYS','MONTHLY_LOGIC','MONTHLY_START_DAY']



        backup = {k: self.config.get(k) for k in keys}



        try:



            self.config['DAILY_LOGIC'] = daily_logic



            self.config['WEEKLY_LOGIC'] = weekly_logic



            self.config['WEEKLY_START_OFFSET_DAYS'] = weekly_offset



            self.config['MONTHLY_LOGIC'] = monthly_logic



            self.config['MONTHLY_START_DAY'] = monthly_start_day



            self._refresh_sod_task_periods()



        except Exception:



            pass



        finally:



            try:



                for k,v in backup.items():



                    self.config[k] = v



            except Exception:



                pass

    def get_task_dropdown_display_list(self):
        # Return list of task names for the SOD 'Saved Tasks' dropdown
        self._normalize_taskdropdown_items()
        return [t for t, _ in self.presets.get('TaskDropdown', [])]

    def get_task_dropdown_frequency(self, task_name: str) -> str:
        # Return frequency mapped to a task in TaskDropdown ('' if none)
        if not hasattr(self, 'taskdropdown_freq_map'):
            self._refresh_taskdropdown_cache()
        return (self.taskdropdown_freq_map or {}).get(task_name, '')

    def save_presets(self):
        try:
            preset_type = self.preset_type_var.get()
            preset_key = self.preset_key_var.get()
            tasks = [self.preset_tree.item(item, 'values') for item in self.preset_tree.get_children()]
            tasks_to_save = [list(row[:-1]) for row in tasks]

            # Automatic frequencies for Daily/Weekday/Monthly preset types
            try:
                if preset_type in ('Daily', 'Weekday', 'Monthly'):
                    rebuilt = []
                    for row in tasks_to_save:
                        if not row or not row[0]:
                            continue
                        task = row[0]
                        stream_in = (row[1] if len(row) > 1 else '')
                        freq_in = (row[2] if len(row) > 2 else '')
                        # Daily preset is always Daily
                        if preset_type == 'Daily':
                            freq_final = 'Daily'
                        else:
                            _fl, _fd = normalize_frequency_string(str(freq_in))
                            freq_final = (_fd or str(freq_in).strip())
                            if not freq_final and preset_type == 'Weekday':
                                freq_final = 'Weekly'
                            if not freq_final and preset_type == 'Monthly':
                                freq_final = 'Monthly'
                        st = row[3] if len(row) > 3 else ''
                        et = row[4] if len(row) > 4 else ''
                        rebuilt.append([task, stream_in, freq_final, '', st, et])
                    tasks_to_save = rebuilt
            except Exception:
                pass

            save_location_text = ""
            if preset_type == "Daily":
                self.presets["Daily"] = tasks_to_save
                save_location_text = "Daily"

            elif preset_type == "Weekday":
                if not preset_key:
                    messagebox.showwarning("Warning", "Please select a weekday.")
                    return
                self.presets["Weekdays"][preset_key] = tasks_to_save
                save_location_text = f"Weekday: {preset_key}"

            elif preset_type == "Monthly":
                if not preset_key:
                    messagebox.showwarning("Warning", "Please select a day of the month.")
                    return
                self.presets["Monthly"][preset_key] = tasks_to_save
                save_location_text = f"Monthly: Day {preset_key}"

            elif preset_type == "Task Dropdown Options":
                # Backward compatible save:
                # - If all frequencies blank -> keep original presets.json format (list of strings)
                # - If any frequency set -> save as list of [task, frequency] pairs
                task_pairs = []
                any_freq = False
                for row in tasks_to_save:
                    if not row:
                        continue
                    task = row[0]
                    freq = row[1] if len(row) > 1 else ''
                    task_pairs.append([task, freq])
                    if str(freq).strip() != "":
                        any_freq = True

                if any_freq:
                    self.presets["TaskDropdown"] = task_pairs
                else:
                    self.presets["TaskDropdown"] = [tp[0] for tp in task_pairs if tp and tp[0]]

                save_location_text = "Task Dropdown Options"

            with open(PRESETS_FILE, 'w') as f:
                json.dump(self.presets, f, indent=4)

            if preset_type == "Task Dropdown Options":
                try:
                    self._refresh_taskdropdown_cache()
                except Exception:
                    pass
            try:
                self._refresh_sod_task_frequencies()
            except Exception:
                pass

            messagebox.showinfo("Success", f"Presets for {save_location_text} have been saved!")

        except Exception as e:
            messagebox.showerror("Error", f"Could not save presets.\n\nDetails: {e}")

    # -----------------------------
    # FRAME NAVIGATION
    # -----------------------------
    def show_frame(self, frame_to_show):
        self.main_frame.pack_forget()
        self.settings_frame.pack_forget()
        self.presets_frame.pack_forget()
        self.management_frame.pack_forget()
        self.help_frame.pack_forget()
        frame_to_show.pack(fill="both", expand=True)
        if frame_to_show == self.main_frame:
            self.load_config()

            try:
                self._refresh_sod_task_frequencies()
            except Exception:
                pass
            try:
                self._refresh_sod_task_periods()
            except Exception:
                pass
    # -----------------------------
    # SETTINGS UI
    # -----------------------------
    def show_settings_frame(self):
        self.show_frame(self.settings_frame)

        self.name_entry.delete(0, tk.END)
        self.name_entry.insert(0, self.config.get('YOUR_NAME', ''))

        self.signature_entry.delete(0, tk.END)
        self.signature_entry.insert(0, self.config.get('SIGNATURE_NAME', ''))
        try:
            if hasattr(self, 'disable_signature_new_var'):
                self.disable_signature_new_var.set(bool(self.config.get('NEW_OUTLOOK_SIGNATURE_DISABLED', False)))
                state = 'normal' if str(self.config.get('OUTLOOK_VERSION','Classic')).lower().startswith('new') else 'disabled'
                try:
                    self.disable_signature_new_chk.config(state=state)
                except Exception:
                    pass
        except Exception:
            pass

        self.stream_entry.delete(0, tk.END)
        self.stream_entry.insert(0, self.config.get('FIXED_STREAM', ''))

        self.country_entry.delete(0, tk.END)
        self.country_entry.insert(0, self.config.get('FIXED_COUNTRY', ''))

        # --- General (Period Logic) sync ---
        try:
            if hasattr(self, 'daily_logic_var'):
                self.daily_logic_var.set(self.config.get('DAILY_LOGIC', 'Delayed (Day-1)'))
        except Exception:
            pass
        try:
            if hasattr(self, 'weekly_logic_var'):
                self.weekly_logic_var.set(self.config.get('WEEKLY_LOGIC', 'Delayed (ISO-1)'))
        except Exception:
            pass
        try:
            if hasattr(self, 'weekly_start_offset_var'):
                self.weekly_start_offset_var.set(int(self.config.get('WEEKLY_START_OFFSET_DAYS', 0) or 0))
        except Exception:
            pass
        try:
            if hasattr(self, 'monthly_logic_var'):
                self.monthly_logic_var.set(self.config.get('MONTHLY_LOGIC', 'Delayed (Month-1)'))
        except Exception:
            pass
        try:
            if hasattr(self, 'monthly_start_day_var'):
                self.monthly_start_day_var.set(int(self.config.get('MONTHLY_START_DAY', 1) or 1))
        except Exception:
            pass
        try:
            self.update_daily_preview()
            self.update_weekly_preview()
            self.update_monthly_preview()
        except Exception:
            pass
        # --- Recipients (populate BOTH modes even if a field is currently disabled) ---
        def _set_entry(_w, _v):
            try:
                _old = None
                try: _old = str(_w.cget('state'))
                except Exception: _old = None
                try:
                    if _old == 'disabled':
                        _w.config(state='normal')
                except Exception: pass
                try:
                    _w.delete(0, tk.END)
                    _w.insert(0, _v if _v is not None else '')
                except Exception: pass
                try:
                    if _old == 'disabled':
                        _w.config(state='disabled')
                except Exception: pass
            except Exception: pass

        def _set_text(_w, _v):
            try:
                _old = None
                try: _old = str(_w.cget('state'))
                except Exception: _old = None
                try:
                    if _old == 'disabled':
                        _w.config(state='normal')
                except Exception: pass
                try:
                    _w.delete('1.0', tk.END)
                    _w.insert('1.0', _v if _v is not None else '')
                except Exception: pass
                try:
                    if _old == 'disabled':
                        _w.config(state='disabled')
                except Exception: pass
            except Exception: pass

        try:
            if hasattr(self, 'to_entry'):
                _set_text(self.to_entry, self.config.get('RECIPIENTS_TO', ''))
        except Exception: pass
        try:
            if hasattr(self, 'to_new_entry'):
                _set_text(self.to_new_entry, self.config.get('RECIPIENTS_TO_NEW', ''))
        except Exception: pass
        try:
            if hasattr(self, 'cc_entry'):
                _set_text(self.cc_entry, self.config.get('RECIPIENTS_CC', ''))
        except Exception: pass
        try:
            if hasattr(self, 'cc_new_entry'):
                _set_text(self.cc_new_entry, self.config.get('RECIPIENTS_CC_NEW', ''))
        except Exception: pass
        try:
            start_time_obj = datetime.datetime.strptime(self.config.get('START_TIME'), "%I:%M%p")
            self.start_hour_var.set(start_time_obj.strftime("%I"))
            self.start_minute_var.set(start_time_obj.strftime("%M"))
            self.start_ampm_var.set(start_time_obj.strftime("%p"))
        except Exception:
            pass
        try:
            end_time_obj = datetime.datetime.strptime(self.config.get('END_TIME'), "%I:%M%p")
            self.end_hour_var.set(end_time_obj.strftime("%I"))
            self.end_minute_var.set(end_time_obj.strftime("%M"))
            self.end_ampm_var.set(end_time_obj.strftime("%p"))
        except Exception:
            pass

        try:
            if hasattr(self, 'outlook_version_var'):
                self.outlook_version_var.set(str(self.config.get('OUTLOOK_VERSION','Classic') or 'Classic'))
        except Exception:
            pass

        try:
            self.config['UPDATE_MANIFEST_URL'] = HARDCODED_UPDATE_MANIFEST_URL
            if hasattr(self, 'update_manifest_entry'):
                try:
                    self.update_manifest_entry.config(state='normal')
                except Exception:
                    pass
                self.update_manifest_entry.delete(0, tk.END)
                self.update_manifest_entry.insert(0, HARDCODED_UPDATE_MANIFEST_URL)
                try:
                    self.update_manifest_entry.config(state='readonly')
                except Exception:
                    pass
            if hasattr(self, 'auto_check_updates_var'):
                self.auto_check_updates_var.set(bool(self.config.get('AUTO_CHECK_UPDATES', False)))
            if hasattr(self, 'update_status_label'):
                skipped = self.config.get('SKIPPED_UPDATE_VERSION', '') or '(none)'
                last_check = self.config.get('LAST_UPDATE_CHECK', '') or '(never)'
                last_seen = self.config.get('LAST_UPDATE_VERSION_SEEN', '') or '(none)'
                self.update_status_label.config(text=f'Last check: {last_check} | Last seen version: {last_seen} | Skipped: {skipped}')
        except Exception:
            pass

        try:
            self._wire_to_tab_behavior()
        except Exception:
            pass
        try:
            self._wire_recipient_hover_behavior()
        except Exception:
            pass
        try:
            self._refresh_recipient_inputs_state()
        except Exception:
            pass

    def _on_settings_tab_changed(self, event=None):

        """Keep Settings UI elements in sync when switching tabs (esp. Signature toggle)."""

        # Ensure Signature checkbox state reflects SAVED Outlook version

        try:

            if hasattr(self, 'disable_signature_new_chk'):

                state = 'normal' if self._is_new_outlook() else 'disabled'

                self.disable_signature_new_chk.config(state=state)

        except Exception:

            pass

        # If user navigates to Signature tab, ensure checkbox value reflects config

        try:

            nb = getattr(self, 'settings_notebook', None)

            if nb is not None:

                tab_text = str(nb.tab(nb.select(), 'text') or '')

            else:

                tab_text = ''

            if tab_text == 'Signature' and hasattr(self, 'disable_signature_new_var'):

                self.disable_signature_new_var.set(bool(self.config.get('NEW_OUTLOOK_SIGNATURE_DISABLED', False)))

        except Exception:

            pass

    def create_settings_widgets(self):
        self.settings_frame = ttk.Frame(self)
        settings_notebook = ttk.Notebook(self.settings_frame)
        self.settings_notebook = settings_notebook
        settings_notebook.pack(pady=10, padx=10, fill="both", expand=True)
        try:
            settings_notebook.bind('<<NotebookTabChanged>>', self._on_settings_tab_changed)
        except Exception:
            pass


        tabs = {name: ttk.Frame(settings_notebook) for name in ["General", "Name", "Signature", "Stream", "Country", "Schedule", "To", "CC", "Outlook Version"]}
        for name, tab in tabs.items():
            settings_notebook.add(tab, text=name)

        # --- General tab ---
        ttk.Label(tabs["General"], text="Daily Period Logic:").pack(padx=10, pady=(10, 5), anchor="w")
        self.daily_logic_var = tk.StringVar()
        try:
            self.daily_logic_var.set(self.config.get('DAILY_LOGIC', 'Delayed (Day-1)'))
        except Exception:
            pass
        ttk.Combobox(
            tabs["General"],
            textvariable=self.daily_logic_var,
            values=['Delayed (Day-1)', 'Current (Day)'],
            state="readonly",
            width=30
        ).pack(padx=10, pady=0, anchor="w")
        self.daily_preview_label = ttk.Label(tabs["General"], text="Preview (today): —", foreground="#1F4E79")
        self.daily_preview_label.pack(padx=10, pady=(5, 0), anchor="w")
        self.update_daily_preview()
        self.daily_logic_var.trace_add('write', lambda *_: self.update_daily_preview())
        self.daily_logic_var.trace_add('write', lambda *_: self._refresh_sod_task_periods_preview())


        ttk.Label(tabs["General"], text="Weekly Period Logic:").pack(padx=10, pady=(15, 5), anchor="w")
        self.weekly_logic_var = tk.StringVar()
        try:
            self.weekly_logic_var.set(self.config.get('WEEKLY_LOGIC', 'Delayed (ISO-1)'))
        except Exception:
            pass
        ttk.Combobox(
            tabs["General"],
            textvariable=self.weekly_logic_var,
            values=['Delayed (ISO-1)', 'Current (ISO)'],
            state="readonly",
            width=30
        ).pack(padx=10, pady=0, anchor="w")

        ttk.Label(tabs["General"], text="Week Start Offset (days from Monday):").pack(padx=10, pady=(15, 5), anchor="w")
        self.weekly_start_offset_var = tk.IntVar(value=int(self.config.get('WEEKLY_START_OFFSET_DAYS', 0)))
        weekly_offset_frame = ttk.Frame(tabs["General"])
        weekly_offset_frame.pack(padx=10, pady=0, anchor="w")
        ttk.Spinbox(weekly_offset_frame, from_=-6, to=6, textvariable=self.weekly_start_offset_var, width=5, command=self.update_weekly_preview).pack(side="left")
        ttk.Label(weekly_offset_frame, text="  (0=Mon, 4=Fri, 6=Sun; negatives go backward)").pack(side="left")
        self.weekly_preview_label = ttk.Label(tabs["General"], text="Preview (today): —", foreground="#1F4E79")
        self.weekly_preview_label.pack(padx=10, pady=(5, 0), anchor="w")
        self.update_weekly_preview()
        self.weekly_logic_var.trace_add('write', lambda *_: self.update_weekly_preview())
        self.weekly_logic_var.trace_add('write', lambda *_: self._refresh_sod_task_periods_preview())
        self.weekly_start_offset_var.trace_add('write', lambda *_: self.update_weekly_preview())
        self.weekly_start_offset_var.trace_add('write', lambda *_: self._refresh_sod_task_periods_preview())

        ttk.Label(tabs["General"], text="Monthly Period Logic:").pack(padx=10, pady=(15, 5), anchor="w")
        self.monthly_logic_var = tk.StringVar()
        try:
            self.monthly_logic_var.set(self.config.get('MONTHLY_LOGIC', 'Delayed (Month-1)'))
        except Exception:
            pass
        ttk.Combobox(
            tabs["General"],
            textvariable=self.monthly_logic_var,
            values=['Delayed (Month-1)', 'Current (Month)'],
            state="readonly",
            width=30
        ).pack(padx=10, pady=0, anchor="w")

        ttk.Label(tabs["General"], text="Monthly Start Day (period flips on this day):").pack(padx=10, pady=(15, 5), anchor="w")
        self.monthly_start_day_var = tk.IntVar(value=int(self.config.get('MONTHLY_START_DAY', 1)))
        ttk.Spinbox(tabs["General"], from_=1, to=31, textvariable=self.monthly_start_day_var, width=5).pack(padx=10, pady=0, anchor="w")
        self.monthly_preview_label = ttk.Label(tabs["General"], text="Preview (today): —", foreground="#1F4E79")
        self.monthly_preview_label.pack(padx=10, pady=(5, 0), anchor="w")
        self.update_monthly_preview()
        self.monthly_logic_var.trace_add('write', lambda *_: self.update_monthly_preview())
        self.monthly_logic_var.trace_add('write', lambda *_: self._refresh_sod_task_periods_preview())
        self.monthly_start_day_var.trace_add('write', lambda *_: self.update_monthly_preview())
        self.monthly_start_day_var.trace_add('write', lambda *_: self._refresh_sod_task_periods_preview())


        ttk.Button(tabs["General"], text="Save General Settings", command=self.save_period_logic).pack(pady=20)

        # --- Name tab ---
        ttk.Label(tabs["Name"], text="Your Full Name:").pack(padx=10, pady=5)
        self.name_entry = ttk.Entry(tabs["Name"], width=50)
        self.name_entry.pack(padx=10, pady=5)
        ttk.Button(tabs["Name"], text="Save Name", command=self.save_name).pack(pady=10)

        # --- Signature tab ---

        # New Outlook only: optionally disable tool signature (use Outlook preset signature instead)
        self.disable_signature_new_var = tk.BooleanVar(value=bool(self.config.get('NEW_OUTLOOK_SIGNATURE_DISABLED', False)))
        self.disable_signature_new_chk = ttk.Checkbutton(
            tabs["Signature"],
            text="New Outlook: Disable tool signature (use Outlook preset signature)",
            variable=self.disable_signature_new_var
        )
        self.disable_signature_new_chk.pack(padx=10, pady=(8, 0), anchor='w')

        ttk.Label(tabs["Signature"], text="Signature Name (bolded):").pack(padx=10, pady=5)
        self.signature_entry = ttk.Entry(tabs["Signature"], width=50)
        self.signature_entry.pack(padx=10, pady=5)
        ttk.Button(tabs["Signature"], text="Save Signature", command=self.save_signature).pack(pady=10)

        # --- Stream tab ---
        ttk.Label(tabs["Stream"], text="Stream (e.g., CH):").pack(padx=10, pady=5)
        self.stream_entry = ttk.Entry(tabs["Stream"], width=50)
        self.stream_entry.pack(padx=10, pady=5)
        ttk.Button(tabs["Stream"], text="Save Stream", command=self.save_stream).pack(pady=10)

        # --- Country tab ---
        ttk.Label(tabs["Country"], text="Country:").pack(padx=10, pady=5)
        self.country_entry = ttk.Entry(tabs["Country"], width=50)
        self.country_entry.pack(padx=10, pady=5)
        ttk.Button(tabs["Country"], text="Save Country", command=self.save_country).pack(pady=10)

        # --- Schedule tab ---
        hours = [str(i).zfill(2) for i in range(1, 13)]
        minutes = [str(i).zfill(2) for i in range(60)]

        start_frame = ttk.Frame(tabs["Schedule"])
        start_frame.pack(pady=5)
        ttk.Label(start_frame, text="Start Time:").pack(side="left", padx=5)
        self.start_hour_var, self.start_minute_var, self.start_ampm_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        # Initialize Start Time controls from saved config
        try:
            _st = datetime.datetime.strptime(self.config.get('START_TIME'), "%I:%M%p")
            self.start_hour_var.set(_st.strftime("%I"))
            self.start_minute_var.set(_st.strftime("%M"))
            self.start_ampm_var.set(_st.strftime("%p"))
        except Exception:
            pass
        ttk.Combobox(start_frame, textvariable=self.start_hour_var, values=hours, width=3).pack(side="left")
        ttk.Combobox(start_frame, textvariable=self.start_minute_var, values=minutes, width=3).pack(side="left")
        ttk.Combobox(start_frame, textvariable=self.start_ampm_var, values=['AM', 'PM'], width=3).pack(side="left")

        end_frame = ttk.Frame(tabs["Schedule"])
        end_frame.pack(pady=5)
        ttk.Label(end_frame, text="End Time:  ").pack(side="left", padx=5)
        self.end_hour_var, self.end_minute_var, self.end_ampm_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        ttk.Combobox(end_frame, textvariable=self.end_hour_var, values=hours, width=3).pack(side="left")
        ttk.Combobox(end_frame, textvariable=self.end_minute_var, values=minutes, width=3).pack(side="left")
        ttk.Combobox(end_frame, textvariable=self.end_ampm_var, values=['AM', 'PM'], width=3).pack(side="left")

        ttk.Button(tabs["Schedule"], text="Save Schedule", command=self.save_schedule).pack(pady=10)

        # --- To tab ---

        # Classic Outlook window (Names) - expects "Surname, First Name" format

        ttk.Label(tabs["To"], text="Classic Outlook To (Names)", font=('Segoe UI', 11, 'bold')).pack(padx=10, pady=(12, 6), anchor='w')

        ttk.Label(tabs["To"], text="Note: Use 'Surname, First Name' (e.g., Doe, John; Smith, Jane)", foreground='#6B6B6B').pack(padx=10, pady=(0, 8), anchor='w')

        

        # Blue border container (turns grey when New Outlook is selected)

        self.to_classic_border = tk.Frame(tabs["To"], bg='#1F4E79')

        self.to_classic_border.pack(padx=10, pady=(0, 8), fill='x')

        

        self.to_classic_frame = ttk.LabelFrame(self.to_classic_border, text="Classic Outlook (Names)")

        try:

            self.to_classic_frame.configure(bootstyle='primary')

        except Exception:

            pass

        self.to_classic_frame.pack(padx=2, pady=2, fill='x')

        

        ttk.Label(self.to_classic_frame, text="To (names separated by ';')").pack(padx=8, pady=(8, 4), anchor='w')

        self.to_classic_wrap = ttk.Frame(self.to_classic_frame)
        self.to_classic_wrap.pack(padx=8, pady=(0, 10), fill='both', expand=False)
        self.to_entry = tk.Text(self.to_classic_wrap, height=2, width=80, wrap='word')
        self.to_entry.pack(side='left', fill='both', expand=True)

        

        ttk.Button(tabs["To"], text="Save To", command=self.save_to).pack(pady=(4, 10))


        

        # --- New Outlook To tab window ---

        

        ttk.Separator(tabs["To"], orient='horizontal').pack(fill='x', padx=10, pady=(6, 8))

        

        ttk.Label(tabs["To"], text="New Outlook To (Recipients)", font=('Segoe UI', 11, 'bold')).pack(padx=10, pady=(0, 6), anchor='w')

        

        ttk.Label(tabs["To"], text="Format: firstname.lastname@iqvia.com (e.g., john.doe@iqvia.com; jane.smith@iqvia.com). Separate by ';' or new line", foreground='#6B6B6B').pack(padx=10, pady=(0, 8), anchor='w')

        

        

        

        # Blue border container (turns grey when Classic Outlook is selected)

        

        self.to_new_border = tk.Frame(tabs["To"], bg='#C0C0C0')

        

        self.to_new_border.pack(padx=10, pady=(0, 8), fill='x')

        

        

        

        self.to_new_frame = ttk.LabelFrame(self.to_new_border, text="New Outlook (Recipients)")

        

        try:

        

            self.to_new_frame.configure(bootstyle='primary')

        

        except Exception:

        

            pass

        

        self.to_new_frame.pack(padx=2, pady=2, fill='x')

        

        

        

        ttk.Label(self.to_new_frame, text="To (emails separated by ';' or new line)").pack(padx=8, pady=(8, 4), anchor='w')

        

        self.to_new_wrap = ttk.Frame(self.to_new_frame)

        

        self.to_new_wrap.pack(padx=8, pady=(0, 8), fill='both', expand=False)

        

        self.to_new_entry = tk.Text(self.to_new_wrap, height=3, width=80, wrap='word')

        

        self.to_new_entry.pack(side='left', fill='both', expand=True)

        

        self.to_new_scroll = ttk.Scrollbar(self.to_new_wrap, orient='vertical', command=self.to_new_entry.yview)

        

        self.to_new_scroll.pack(side='right', fill='y')

        

        try:

        

            self.to_new_entry.configure(yscrollcommand=self.to_new_scroll.set)

        

        except Exception:

        

            pass

        

        try:

        

            self.to_new_entry.config(font=('Segoe UI', 10))

        

        except Exception:

        

            pass

        

        

        

        self.to_new_save_btn = ttk.Button(tabs["To"], text="Save To (New Outlook)", command=self.save_to_new)

        

        self.to_new_save_btn.pack(pady=(4, 10))

        # Classic Outlook window (Names) - expects "Surname, First Name" format

        
        ttk.Label(tabs["CC"], text="Classic Outlook CC (Names)", font=('Segoe UI', 11, 'bold')).pack(padx=10, pady=(12, 6), anchor='w')

        
        ttk.Label(tabs["CC"], text="Note: Use 'Surname, First Name' (e.g., Doe, John; Smith, Jane; Brown, Alex)", foreground='#6B6B6B').pack(padx=10, pady=(0, 8), anchor='w')

        
        

        
        # Blue border container (turns grey when New Outlook is selected)

        
        self.cc_classic_border = tk.Frame(tabs["CC"], bg='#1F4E79')

        
        self.cc_classic_border.pack(padx=10, pady=(0, 8), fill='both')

        
        

        
        self.cc_classic_frame = ttk.LabelFrame(self.cc_classic_border, text="Classic Outlook (Names)")

        
        try:

        
            self.cc_classic_frame.configure(bootstyle='primary')

        
        except Exception:

        
            pass

        
        self.cc_classic_frame.pack(padx=2, pady=2, fill='both')

        
        

        
        ttk.Label(self.cc_classic_frame, text="CC (names separated by ';' or new line)").pack(padx=8, pady=(8, 4), anchor='w')

        
        self.cc_classic_wrap = ttk.Frame(self.cc_classic_frame)

        
        self.cc_classic_wrap.pack(padx=8, pady=(0, 8), fill='both', expand=False)

        
        self.cc_entry = tk.Text(self.cc_classic_wrap, height=4, width=80, wrap='word')

        
        self.cc_entry.pack(side='left', fill='both', expand=True)

        
        self.cc_classic_scroll = ttk.Scrollbar(self.cc_classic_wrap, orient='vertical', command=self.cc_entry.yview)

        
        self.cc_classic_scroll.pack(side='right', fill='y')

        
        try:

        
            self.cc_entry.configure(yscrollcommand=self.cc_classic_scroll.set)

        
        except Exception:

        
            pass

        
        

        
        ttk.Button(tabs["CC"], text="Save CC", command=self.save_cc).pack(pady=(4, 10))



        
        

        
        # --- New Outlook CC tab window ---

        
        

        
        ttk.Separator(tabs["CC"], orient='horizontal').pack(fill='x', padx=10, pady=(6, 8))

        
        

        
        ttk.Label(tabs["CC"], text="New Outlook CC (Recipients)", font=('Segoe UI', 11, 'bold')).pack(padx=10, pady=(0, 6), anchor='w')

        
        

        
        ttk.Label(tabs["CC"], text="Format: firstname.lastname@iqvia.com (e.g., john.doe@iqvia.com; jane.smith@iqvia.com). Separate by ';' or new line", foreground='#6B6B6B').pack(padx=10, pady=(0, 8), anchor='w')

        
        

        
        

        
        

        
        # Blue border container (turns grey when Classic Outlook is selected)

        
        

        
        self.cc_new_border = tk.Frame(tabs["CC"], bg='#C0C0C0')

        
        

        
        self.cc_new_border.pack(padx=10, pady=(0, 8), fill='x')

        
        

        
        

        
        

        
        self.cc_new_frame = ttk.LabelFrame(self.cc_new_border, text="New Outlook (Recipients)")

        
        

        
        try:

        
        

        
            self.cc_new_frame.configure(bootstyle='primary')

        
        

        
        except Exception:

        
        

        
            pass

        
        

        
        self.cc_new_frame.pack(padx=2, pady=2, fill='x')

        
        

        
        

        
        

        
        ttk.Label(self.cc_new_frame, text="CC (emails separated by ';' or new line)").pack(padx=8, pady=(8, 4), anchor='w')

        
        

        
        self.cc_new_wrap = ttk.Frame(self.cc_new_frame)

        
        

        
        self.cc_new_wrap.pack(padx=8, pady=(0, 8), fill='both', expand=False)

        
        

        
        self.cc_new_entry = tk.Text(self.cc_new_wrap, height=4, width=80, wrap='word')

        
        

        
        self.cc_new_entry.pack(side='left', fill='both', expand=True)

        
        

        
        self.cc_new_scroll = ttk.Scrollbar(self.cc_new_wrap, orient='vertical', command=self.cc_new_entry.yview)

        
        

        
        self.cc_new_scroll.pack(side='right', fill='y')

        
        

        
        try:

        
        

        
            self.cc_new_entry.configure(yscrollcommand=self.cc_new_scroll.set)

        
        

        
        except Exception:

        
        

        
            pass

        
        

        
        try:

        
        

        
            self.cc_new_entry.config(font=('Segoe UI', 10))

        
        

        
        except Exception:

        
        

        
            pass

        
        

        
        

        
        

        
        self.cc_new_save_btn = ttk.Button(tabs["CC"], text="Save CC (New Outlook)", command=self.save_cc_new)

        
        

        
        self.cc_new_save_btn.pack(pady=(4, 10))

        # --- Outlook Version tab ---

        ttk.Label(tabs["Outlook Version"], text="Select Outlook Version:").pack(padx=10, pady=(15, 8), anchor='w')
        self.outlook_version_var = tk.StringVar(value=str(self.config.get('OUTLOOK_VERSION','Classic') or 'Classic'))
        ttk.Radiobutton(tabs["Outlook Version"], text="Classic Outlook (Desktop)", value="Classic", variable=self.outlook_version_var).pack(padx=20, pady=5, anchor='w')
        ttk.Radiobutton(tabs["Outlook Version"], text="New Outlook (Copy/Paste)", value="New", variable=self.outlook_version_var).pack(padx=20, pady=5, anchor='w')
        ttk.Label(tabs["Outlook Version"], text="New Outlook mode uses Copy buttons and disables screenshots.", foreground="#1F4E79").pack(padx=10, pady=(10, 0), anchor='w')
        ttk.Button(tabs["Outlook Version"], text="Save Outlook Version", command=self.save_outlook_version).pack(pady=20)

# Navigation handled via header buttons
    def save_period_logic(self):
        self.config['WEEKLY_LOGIC'] = self.weekly_logic_var.get()
        try:
            self.config['WEEKLY_START_OFFSET_DAYS'] = int(self.weekly_start_offset_var.get())
        except Exception:
            self.config['WEEKLY_START_OFFSET_DAYS'] = 0
        self.config['DAILY_LOGIC'] = self.daily_logic_var.get()
        self.config['MONTHLY_LOGIC'] = self.monthly_logic_var.get()
        try:
            self.config['MONTHLY_START_DAY'] = int(self.monthly_start_day_var.get())
        except Exception:
            self.config['MONTHLY_START_DAY'] = 1
        self.save_config()
        try:
            self._refresh_sod_task_periods()
        except Exception:
            pass
    def save_name(self): self.config['YOUR_NAME'] = self.name_entry.get(); self.save_config()
    def save_signature(self):
        self.config['SIGNATURE_NAME'] = self.signature_entry.get()
        try:
            if hasattr(self, 'disable_signature_new_var'):
                self.config['NEW_OUTLOOK_SIGNATURE_DISABLED'] = bool(self.disable_signature_new_var.get())
        except Exception:
            pass
        self.save_config()
    def save_stream(self): self.config['FIXED_STREAM'] = self.stream_entry.get(); self.save_config()
    def save_country(self): self.config['FIXED_COUNTRY'] = self.country_entry.get(); self.save_config()
    def save_schedule(self):
        self.config['START_TIME'] = f"{self.start_hour_var.get()}:{self.start_minute_var.get()}{self.start_ampm_var.get()}"
        self.config['END_TIME'] = f"{self.end_hour_var.get()}:{self.end_minute_var.get()}{self.end_ampm_var.get()}"
        self.save_config()
    def save_to(self):
        """Save To recipients (Classic Outlook names only)."""
        try:
            self.config['RECIPIENTS_TO'] = self.to_entry.get('1.0', 'end-1c')
        except Exception:
            pass
        self.save_config()

    def save_to_new(self):
        """Save To recipients for New Outlook (emails/paste list)."""
        try:
            self.config['RECIPIENTS_TO_NEW'] = self.to_new_entry.get('1.0', 'end-1c')
        except Exception:
            pass
        self.save_config()
    def save_cc(self):
        """Save CC recipients (Classic Outlook names only)."""
        try:
            self.config['RECIPIENTS_CC'] = self.cc_entry.get('1.0', 'end-1c')
        except Exception:
            pass
        self.save_config()

    def save_cc_new(self):
        """Save CC recipients for New Outlook (emails/paste list)."""
        try:
            self.config['RECIPIENTS_CC_NEW'] = self.cc_new_entry.get('1.0', 'end-1c')
        except Exception:
            pass
        self.save_config()
    def _on_outlook_version_changed(self):
        """Live-apply Outlook Version changes (no restart)."""
        # Defer to end of event loop so Tk applies the radio variable first
        try:
            self.after(0, self._apply_outlook_version_changed_now)
        except Exception:
            try:
                self._apply_outlook_version_changed_now()
            except Exception:
                pass

    def _apply_outlook_version_changed_now(self):
        """Internal: run the actual refresh work."""
        try:
            self._refresh_recipient_inputs_state()
        except Exception:
            pass
        try:
            self._apply_outlook_version_ui()
        except Exception:
            pass
        try:
            if hasattr(self, 'disable_signature_new_chk'):
                state = 'normal' if self._is_new_outlook() else 'disabled'
                self.disable_signature_new_chk.config(state=state)
        except Exception:
            pass
        try:
            # Force border redraw
            if hasattr(self, 'to_new_border'): self.to_new_border.update_idletasks()
            if hasattr(self, 'cc_new_border'): self.cc_new_border.update_idletasks()
            self.update_idletasks()
        except Exception:
            pass

    def save_outlook_version(self):

        """Save Outlook version selection (Classic/New) and apply UI changes."""

        try:

            val = str(getattr(self, 'outlook_version_var', tk.StringVar(value='Classic')).get() or 'Classic').strip()

        except Exception:

            val = 'Classic'

        self.config['OUTLOOK_VERSION'] = 'New' if val.lower().startswith('new') else 'Classic'

        self.save_config()

        # Apply UI based on SAVED config immediately

        try:

            self._apply_outlook_version_ui()

        except Exception:

            pass

        try:

            self._refresh_recipient_inputs_state()

        except Exception:

            pass

        # Keep signature toggle enabled/disabled in sync with Outlook mode (no restart)

        try:

            if hasattr(self, 'disable_signature_new_chk'):

                state = 'normal' if self._is_new_outlook() else 'disabled'

                self.disable_signature_new_chk.config(state=state)

        except Exception:

            pass
    def create_management_widgets(self):
        self.management_frame = ttk.Frame(self)
        ttk.Label(self.management_frame, text="Management", font="-size 16 -weight bold").pack(pady=20)
        ttk.Button(self.management_frame, text="Edit Application Settings", command=self.show_settings_frame).pack(pady=10, ipadx=10, ipady=5)
        ttk.Button(self.management_frame, text="Edit Daily Presets", command=lambda: self.show_frame(self.presets_frame)).pack(pady=10, ipadx=10, ipady=5)
# Navigation handled via header buttons
    # -----------------------------
    # TASK MOVE
    # -----------------------------


    def create_help_widgets(self):
        self.help_frame = ttk.Frame(self)

        ttk.Label(self.help_frame, text='Help', font=('Segoe UI', 16, 'bold')).pack(pady=(15, 5))
        ttk.Label(self.help_frame, text='Instructions and app information', foreground='#6B6B6B').pack(pady=(0, 10))

        nb = ttk.Notebook(self.help_frame)
        nb.pack(padx=10, pady=10, fill='both', expand=True)

        tab_instr = ttk.Frame(nb)
        tab_about = ttk.Frame(nb)
        tab_updates = ttk.Frame(nb)
        nb.add(tab_instr, text='Instructions')
        nb.add(tab_about, text='About')
        nb.add(tab_updates, text='Updates')

        def _make_readonly_text(parent, content):
            wrap = ttk.Frame(parent)
            wrap.pack(fill='both', expand=True, padx=10, pady=10)
            txt = tk.Text(wrap, wrap='word')
            txt.pack(side='left', fill='both', expand=True)
            sb = ttk.Scrollbar(wrap, orient='vertical', command=txt.yview)
            sb.pack(side='right', fill='y')
            try:
                txt.configure(yscrollcommand=sb.set)
            except Exception:
                pass
            try:
                txt.configure(font=('Segoe UI', 10))
            except Exception:
                pass
            txt.insert('1.0', content)
            txt.config(state='disabled')
            return txt

        _make_readonly_text(tab_instr, 'Welcome! This tool helps you create accurate Start-of-Day (SOD) and End-of-Day (EOD) email drafts quickly and consistently.\n\nQUICK START (Most common flow)\n1) Home → Start of Day (SOD)\n   • Click “Load Preset” to load today’s tasks automatically.\n   • Review tasks, then click “Prepare SOD Email Draft” (Classic Outlook) OR use “Copy Body/Subject” (New Outlook).\n\n2) Home → End of Day (EOD)\n   • Click “Load Tasks” to load the tasks you prepared in SOD.\n   • Mark each task status (Done / In Progress / Carried Over).\n   • Add optional attendance screenshot (Classic Outlook only).\n   • Click “Prepare EOD Email Draft” OR use Copy buttons (New Outlook).\n\nHOME PAGE TABS\nA) Start of Day (SOD)\n• Add Task: Enter Tasklist + Stream + Frequency (optional) + Start/End times (optional).\n• Saved Tasks dropdown: choose from “Task Dropdown Options” (Management → Edit Daily Presets).\n• Load Preset: loads Daily + today’s Weekday + today’s Monthly tasks.\n• Load Unfinished Tasks: loads tasks from the latest EOD where status was In Progress / Carried Over.\n• Reorder: use ▲/▼ buttons.\n• Edit: double‑click a task row to edit values.\n• Remove: click the ❌ Action column.\n\nB) End of Day (EOD)\n• Load Tasks: loads tasks from the most recent SOD snapshot.\n• Set Status: click Done / In Progress / Carried Over.\n• Screenshot: Paste/Browse (disabled in New Outlook mode).\n• Prepare Draft / Copy: build the EOD email content.\n\nC) OT In / OT Out\n• OT In: select OT date, set OT From/To, add OT tasks and status.\n• OT Out: load OT In tasks, set actual end time, update status, then generate draft.\n\nMANAGEMENT\n1) Edit Application Settings\n   • General: controls how Period is calculated for Daily/Weekly/Monthly.\n   • Name / Signature / Stream / Country / Schedule\n   • To / CC: classic recipients (names) and New Outlook recipients (emails).\n   • Outlook Version: choose Classic vs New Outlook behavior.\n\n2) Edit Daily Presets\n   • Daily: always daily tasks.\n   • Weekday: tasks per weekday.\n   • Monthly: tasks per day‑of‑month.\n   • Task Dropdown Options: saved tasks shown under SOD “Saved Tasks”. You can also assign a default frequency per task.\n\nOUTLOOK MODES\n• Classic Outlook (Desktop)\n  – Uses Outlook automation to create a draft automatically.\n  – Supports Attendance screenshot attachment in EOD.\n\n• New Outlook (Copy/Paste)\n  – Use Copy buttons (Copy Body / Copy Subject / Copy To / Copy CC).\n  – Screenshots are disabled.\n  – New Outlook To/CC expects: firstname.lastname@iqvia.com\n\nTIPS\n• Period is computed from Frequency using General settings.\n• If you edit Task Dropdown Options frequencies, existing listed tasks update accordingly.\n\nTROUBLESHOOTING\n• If Classic Outlook draft creation fails, ensure Outlook Desktop is installed and running.\n• If HTML paste looks off in New Outlook, paste normally (Ctrl+V) into the message body.\n')
        _make_readonly_text(tab_about, 'Prod_task_generator\n\nThis app streamlines daily reporting by generating structured SOD, EOD, and OT email drafts.\n\nCreator\n• Allen Paul Olguera\n• Vibe coder 😄 — the idea, direction, testing, and compilation were done by Allen Paul Olguera.\n• Implementation is AI‑assisted and hard‑coded into this tool following the creator’s specifications.\n\nSupport / Suggestions\nIf you find a bug or have improvement suggestions, please contact:\nallenpaul.olguera@iqvia.com\n')

        # --- Updates tab inside Help ---
        ttk.Label(tab_updates, text='Update Settings', font=('Segoe UI', 11, 'bold')).pack(padx=10, pady=(15, 5), anchor='w')
        ttk.Label(tab_updates, text='Hardcoded GitHub Release API URL:').pack(padx=10, pady=(5, 5), anchor='w')
        self.update_manifest_entry = ttk.Entry(tab_updates, width=90)
        self.update_manifest_entry.pack(padx=10, pady=(0, 8), fill='x')
        try:
            self.update_manifest_entry.insert(0, self.config.get('UPDATE_MANIFEST_URL', ''))
        except Exception:
            pass
        ttk.Label(tab_updates, text='Supports either a custom manifest.json URL or the GitHub Releases latest API URL.', foreground='#6B6B6B').pack(padx=10, pady=(0, 8), anchor='w')
        self.auto_check_updates_var = tk.BooleanVar(value=bool(self.config.get('AUTO_CHECK_UPDATES', False)))
        ttk.Checkbutton(tab_updates, text='Automatically check for updates on startup', variable=self.auto_check_updates_var).pack(padx=10, pady=(0, 10), anchor='w')
        self.update_status_label = ttk.Label(tab_updates, text='Last check: (never) | Last seen version: (none) | Skipped: (none)', foreground='#1F4E79')
        self.update_status_label.pack(padx=10, pady=(0, 10), anchor='w')
        updates_btn_row = ttk.Frame(tab_updates)
        updates_btn_row.pack(padx=10, pady=(5, 10), anchor='w')
        ttk.Button(updates_btn_row, text='Save Update Preferences', command=self.save_update_settings).pack(side='left')
        ttk.Button(updates_btn_row, text='Check for Updates', command=lambda: self.check_for_updates(show_no_update=True, startup=False)).pack(side='left', padx=(8, 0))
        ttk.Button(updates_btn_row, text='Clear Skipped Version', command=self.clear_skipped_update_version).pack(side='left', padx=(8, 0))
        try:
            self.refresh_update_ui()
        except Exception:
            pass
    def move_task(self, tree, direction):
        selected_item = tree.focus()
        if not selected_item:
            return
        tree.move(selected_item, tree.parent(selected_item), tree.index(selected_item) + direction)

    # -----------------------------
    # SOD UI
    # -----------------------------
    def create_sod_widgets(self):
        sod_input_frame = ttk.LabelFrame(self.sod_frame, text="Add New Task")
        sod_input_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
        sod_input_frame.columnconfigure(1, weight=1)

        ttk.Label(sod_input_frame, text="Tasklist:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.tasklist_entry = ttk.Entry(sod_input_frame)
        self.tasklist_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=2)

        ttk.Label(sod_input_frame, text="Saved Tasks:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.task_helper_var = tk.StringVar()
        self.task_helper_combo = ttk.Combobox(
            sod_input_frame,
            textvariable=self.task_helper_var,
            values=self.get_task_dropdown_display_list(),
            state="readonly"
        )
        self.task_helper_combo.grid(row=1, column=1, sticky="ew", padx=5, pady=0)
        self.task_helper_combo.bind("<<ComboboxSelected>>", self.on_task_helper_select)

        ttk.Label(sod_input_frame, text="Frequency:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.frequency_var = tk.StringVar()
        self.frequency_combo = ttk.Combobox(
            sod_input_frame,
            textvariable=self.frequency_var,
            values=['', 'Daily', 'Weekly', 'Monthly', 'Weekly, Monthly'],
            width=15
        )
        self.frequency_combo.grid(row=2, column=1, sticky="w", padx=5, pady=2)
        # Start/End time pickers (optional)
        time_hours = [str(i).zfill(2) for i in range(1, 13)]
        time_minutes = [str(i).zfill(2) for i in range(60)]

        ttk.Label(sod_input_frame, text="Start Time:").grid(row=3, column=0, sticky="w", padx=5, pady=2)
        self.sod_start_hour_var, self.sod_start_minute_var, self.sod_start_ampm_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        st_frame = ttk.Frame(sod_input_frame)
        st_frame.grid(row=3, column=1, sticky="w", padx=5, pady=2)
        ttk.Combobox(st_frame, textvariable=self.sod_start_hour_var, values=time_hours, width=3, state="readonly").pack(side="left")
        ttk.Combobox(st_frame, textvariable=self.sod_start_minute_var, values=time_minutes, width=3, state="readonly").pack(side="left", padx=(5,0))
        ttk.Combobox(st_frame, textvariable=self.sod_start_ampm_var, values=['AM','PM'], width=3, state="readonly").pack(side="left", padx=(5,0))

        ttk.Label(sod_input_frame, text="End Time:").grid(row=4, column=0, sticky="w", padx=5, pady=2)
        self.sod_end_hour_var, self.sod_end_minute_var, self.sod_end_ampm_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        et_frame = ttk.Frame(sod_input_frame)
        et_frame.grid(row=4, column=1, sticky="w", padx=5, pady=2)
        ttk.Combobox(et_frame, textvariable=self.sod_end_hour_var, values=time_hours, width=3, state="readonly").pack(side="left")
        ttk.Combobox(et_frame, textvariable=self.sod_end_minute_var, values=time_minutes, width=3, state="readonly").pack(side="left", padx=(5,0))
        ttk.Combobox(et_frame, textvariable=self.sod_end_ampm_var, values=['AM','PM'], width=3, state="readonly").pack(side="left", padx=(5,0))

        ttk.Label(sod_input_frame, text="Stream:").grid(row=5, column=0, sticky="w", padx=5, pady=2)
        self.sod_stream_var = tk.StringVar()
        ttk.Entry(sod_input_frame, textvariable=self.sod_stream_var).grid(row=5, column=1, sticky="ew", padx=5, pady=2)
        ttk.Button(sod_input_frame, text="Add Task", command=self.add_task).grid(row=6, column=1, pady=10, sticky="w")


        

        sod_list_frame = ttk.LabelFrame(self.sod_frame, text="Today's Task List")
        sod_list_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=10, pady=5)
        self.sod_frame.rowconfigure(1, weight=1)
        self.sod_frame.columnconfigure(0, weight=1)

        self.sod_display_cols = ["Tasklist", "Stream", "Frequency", "Period", "Start Time", "End Time", "Issue Encountered", "Remarks", "Action"]
        self.sod_tree = ttk.Treeview(sod_list_frame, columns=self.sod_display_cols, show='headings')
        self.sod_tree.heading('Tasklist', text='Tasklist')
        self.sod_tree.column('Tasklist', width=280, minwidth=180, anchor='w', stretch=True)
        self.sod_tree.heading('Stream', text='Stream')
        self.sod_tree.column('Stream', width=90, minwidth=60, anchor='w', stretch=True)
        self.sod_tree.heading('Frequency', text='Frequency')
        self.sod_tree.column('Frequency', width=100, anchor='w')
        self.sod_tree.heading('Period', text='Period')
        self.sod_tree.column('Period', width=110, anchor='w')
        self.sod_tree.heading('Start Time', text='Start Time')
        self.sod_tree.column('Start Time', width=90, anchor='w')
        self.sod_tree.heading('End Time', text='End Time')
        self.sod_tree.column('End Time', width=90, anchor='w')
        self.sod_tree.heading('Issue Encountered', text='Issue Encountered')
        self.sod_tree.column('Issue Encountered', width=160, anchor='w')
        self.sod_tree.heading('Remarks', text='Status')
        self.sod_tree.column('Remarks', width=100, anchor='w')
        self.sod_tree.heading('Action', text='Action')
        self.sod_tree.column('Action', width=70, minwidth=70, anchor='center', stretch=False)
        # Make SOD columns resizable (drag separators) – keep Action fixed
        for _c in self.sod_display_cols:
            if _c == 'Action':
                continue
            try:
                self.sod_tree.column(_c, stretch=True, minwidth=80)
            except Exception:
                pass

        self.sod_tree.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        self.sod_tree.bind("<Double-1>", self.edit_task)
        self.sod_tree.bind("<ButtonRelease-1>", self.handle_sod_tree_click)

        sod_reorder_frame = ttk.Frame(sod_list_frame)
        sod_reorder_frame.pack(side="right", fill="y", padx=(0, 5))
        ttk.Button(sod_reorder_frame, text="▲", command=lambda: self.move_task(self.sod_tree, -1), width=3).pack(pady=2)
        ttk.Button(sod_reorder_frame, text="▼", command=lambda: self.move_task(self.sod_tree, 1), width=3).pack(pady=2)

        button_frame = ttk.Frame(self.sod_frame)
        button_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 5))
        # --- Actual Start Shift (optional override) ---
        shift_picker = ttk.Frame(button_frame)
        shift_picker.pack(side="top", fill="x", anchor="w", pady=(0, 5))
        ttk.Label(shift_picker, text="Actual Start Shift:").pack(side="left")
        self.sod_shift_hour_var, self.sod_shift_minute_var, self.sod_shift_ampm_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        hours = [str(i).zfill(2) for i in range(1, 13)]
        minutes = [str(i).zfill(2) for i in range(60)]
        ttk.Combobox(shift_picker, textvariable=self.sod_shift_hour_var, values=hours, width=3, state="readonly").pack(side="left", padx=(5, 0))
        ttk.Combobox(shift_picker, textvariable=self.sod_shift_minute_var, values=minutes, width=3, state="readonly").pack(side="left", padx=(5, 0))
        ttk.Combobox(shift_picker, textvariable=self.sod_shift_ampm_var, values=['AM','PM'], width=3, state="readonly").pack(side="left", padx=(5, 0))
        self.set_default_start_shift()

        buttons_row = ttk.Frame(button_frame)
        buttons_row.pack(side="top", fill="x")
        self.btn_prepare_sod = ttk.Button(buttons_row, text="Prepare SOD Email Draft", command=self.prepare_sod)
        self._btn_prepare_sod_pack = {'side': 'left'}
        self.btn_prepare_sod.pack(**self._btn_prepare_sod_pack)
        ttk.Button(buttons_row, text="Load Preset", command=self.load_preset_tasks).pack(side="left", padx=5)
        ttk.Button(buttons_row, text="Load Unfinished Tasks", command=self.load_unfinished_tasks).pack(side="left")

        # --- Copy bar (New Outlook mode) ---
        self.sod_copy_bar = ttk.LabelFrame(self.sod_frame, text='Copy (New Outlook)')
        self.sod_copy_bar.grid(row=3, column=0, columnspan=2, sticky='ew', padx=10, pady=(0, 8))
        self.sod_copy_bar.columnconfigure(0, weight=1)
        _bar = ttk.Frame(self.sod_copy_bar)
        _bar.grid(row=0, column=0, sticky='ew', padx=8, pady=6)
        self._populate_copy_bar(_bar, self.copy_sod_subject, self.copy_sod_body)
        # Hide in Classic mode (New Outlook mode will show it immediately)
        try:
            if not self._is_new_outlook():
                self.sod_copy_bar.grid_remove()
        except Exception:
            pass


    def on_task_helper_select(self, event):
        selected_task = self.task_helper_var.get()
        self.tasklist_entry.delete(0, tk.END)
        self.tasklist_entry.insert(0, selected_task)
        # If TaskDropdown has an assigned frequency, auto-fill it for faster entry
        try:
            mapped_freq = self.get_task_dropdown_frequency(selected_task)
            if mapped_freq is not None and str(mapped_freq).strip() != "":
                self.frequency_var.set(mapped_freq)
        except Exception:
            pass

    def handle_sod_tree_click(self, event):
        if self.sod_tree.identify_region(event.x, event.y) == "cell":
            if self.sod_tree.identify_column(event.x) == f'#{len(self.sod_display_cols)}':
                selected_item_id = self.sod_tree.focus()
                if selected_item_id:
                    self.remove_task(selected_item_id)
    def edit_task(self, event):
        # Ignore double click on Action column
        if self.sod_tree.identify_column(event.x) == f'#{len(self.sod_display_cols)}':
            return
        selected_item_id = self.sod_tree.focus()
        if not selected_item_id:
            return
        row = self.sod_full_data_storage.get(selected_item_id)
        if not row:
            return
        # row is (tasklist, stream, frequency, period, start, end, issue, remarks)
        if len(row) == 8:
            tasklist, stream, frequency, _period, start_time, end_time, issue, remarks = row
        else:
            tasklist = row[0] if len(row) > 0 else ''
            stream = ''
            frequency = row[1] if len(row) > 1 else ''
            start_time = row[3] if len(row) > 3 else ''
            end_time = row[4] if len(row) > 4 else ''
            issue = row[5] if len(row) > 5 else ''
            remarks = row[6] if len(row) > 6 else ''
        edit_window = tk.Toplevel(self)
        edit_window.title('Edit Task')
        ttk.Label(edit_window, text='Tasklist:').pack(padx=10, pady=(10, 0))
        task_entry = ttk.Entry(edit_window, width=60)
        task_entry.pack(padx=10, pady=5)
        task_entry.insert(0, tasklist)

        ttk.Label(edit_window, text='Stream:').pack(padx=10, pady=(10, 0))
        stream_var = tk.StringVar(value=stream)
        ttk.Entry(edit_window, width=30, textvariable=stream_var).pack(padx=10, pady=5, anchor='w')

        ttk.Label(edit_window, text='Frequency:').pack(padx=10, pady=(10, 0))
        freq_var = tk.StringVar(value=frequency)
        ttk.Combobox(edit_window, textvariable=freq_var, values=['', 'Daily', 'Weekly', 'Monthly', 'Weekly, Monthly']).pack(padx=10, pady=5)

        time_hours = [str(i).zfill(2) for i in range(1, 13)]
        time_minutes = [str(i).zfill(2) for i in range(60)]
        def split_time(t):
            try:
                dt = datetime.datetime.strptime(str(t).strip(), '%I:%M %p')
                return dt.strftime('%I'), dt.strftime('%M'), dt.strftime('%p')
            except Exception:
                return '', '', ''
        sh, sm, sa = split_time(start_time)
        eh, em, ea = split_time(end_time)

        ttk.Label(edit_window, text='Start Time:').pack(padx=10, pady=(10, 0))
        st_frame = ttk.Frame(edit_window)
        st_frame.pack(padx=10, pady=5, anchor='w')
        st_h=tk.StringVar(value=sh); st_m=tk.StringVar(value=sm); st_a=tk.StringVar(value=sa)
        ttk.Combobox(st_frame, textvariable=st_h, values=time_hours, width=3, state='readonly').pack(side='left')
        ttk.Combobox(st_frame, textvariable=st_m, values=time_minutes, width=3, state='readonly').pack(side='left', padx=(5,0))
        ttk.Combobox(st_frame, textvariable=st_a, values=['AM','PM'], width=3, state='readonly').pack(side='left', padx=(5,0))

        ttk.Label(edit_window, text='End Time:').pack(padx=10, pady=(10, 0))
        et_frame = ttk.Frame(edit_window)
        et_frame.pack(padx=10, pady=5, anchor='w')
        et_h=tk.StringVar(value=eh); et_m=tk.StringVar(value=em); et_a=tk.StringVar(value=ea)
        ttk.Combobox(et_frame, textvariable=et_h, values=time_hours, width=3, state='readonly').pack(side='left')
        ttk.Combobox(et_frame, textvariable=et_m, values=time_minutes, width=3, state='readonly').pack(side='left', padx=(5,0))
        ttk.Combobox(et_frame, textvariable=et_a, values=['AM','PM'], width=3, state='readonly').pack(side='left', padx=(5,0))

        def save_changes():
            new_task = task_entry.get().strip()
            if not new_task:
                messagebox.showwarning('Warning', 'Tasklist field cannot be empty.')
                return
            new_stream = (stream_var.get() or '').strip()
            new_freq = (freq_var.get() or '').strip()
            _fl2, _fd2 = normalize_frequency_string(new_freq)
            new_freq = _fd2 or new_freq
            new_period = self.calculate_period(new_freq)
            new_start = f"{st_h.get()}:{st_m.get()} {st_a.get()}" if (st_h.get() and st_m.get() and st_a.get()) else ''
            new_end = f"{et_h.get()}:{et_m.get()} {et_a.get()}" if (et_h.get() and et_m.get() and et_a.get()) else ''
            self.sod_tree.item(selected_item_id, values=(new_task, new_stream, new_freq, new_period, new_start, new_end, issue, remarks, '❌'))
            self.sod_full_data_storage[selected_item_id] = (new_task, new_stream, new_freq, new_period, new_start, new_end, issue, remarks)
            edit_window.destroy()
        ttk.Button(edit_window, text='Save Changes', command=save_changes).pack(pady=10)

    def remove_task(self, item_id):
        if item_id in self.sod_full_data_storage:
            del self.sod_full_data_storage[item_id]
        self.sod_tree.delete(item_id)

    # -----------------------------
    # PERIOD LOGIC (UPDATED)
    # -----------------------------

    def update_weekly_preview(self):
        """Live preview for Weekly logic: Current/Delayed + rollover offset."""
        try:
            if not hasattr(self, 'weekly_preview_label'):
                return
            today = datetime.date.today()
            logic = self.weekly_logic_var.get() if hasattr(self, 'weekly_logic_var') else self.config.get('WEEKLY_LOGIC', 'Current (ISO)')
            try:
                offset_days = int(self.weekly_start_offset_var.get())
            except Exception:
                offset_days = int(self.config.get('WEEKLY_START_OFFSET_DAYS', 0) or 0)
            if offset_days < -6: offset_days = -6
            if offset_days > 6: offset_days = 6
            iso_monday = today - datetime.timedelta(days=today.weekday())
            rollover_date = iso_monday + datetime.timedelta(days=offset_days)
            active_week_monday = iso_monday if today >= rollover_date else (iso_monday - datetime.timedelta(days=7))
            label_week_monday = active_week_monday - datetime.timedelta(days=7) if logic == 'Delayed (ISO-1)' else active_week_monday
            iso_year, iso_week, _ = label_week_monday.isocalendar()
            period = f"W{iso_week:02d}{yy_from_year(iso_year)}"
            self.weekly_preview_label.config(text=f"Preview (today): {period} | Rollover: {rollover_date.strftime('%a %d/%m/%Y')} | Offset: {offset_days} (0=Mon, 4=Fri)")
        except Exception:
            pass

    def update_daily_preview(self):
        """Live preview for Daily period logic."""
        try:
            if not hasattr(self, 'daily_preview_label'):
                return
            today = datetime.date.today()
            logic = self.daily_logic_var.get() if hasattr(self, 'daily_logic_var') else self.config.get('DAILY_LOGIC', 'Delayed (Day-1)')
            date_to_use = today if logic == 'Current (Day)' else (today - datetime.timedelta(days=1))
            day_of_year = date_to_use.timetuple().tm_yday
            year_short = date_to_use.strftime('%Y')
            period = f"D{day_of_year:03d}{year_short}"
            self.daily_preview_label.config(text=f"Preview (today): {period} | Effective date: {date_to_use.strftime('%a %d/%m/%Y')}")
        except Exception:
            pass

    def update_monthly_preview(self):
        """Live preview for Monthly period logic (cutoff + delayed/current)."""
        try:
            if not hasattr(self, 'monthly_preview_label'):
                return
            today = datetime.date.today()
            m_logic = self.monthly_logic_var.get() if hasattr(self, 'monthly_logic_var') else self.config.get('MONTHLY_LOGIC', 'Delayed (Month-1)')
            try:
                start_day_cfg = int(self.monthly_start_day_var.get())
            except Exception:
                start_day_cfg = int(self.config.get('MONTHLY_START_DAY', 1) or 1)
            start_day = clamp_monthly_start_day(today.year, today.month, start_day_cfg)
            effective_date = today - relativedelta(months=1) if today.day < start_day else today
            if m_logic == 'Delayed (Month-1)':
                effective_date = effective_date - relativedelta(months=1)
            period = f"M{effective_date.month:02d}{effective_date.strftime('%Y')}"
            self.monthly_preview_label.config(text=f"Preview (today): {period} | Effective month: {effective_date.strftime('%b %Y')} | Cutoff day: {start_day}")
        except Exception:
            pass

    def calculate_period(self, frequency):
        today = datetime.date.today()
        # Support multi-frequency input (e.g., 'Weekly, Monthly')
        freq_list, _freq_display = normalize_frequency_string(frequency)
        if not freq_list:
            return ""

        def _period_for(single_freq: str) -> str:
            if single_freq == 'Daily':
                logic = self.config.get('DAILY_LOGIC', 'Delayed (Day-1)')
                date_to_use = today if logic == 'Current (Day)' else (today - datetime.timedelta(days=1))
                day_of_year = date_to_use.timetuple().tm_yday
                year_short = date_to_use.strftime('%Y')
                return f"D{day_of_year:03d}{year_short}"
            elif single_freq == 'Weekly':
                logic = self.config.get('WEEKLY_LOGIC', 'Current (ISO)')
                try:
                    offset_days = int(self.config.get('WEEKLY_START_OFFSET_DAYS', 0))
                except Exception:
                    offset_days = 0
                if offset_days < -6: offset_days = -6
                if offset_days > 6: offset_days = 6
                iso_monday = today - datetime.timedelta(days=today.weekday())
                rollover_date = iso_monday + datetime.timedelta(days=offset_days)
                active_week_monday = iso_monday if today >= rollover_date else (iso_monday - datetime.timedelta(days=7))
                if logic == 'Delayed (ISO-1)':
                    active_week_monday = active_week_monday - datetime.timedelta(days=7)
                iso_year, iso_week, _ = active_week_monday.isocalendar()
                return f"W{iso_week:02d}{yy_from_year(iso_year)}"
            elif single_freq == 'Monthly':
                m_logic = self.config.get('MONTHLY_LOGIC', 'Delayed (Month-1)')
                start_day_cfg = self.config.get('MONTHLY_START_DAY', 1)
                start_day = clamp_monthly_start_day(today.year, today.month, start_day_cfg)
                effective_date = today - relativedelta(months=1) if today.day < start_day else today
                if m_logic == 'Delayed (Month-1)':
                    effective_date = effective_date - relativedelta(months=1)
                return f"M{effective_date.month:02d}{effective_date.strftime('%Y')}"
            return ''

        if len(freq_list) == 1:
            return _period_for(freq_list[0])

        # Multi-frequency: combine period tags in the same order (Weekly then Monthly)
        periods = [p for p in (_period_for(f) for f in freq_list) if p]
        return ' | '.join(periods)

    # -----------------------------
    # SOD ACTIONS
    # -----------------------------
    def add_task(self, task_values=None):
        # Fields: tasklist, stream, frequency, period, start_time, end_time, issue, remarks
        start_time = ''
        end_time = ''
        issue = ''
        remarks = 'Not started'
        stream = ''
        period_override = ''  # preserve period when loading carried-over tasks
        if task_values:
            tasklist = task_values[0] if len(task_values) > 0 else ''
            def _looks_like_freq(s):
                ss = (str(s) if s is not None else '').strip().lower()
                return ss in ('daily','weekly','monthly') or 'weekly' in ss or 'monthly' in ss or 'daily' in ss
            # Preset/task_values stream support:
            # - Old format: [task, freq, ...] -> stream=''
            # - New format: [task, stream, freq, ...] -> stream=task_values[1]
            stream = ''
            try:
                if len(task_values) > 2 and (not _looks_like_freq(task_values[1])) and _looks_like_freq(task_values[2]):
                    stream = str(task_values[1] or '').strip()
            except Exception:
                stream = ''

            # Support old/new preset rows
            if len(task_values) > 2 and (not _looks_like_freq(task_values[1])) and _looks_like_freq(task_values[2]):
                frequency = str(task_values[2])
                base = 3
            else:
                frequency = str(task_values[1]) if len(task_values) > 1 else ''
                base = 2
                # If task_values provides an explicit Period (from prior day/week/month), keep it
                try:
                    if len(task_values) > base:
                        cand = str(task_values[base] or '').strip()
                        cand_norm = re.sub(r'\s+', ' ', cand)
                        if re.match(r'^(?:D\d{3}(?:\d{2}|\d{4})|W\d{2}(?:\d{2}|\d{4})|M\d{2}(?:\d{2}|\d{4}))(?:\s+(?:D\d{3}(?:\d{2}|\d{4})|W\d{2}(?:\d{2}|\d{4})|M\d{2}(?:\d{2}|\d{4})))*$', cand_norm):
                            period_override = cand
                            base += 1
                except Exception:
                    pass
            if len(task_values) > base:
                _r = str(task_values[base] or '').strip()
                if _r:
                    remarks = _r
            if len(task_values) > base+1:
                start_time = str(task_values[base+1] or '').strip()
            if len(task_values) > base+2:
                end_time = str(task_values[base+2] or '').strip()
            if len(task_values) > base+3:
                issue = str(task_values[base+3] or '').strip()
        else:
            tasklist = self.tasklist_entry.get().strip()
            frequency = self.frequency_var.get().strip()
            stream = (getattr(self, 'sod_stream_var', tk.StringVar()).get() or '').strip()
            _fl, _fd = normalize_frequency_string(frequency)
            frequency = _fd or frequency
            sh, sm, sa = self.sod_start_hour_var.get(), self.sod_start_minute_var.get(), self.sod_start_ampm_var.get()
            eh, em, ea = self.sod_end_hour_var.get(), self.sod_end_minute_var.get(), self.sod_end_ampm_var.get()
            if sh and sm and sa:
                start_time = f"{sh}:{sm} {sa}"
            if eh and em and ea:
                end_time = f"{eh}:{em} {ea}"
        if not tasklist:
            messagebox.showwarning('Warning', 'Tasklist field cannot be empty.')
            return
        period = period_override if str(period_override).strip() != "" else self.calculate_period(frequency)
        display_values = (tasklist, stream, frequency, period, start_time, end_time, issue, remarks, '❌')
        full_data_values = (tasklist, stream, frequency, period, start_time, end_time, issue, remarks)
        item_id = self.sod_tree.insert('', tk.END, values=display_values)
        self.sod_full_data_storage[item_id] = full_data_values
        if not task_values:
            self.tasklist_entry.delete(0, tk.END)
            self.frequency_combo.set('')
            self.task_helper_combo.set('')
            self.sod_start_hour_var.set(''); self.sod_start_minute_var.set(''); self.sod_start_ampm_var.set('')
            self.sod_end_hour_var.set(''); self.sod_end_minute_var.set(''); self.sod_end_ampm_var.set('')
            try:
                self.sod_stream_var.set('')
            except Exception:
                pass
    def load_preset_tasks(self):
        today = datetime.date.today()
        today_name = today.strftime('%A')
        today_day_num = today.day

        tasks_to_load = []
        tasks_to_load.extend(self.presets.get("Daily", []))
        tasks_to_load.extend(self.presets.get("Weekdays", {}).get(today_name, []))
        tasks_to_load.extend(self.presets.get("Monthly", {}).get(str(today_day_num), []))

        # --- Weekend Monthly Task Shifting (Friday only) ---
        # If a monthly task falls on Saturday/Sunday, make it available on Friday (pre-load).
        # Note: We intentionally do NOT do a general Monday catch-up for weekend tasks.
        # Month rollover (e.g., next month day "1") is handled separately on the first Monday.

        if today.weekday() == 4:  # Friday
            sat = today + datetime.timedelta(days=1)
            sun = today + datetime.timedelta(days=2)
            for d in (sat, sun):
                # Only pre-load weekend monthly tasks if the weekend is still in the same month.
                if d.month == today.month:
                    tasks_to_load.extend(self.presets.get("Monthly", {}).get(str(d.day), []))

        # --- Month Start Weekend Rule (Day "1" -> first Monday) ---
        # If the 1st day of the current month falls on a weekend, load Monthly["1"] tasks
        # on the first Monday of the month.
        if today.weekday() == 0:  # Monday
            first_day = datetime.date(today.year, today.month, 1)
            if first_day.weekday() in (5, 6):  # Saturday=5, Sunday=6
                first_monday = first_day + datetime.timedelta(days=(7 - first_day.weekday()) % 7)
                if today == first_monday:
                    tasks_to_load.extend(self.presets.get("Monthly", {}).get("1", []))

        _, num_days_in_month = calendar.monthrange(today.year, today.month)
        if today_day_num == num_days_in_month:
            for check_day in range(today_day_num + 1, 32):
                tasks_for_missing_day = self.presets.get("Monthly", {}).get(str(check_day), [])
                if tasks_for_missing_day:
                    tasks_to_load.extend(tasks_for_missing_day)

        if tasks_to_load:
            if messagebox.askyesno("Confirm", f"Found {len(tasks_to_load)} preset tasks for today.\n\nLoad them into your list?"):
                for task_values in tasks_to_load:
                    self.add_task(task_values=task_values)
        else:
            messagebox.showinfo("Info", "No presets found for today.")

    def prepare_sod(self):
        if not self.sod_tree.get_children():
            messagebox.showerror("Error", "Task list is empty.")
            return

        tasks_data = []
        offsite_schedule = datetime.date.today().strftime('%d/%m/%Y')
        for item_id in self.sod_tree.get_children():
            tasklist, stream, frequency, period, start_time, end_time, issue, remarks = self.sod_full_data_storage[item_id]
            full_row = (
                self.config['FIXED_COUNTRY'],
                self.config['YOUR_NAME'],
                stream,
                tasklist,
                offsite_schedule,
                frequency,
                period,
                start_time,
                end_time,
                issue,
                remarks
            )
            tasks_data.append(full_row)

        sod_created_time_str = now_display_time()

        # Use Actual Start Shift if set; otherwise fall back to scheduled START_TIME
        ah = getattr(self, 'sod_shift_hour_var', tk.StringVar()).get()
        am = getattr(self, 'sod_shift_minute_var', tk.StringVar()).get()
        ap = getattr(self, 'sod_shift_ampm_var', tk.StringVar()).get()
        actual_start_shift = (f"{ah}:{am}{ap}" if (ah and am and ap) else (self.config.get('START_TIME') or ''))

        sod_file = os.path.join(RESOURCES_DIR, f"sod_tasks_{datetime.date.today().strftime('%Y-%m-%d')}.json")
        payload = {
            "meta": {
                "sod_created_time": sod_created_time_str,
                "sod_created_iso": datetime.datetime.now().isoformat(timespec="seconds"),
                "actual_start_shift": actual_start_shift
            },
            "tasks": tasks_data
        }
        with open(sod_file, 'w') as f:
            json.dump(payload, f, indent=4)

        subject = f"WFH SOD Notification | {self.config['YOUR_NAME']} | {datetime.date.today().strftime('%d/%m/%Y')}"
        body = create_sod_html_body(tasks_data, self.config, sod_created_time_str, actual_start_shift=actual_start_shift)

        if generate_email_draft(subject, body, self.config):
            messagebox.showinfo("Success", "SOD Email draft created!")

    def load_unfinished_tasks(self):
        try:
            today = datetime.date.today()
            # Option A: Load unfinished tasks from the most recent available EOD report (not just yesterday).
            eod_files = []
            for filename in os.listdir(RESOURCES_DIR):
                if filename.startswith("eod_report_") and filename.endswith(".json"):
                    date_str = filename.replace("eod_report_", "").replace(".json", "")
                    try:
                        file_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
                        eod_files.append((file_date, os.path.join(RESOURCES_DIR, filename)))
                    except Exception:
                        continue
            candidates = [p for p in eod_files if p[0] < today]
            chosen = max(candidates, default=None, key=lambda x: x[0])
            if chosen is None:
                chosen = max(eod_files, default=None, key=lambda x: x[0])
            if chosen is None:
                messagebox.showinfo("No Data Found", "No EOD report file found yet.")
                return
            eod_file_to_load = chosen[1]
            with open(eod_file_to_load, "r") as f:
                last_eod_data = json.load(f)
            unfinished_tasks = []
            for task_row in (last_eod_data or []):
                if not isinstance(task_row, (list, tuple)):
                    continue
                detected_status = None
                for item in reversed(task_row):
                    if item in ("🔄 In Progress", "➡️ Carried Over"):
                        detected_status = item
                        break
                if detected_status == "🔄 In Progress":
                    tasklist = task_row[3] if len(task_row) > 3 else (task_row[0] if task_row else "")
                    freq = task_row[5] if len(task_row) > 5 else ""
                    period = task_row[6] if len(task_row) > 6 else ''
                    unfinished_tasks.append((tasklist, freq, period, "In Progress"))
                elif detected_status == "➡️ Carried Over":
                    tasklist = task_row[3] if len(task_row) > 3 else (task_row[0] if task_row else "")
                    freq = task_row[5] if len(task_row) > 5 else ""
                    period = task_row[6] if len(task_row) > 6 else ''
                    unfinished_tasks.append((tasklist, freq, period, "Carried Over"))
            if not unfinished_tasks:
                messagebox.showinfo("All Clear!", "No unfinished tasks found from your last report.")
                return
            if messagebox.askyesno("Confirm", f"Found {len(unfinished_tasks)} unfinished task(s).\n\nLoad them into today's SOD list?"):
                for task_vals in unfinished_tasks:
                    self.add_task(task_values=task_vals)
                messagebox.showinfo("Success", f"{len(unfinished_tasks)} task(s) loaded.")
        except Exception as e:
            messagebox.showerror("Error Loading Tasks", f"An error occurred while loading unfinished tasks.\n\nDetails: {e}")

    # -----------------------------
    # EOD UI + ACTIONS
    # -----------------------------
    def create_eod_widgets(self):
        eod_main_frame = ttk.LabelFrame(self.eod_frame, text="EOD Task Status")
        eod_main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.eod_display_cols = ('task', 'done', 'progress', 'carried')
        self.eod_tree = ttk.Treeview(eod_main_frame, columns=self.eod_display_cols, show='headings')
        self.eod_tree.heading('task', text='Tasklist')
        self.eod_tree.column('task', width=400)
        self.eod_tree.heading('done', text='Done')
        self.eod_tree.column('done', width=80, anchor='center')
        self.eod_tree.heading('progress', text='In Progress')
        self.eod_tree.column('progress', width=80, anchor='center')
        self.eod_tree.heading('carried', text='Carried Over')
        self.eod_tree.column('carried', width=80, anchor='center')
        self.eod_tree.pack(fill="both", expand=True, padx=5, pady=5)
        self.eod_tree.bind("<ButtonRelease-1>", self.handle_eod_tree_click)

        ttk.Button(eod_main_frame, text="Load Tasks", command=self.load_sod_tasks_to_eod).pack(pady=5)

        self.screenshot_frame = ttk.Frame(self.eod_frame)
        self.screenshot_frame.pack(fill="x", padx=10, pady=5)
        ttk.Button(self.screenshot_frame, text="Paste from Clipboard", command=self.paste_from_clipboard).pack(side="left", padx=(0, 5))
        ttk.Button(self.screenshot_frame, text="Browse...", command=self.browse_for_screenshot, width=10).pack(side="left", padx=(0, 5))

        self.screenshot_label = ttk.Label(self.screenshot_frame, text="No screenshot selected.", style="TLabel", relief="sunken", anchor="w")
        self.screenshot_label.pack(side="left", fill="x", expand=True, ipady=2)

        time_frame = ttk.Frame(self.eod_frame)
        time_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(time_frame, text="Actual End Time:").pack(side="left", padx=5)
        self.eod_hour_var, self.eod_minute_var, self.eod_ampm_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        hours = [str(i).zfill(2) for i in range(1, 13)]
        minutes = [str(i).zfill(2) for i in range(60)]
        ttk.Combobox(time_frame, textvariable=self.eod_hour_var, values=hours, width=3).pack(side="left")
        ttk.Combobox(time_frame, textvariable=self.eod_minute_var, values=minutes, width=3).pack(side="left")
        ttk.Combobox(time_frame, textvariable=self.eod_ampm_var, values=['AM', 'PM'], width=3).pack(side="left", padx=5)

        # Default Actual End Time to scheduled END_TIME
        self.set_default_end_time()

        self.btn_prepare_eod = ttk.Button(self.eod_frame, text="Prepare EOD Email Draft", command=self.prepare_eod)
        self._btn_prepare_eod_pack = {'pady': 10}
        self.btn_prepare_eod.pack(**self._btn_prepare_eod_pack)

        # --- Copy bar (New Outlook mode) ---
        self.eod_copy_bar = ttk.LabelFrame(self.eod_frame, text='Copy (New Outlook)')
        self._eod_copy_bar_pack = {'side': 'bottom', 'fill': 'x', 'padx': 10, 'pady': (0, 10)}
        self.eod_copy_bar.pack(**self._eod_copy_bar_pack)
        _bar = ttk.Frame(self.eod_copy_bar)
        _bar.pack(fill='x', padx=8, pady=6)
        self._populate_copy_bar(_bar, self.copy_eod_subject, self.copy_eod_body)


    def browse_for_screenshot(self):
        file_path = filedialog.askopenfilename(
            title="Select Screenshot File",
            filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp"), ("All Files", "*.*")]
        )
        if file_path:
            self.eod_screenshot_path.set(file_path)
            self.screenshot_label.config(text=f" {os.path.basename(file_path)}")

    def paste_from_clipboard(self):
        try:
            image = ImageGrab.grabclipboard()
            if image is None:
                messagebox.showwarning("Clipboard Error", "No image found on the clipboard.")
                return
            # Limit pasted screenshot width to 1000px while keeping aspect ratio
            try:
                max_w = 1000
                w, h = image.size
                if w and w > max_w:
                    new_h = int(h * (max_w / float(w)))
                    try:
                        resample = Image.Resampling.LANCZOS
                    except AttributeError:
                        resample = Image.LANCZOS
                    image = image.resize((max_w, new_h), resample=resample)
            except Exception:
                pass

            image.save(TEMP_SCREENSHOT_PATH, "PNG")
            self.eod_screenshot_path.set(TEMP_SCREENSHOT_PATH)
            self.screenshot_label.config(text=" Image pasted from clipboard.")
        except Exception as e:
            messagebox.showerror("Clipboard Error", f"Could not get image from clipboard.\n\nDetails: {e}")

    def handle_eod_tree_click(self, event):
        if self.eod_tree.identify_region(event.x, event.y) != "cell":
            return
        col_name = self._tree_colname_from_event(self.eod_tree, event)
        selected_item_id = self.eod_tree.focus()
        if not selected_item_id:
            return
        status_map = {'done': "✅ Done", 'progress': "🔄 In Progress", 'carried': "➡️ Carried Over"}
        if col_name in status_map:
            self.set_status(selected_item_id, status_map[col_name])

    def set_default_start_shift(self):
        """Default Actual Start Shift dropdown to scheduled START_TIME."""
        try:
            t = datetime.datetime.strptime(self.config.get('START_TIME'), '%I:%M%p')
            self.sod_shift_hour_var.set(t.strftime('%I'))
            self.sod_shift_minute_var.set(t.strftime('%M'))
            self.sod_shift_ampm_var.set(t.strftime('%p'))
        except Exception:
            pass


    def set_default_end_time(self):
        try:
            time_obj = datetime.datetime.strptime(self.config.get('END_TIME'), "%I:%M%p")
            self.eod_hour_var.set(time_obj.strftime("%I"))
            self.eod_minute_var.set(time_obj.strftime("%M"))
            self.eod_ampm_var.set(time_obj.strftime("%p"))
        except Exception:
            pass

    def load_sod_tasks_to_eod(self):
        self.eod_screenshot_path.set("")
        self.screenshot_label.config(text="No screenshot selected.")
        self.eod_tree.delete(*self.eod_tree.get_children())
        self.eod_full_data.clear()

        today_date = datetime.date.today()
        yesterday_date = today_date - datetime.timedelta(days=1)

        today_file = os.path.join(RESOURCES_DIR, f"sod_tasks_{today_date.strftime('%Y-%m-%d')}.json")
        yesterday_file = os.path.join(RESOURCES_DIR, f"sod_tasks_{yesterday_date.strftime('%Y-%m-%d')}.json")

        sod_file_to_load = today_file if os.path.exists(today_file) else (yesterday_file if os.path.exists(yesterday_file) else None)
        date_to_use = today_date if os.path.exists(today_file) else (yesterday_date if os.path.exists(yesterday_file) else today_date)

        if sod_file_to_load:
            self.current_shift_date = date_to_use
            with open(sod_file_to_load, 'r') as f:
                content = json.load(f)

            if isinstance(content, dict):
                tasks_from_file = content.get("tasks", [])
                meta = content.get("meta", {})
                self.loaded_sod_created_time = meta.get("sod_created_time", "N/A")
                self.loaded_actual_start_shift = meta.get("actual_start_shift")
            else:
                tasks_from_file = content
                self.loaded_sod_created_time = "N/A"
                self.loaded_actual_start_shift = None

            for i, full_data in enumerate(tasks_from_file):
                item_id = f"EOD{i:03}"
                self.eod_full_data[item_id] = list(full_data)
                self.eod_tree.insert('', tk.END, iid=item_id, values=(full_data[3], '🔘', '🔘', '🔘'))

            self.set_default_end_time()
        else:
            self.current_shift_date = date_to_use
            self.loaded_sod_created_time = "N/A"
            messagebox.showinfo("Info", "No SOD data found for today or yesterday.")

    def set_status(self, item_id, status):
        values = list(self.eod_tree.item(item_id, 'values'))
        task_description = values[0]
        status_map = {"✅ Done": 1, "🔄 In Progress": 2, "➡️ Carried Over": 3}
        new_values = [task_description, '🔘', '🔘', '🔘']
        if status in status_map:
            new_values[status_map[status]] = '✅'
        self.eod_tree.item(item_id, values=tuple(new_values))

        if item_id in self.eod_full_data:
            if len(self.eod_full_data[item_id]) == 8:
                self.eod_full_data[item_id].append(status)
            else:
                self.eod_full_data[item_id][-1] = status

    def _get_latest_sod_meta(self):
        """Return (meta_dict, date_used) from the most recent SOD snapshot (today else yesterday).
        Used by EOD when the user did not click 'Load Tasks'.
        Returns ({}, None) if no snapshot is found or format is unexpected.
        """
        try:
            today = datetime.date.today()
            yesterday = today - datetime.timedelta(days=1)
            today_file = os.path.join(RESOURCES_DIR, f"sod_tasks_{today.strftime('%Y-%m-%d')}.json")
            yesterday_file = os.path.join(RESOURCES_DIR, f"sod_tasks_{yesterday.strftime('%Y-%m-%d')}.json")
            if os.path.exists(today_file):
                sod_file = today_file
                date_used = today
            elif os.path.exists(yesterday_file):
                sod_file = yesterday_file
                date_used = yesterday
            else:
                return {}, None
            with open(sod_file, 'r') as f:
                content = json.load(f)
            if isinstance(content, dict):
                return (content.get('meta', {}) or {}), date_used
            return {}, date_used
        except Exception:
            return {}, None

    def prepare_eod(self):
        hour, minute, ampm = self.eod_hour_var.get(), self.eod_minute_var.get(), self.eod_ampm_var.get()
        actual_end_time = f"{hour}:{minute}{ampm}" if (hour and minute and ampm) else (self.config.get('END_TIME') or None)

        eod_data = list(self.eod_full_data.values())
        if not self.current_shift_date:
            # If no SOD was loaded, infer the correct shift date for overnight EOD creation.
            # Rule: If scheduled START_TIME is PM and EOD is generated after midnight (next day),
            # treat it as yesterday's shift.
            now_dt = datetime.datetime.now()
            start_time_str = str(self.config.get('START_TIME', '') or '').strip()
            try:
                start_clock = datetime.datetime.strptime(start_time_str, '%I:%M%p').time() if start_time_str else None
            except Exception:
                start_clock = None
            starts_pm = start_time_str.upper().endswith('PM')
            # If current time is earlier than the scheduled PM start time, we are in the next-day window (12AM onwards).
            if starts_pm and start_clock and now_dt.time() < start_clock:
                self.current_shift_date = now_dt.date() - datetime.timedelta(days=1)
            else:
                self.current_shift_date = now_dt.date()

        screenshot_file_path = self.eod_screenshot_path.get()
        absolute_screenshot_path = os.path.abspath(screenshot_file_path) if (screenshot_file_path and os.path.exists(screenshot_file_path)) else None

        has_valid_tasks = bool(eod_data) and all(len(row) >= 9 for row in eod_data)
        tasks_to_render = eod_data if has_valid_tasks else None

        eod_created_time_str = now_display_time()
        sod_created_time_str = self.loaded_sod_created_time or "N/A"

        # Auto-load SOD meta (Actual Start Shift) when user did not click 'Load Tasks'
        if not getattr(self, 'loaded_actual_start_shift', None):
            meta_fallback, date_used = self._get_latest_sod_meta()
            if meta_fallback:
                self.loaded_actual_start_shift = meta_fallback.get('actual_start_shift')
                if (self.loaded_sod_created_time in (None, '', 'N/A')) and meta_fallback.get('sod_created_time'):
                    self.loaded_sod_created_time = meta_fallback.get('sod_created_time')
                if (not self.current_shift_date) and date_used:
                    self.current_shift_date = date_used

        actual_start_shift = self.loaded_actual_start_shift or (self.config.get('START_TIME') or None)

        body = create_eod_html_body(
            tasks_data=tasks_to_render,
            config=self.config,
            actual_end_time=actual_end_time,
            shift_date=self.current_shift_date,
            include_screenshot=bool(absolute_screenshot_path),
            sod_created_time_str=sod_created_time_str,
            eod_created_time_str=eod_created_time_str,
            actual_start_shift=actual_start_shift
        )

        subject_date_str = self.current_shift_date.strftime('%d/%m/%Y')
        subject = f"WFH EOD Notification | {self.config.get('YOUR_NAME','')} | {self.current_shift_date.strftime('%d/%m/%Y')}"

        if generate_email_draft(subject, body, self.config, screenshot_path=absolute_screenshot_path):
            success_msg = "EOD Email draft created successfully."
            if has_valid_tasks:
                try:
                    eod_file_path = os.path.join(RESOURCES_DIR, f"eod_report_{self.current_shift_date.strftime('%Y-%m-%d')}.json")
                    with open(eod_file_path, 'w') as f:
                        json.dump(eod_data, f, indent=4)
                    success_msg += "\nTask history saved for future carry-over."
                except Exception as e:
                    messagebox.showwarning("History Save Error", f"Could not save history file: {e}")
            else:
                success_msg += "\n(Note: Task history was NOT saved because the task list was incomplete or missing.)"

            self.eod_screenshot_path.set("")
            self.screenshot_label.config(text="No screenshot selected.")
            messagebox.showinfo("Success", success_msg)

    # -----------------------------

    # -----------------------------
    # OT IN / OT OUT UI + ACTIONS
    # -----------------------------
    def create_ot_in_widgets(self):
        frame = ttk.LabelFrame(self.ot_in_frame, text="OT In")
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        just_frame = ttk.Frame(frame)
        just_frame.pack(fill="x", pady=(0, 8))
        ttk.Label(just_frame, text="Justification:").pack(side="left")
        self.ot_justification_var = tk.StringVar()
        ttk.Entry(just_frame, textvariable=self.ot_justification_var, width=80).pack(side="left", padx=(8,0), fill="x", expand=True)
        # OT Date picker (calendar)
        date_frame = ttk.Frame(frame)
        date_frame.pack(fill="x", pady=(0, 8))
        ttk.Label(date_frame, text="OT Date:").pack(side="left")
        self.ot_date_var = tk.StringVar()
        # default to today's date (current month/year)
        self.ot_selected_date = datetime.date.today()
        self.ot_date_var.set(self.ot_selected_date.strftime('%d/%m/%Y'))
        ttk.Entry(date_frame, textvariable=self.ot_date_var, width=12, state='readonly').pack(side="left", padx=(8,0))
        ttk.Button(date_frame, text='📅', width=3, command=self._open_ot_date_picker).pack(side="left", padx=(5,0))
        ttk.Label(date_frame, text="(select OT work date)").pack(side="left", padx=(8,0))


        time_frame = ttk.Frame(frame)
        time_frame.pack(fill="x", pady=(0, 8))
        ttk.Label(time_frame, text="OT From:").pack(side="left")
        self.ot_from_h = tk.StringVar(); self.ot_from_m = tk.StringVar(); self.ot_from_a = tk.StringVar()
        self.ot_to_h = tk.StringVar(); self.ot_to_m = tk.StringVar(); self.ot_to_a = tk.StringVar()
        hours = [str(i).zfill(2) for i in range(1, 13)]
        minutes = [str(i).zfill(2) for i in range(60)]
        ttk.Combobox(time_frame, textvariable=self.ot_from_h, values=hours, width=3, state="readonly").pack(side="left", padx=(5,0))
        ttk.Combobox(time_frame, textvariable=self.ot_from_m, values=minutes, width=3, state="readonly").pack(side="left")
        ttk.Combobox(time_frame, textvariable=self.ot_from_a, values=['AM','PM'], width=3, state="readonly").pack(side="left", padx=(0,10))
        ttk.Label(time_frame, text="OT To:").pack(side="left")
        ttk.Combobox(time_frame, textvariable=self.ot_to_h, values=hours, width=3, state="readonly").pack(side="left", padx=(5,0))
        ttk.Combobox(time_frame, textvariable=self.ot_to_m, values=minutes, width=3, state="readonly").pack(side="left")
        ttk.Combobox(time_frame, textvariable=self.ot_to_a, values=['AM','PM'], width=3, state="readonly").pack(side="left")
        ttk.Label(time_frame, text="(leave blank if tbd)").pack(side="left", padx=(8,0))

        try:
            t = datetime.datetime.strptime(self.config.get('END_TIME'), "%I:%M%p")
            self.ot_from_h.set(t.strftime("%I")); self.ot_from_m.set(t.strftime("%M")); self.ot_from_a.set(t.strftime("%p"))
        except Exception:
            pass

        task_input = ttk.LabelFrame(frame, text="Add OT Task")
        task_input.pack(fill="x", pady=(0, 8))
        task_input.columnconfigure(1, weight=1)
        ttk.Label(task_input, text="Tasklist:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.ot_task_entry = ttk.Entry(task_input)
        self.ot_task_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
        ttk.Button(task_input, text="Add Task", command=self.add_ot_task).grid(row=3, column=1, sticky='w', padx=5, pady=8)

        cols = ("Tasklist","Frequency","Period","Status","Action")
        self.ot_in_tree = ttk.Treeview(frame, columns=cols, show='headings', height=10)
        for c in cols:
            self.ot_in_tree.heading(c, text=c)
        self.ot_in_tree.column('Tasklist', width=420, anchor='w')
        self.ot_in_tree.column('Frequency', width=120, anchor='w')
        self.ot_in_tree.column('Period', width=120, anchor='w')
        self.ot_in_tree.column('Status', width=260, anchor='w')
        self.ot_in_tree.column('Action', width=80, anchor='center', stretch=False)
        self.ot_in_tree.pack(fill='both', expand=True, padx=5, pady=5)
        self.ot_in_tree.bind('<ButtonRelease-1>', self.handle_ot_in_click)

        self.btn_prepare_ot_in = ttk.Button(frame, text="Prepare OT In Email Draft", command=self.prepare_ot_in)
        self._btn_prepare_ot_in_pack = {'pady': 8}
        self.btn_prepare_ot_in.pack(**self._btn_prepare_ot_in_pack)

        # --- Copy bar (New Outlook mode) ---
        self.ot_in_copy_bar = ttk.LabelFrame(frame, text='Copy (New Outlook)')
        self._ot_in_copy_bar_pack = {'side': 'bottom', 'fill': 'x', 'padx': 5, 'pady': (0, 8)}
        self.ot_in_copy_bar.pack(**self._ot_in_copy_bar_pack)
        _bar = ttk.Frame(self.ot_in_copy_bar)
        _bar.pack(fill='x', padx=8, pady=6)
        self._populate_copy_bar(_bar, self.copy_ot_in_subject, self.copy_ot_in_body)

        self.ot_in_full_data = {}

    def handle_ot_in_click(self, event):
        if self.ot_in_tree.identify_region(event.x, event.y) != 'cell':
            return
        col_name = self._tree_colname_from_event(self.ot_in_tree, event)
        item = self.ot_in_tree.focus()
        if not item:
            return
        if col_name == 'Action':
            self.ot_in_full_data.pop(item, None)
            self.ot_in_tree.delete(item)
            return
        if col_name == 'Status':
            cur = self.ot_in_tree.item(item, 'values')[3]
            state = _ot_status_state_from_any(cur)
            nxt_state = 'done' if state != 'done' else 'in_progress'
            nxt = ot_status_dual(nxt_state)
            vals = list(self.ot_in_tree.item(item, 'values'))
            vals[3] = nxt
            self.ot_in_tree.item(item, values=tuple(vals))
            if item in self.ot_in_full_data:
                self.ot_in_full_data[item]['status'] = nxt


    def add_ot_task(self):
        task = self.ot_task_entry.get().strip()
        if not task:
            messagebox.showwarning('Warning','Tasklist field cannot be empty.')
            return
        # Frequency/Period are intentionally blank for OT
        freq = ''
        period = ''
        status = ot_status_dual('in_progress')
        iid = self.ot_in_tree.insert('', tk.END, values=(task, freq, period, status, '❌'))
        self.ot_in_full_data[iid] = {'task':task,'freq':freq,'period':period,'status':status,'issue':'','remarks':''}
        self.ot_task_entry.delete(0, tk.END)

    def _ot_time_str(self, h_var, m_var, a_var):
        h=h_var.get().strip(); m=m_var.get().strip(); a=a_var.get().strip()
        return f"{h}:{m}{a}" if (h and m and a) else ''

    def prepare_ot_in(self):
        now_dt = datetime.datetime.now()
        shift_date = getattr(self, 'ot_selected_date', datetime.date.today())
        ot_from = self._ot_time_str(self.ot_from_h, self.ot_from_m, self.ot_from_a)
        if not ot_from:
            messagebox.showerror('Error','OT From time is required.')
            return
        ot_to = self._ot_time_str(self.ot_to_h, self.ot_to_m, self.ot_to_a)
        total = calculate_total_hours(ot_from, ot_to) if ot_to else ''
        offsite = shift_date.strftime('%d/%m/%Y')
        tasks=[]
        for iid in self.ot_in_tree.get_children():
            data=self.ot_in_full_data.get(iid,{})
            status_disp = data.get('status','')
            status_single = ot_status_single(_ot_status_state_from_any(status_disp))
            tasks.append((self.config.get('FIXED_COUNTRY',''), self.config.get('YOUR_NAME',''), '', data.get('task',''), offsite, '', '', '', '', status_single, data.get('issue',''), data.get('remarks','')))
        justification = self.ot_justification_var.get()
        subject = f"OT Notification | {self.config.get('YOUR_NAME','')} | {self.config.get('FIXED_STREAM','')} | ({shift_date.strftime('%d.%m.%Y')})"
        body = create_ot_in_html_body(tasks, self.config, shift_date, ot_from, ot_to, total, justification)
        payload={'meta':{'shift_date':shift_date.isoformat(),'ot_from':ot_from,'ot_to':ot_to,'justification':justification,'created_time':now_display_time(),'created_iso':now_dt.isoformat(timespec='seconds')},'tasks':tasks}
        try:
            with open(f"{OT_IN_FILE_PREFIX}{shift_date.strftime('%Y-%m-%d')}.json",'w') as f:
                json.dump(payload,f,indent=4)
        except Exception:
            pass
        if generate_email_draft(subject, body, self.config):
            messagebox.showinfo('Success','OT In email draft created!')

    def create_ot_out_widgets(self):
        frame = ttk.LabelFrame(self.ot_out_frame, text="OT Out")
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        ttk.Button(frame, text='Load OT In Tasks', command=self.load_ot_in_tasks_to_out).pack(pady=(0,8))
        time_frame = ttk.Frame(frame)
        time_frame.pack(fill='x', pady=(0,8))
        ttk.Label(time_frame, text='Actual End Time:').pack(side='left')
        self.ot_out_h = tk.StringVar(); self.ot_out_m = tk.StringVar(); self.ot_out_a = tk.StringVar()
        hours = [str(i).zfill(2) for i in range(1, 13)]
        minutes = [str(i).zfill(2) for i in range(60)]
        ttk.Combobox(time_frame, textvariable=self.ot_out_h, values=hours, width=3, state='readonly').pack(side='left', padx=(5,0))
        ttk.Combobox(time_frame, textvariable=self.ot_out_m, values=minutes, width=3, state='readonly').pack(side='left')
        ttk.Combobox(time_frame, textvariable=self.ot_out_a, values=['AM','PM'], width=3, state='readonly').pack(side='left', padx=(0,10))
        cols=("Tasklist","Status","Action")
        self.ot_out_tree = ttk.Treeview(frame, columns=cols, show='headings', height=12)
        for c in cols:
            self.ot_out_tree.heading(c, text=c)
        self.ot_out_tree.column('Tasklist', width=420, anchor='w')
        self.ot_out_tree.column('Status', width=260, anchor='w')
        self.ot_out_tree.column('Action', width=80, anchor='center', stretch=False)
        self.ot_out_tree.pack(fill='both', expand=True, padx=5, pady=5)
        self.ot_out_tree.bind('<ButtonRelease-1>', self.handle_ot_out_click)
        self.btn_prepare_ot_out = ttk.Button(frame, text='Prepare OT Out Email Draft', command=self.prepare_ot_out)
        self._btn_prepare_ot_out_pack = {'pady': 8}
        self.btn_prepare_ot_out.pack(**self._btn_prepare_ot_out_pack)

        # --- Copy bar (New Outlook mode) ---
        self.ot_out_copy_bar = ttk.LabelFrame(frame, text='Copy (New Outlook)')
        self._ot_out_copy_bar_pack = {'side': 'bottom', 'fill': 'x', 'padx': 5, 'pady': (0, 8)}
        self.ot_out_copy_bar.pack(**self._ot_out_copy_bar_pack)
        _bar = ttk.Frame(self.ot_out_copy_bar)
        _bar.pack(fill='x', padx=8, pady=6)
        self._populate_copy_bar(_bar, self.copy_ot_out_subject, self.copy_ot_out_body)

        self.ot_out_full_data = {}
        self.loaded_ot_in_meta = {}
        self.loaded_ot_in_tasks = []
        self.loaded_ot_shift_date = None

    def handle_ot_out_click(self, event):
        if self.ot_out_tree.identify_region(event.x, event.y) != 'cell':
            return
        col_name = self._tree_colname_from_event(self.ot_out_tree, event)
        item = self.ot_out_tree.focus()
        if not item:
            return
        if col_name == 'Action':
            self.ot_out_full_data.pop(item, None)
            self.ot_out_tree.delete(item)
            return
        if col_name == 'Status':
            cur = self.ot_out_tree.item(item, 'values')[1]
            state = _ot_status_state_from_any(cur)
            nxt_state = 'done' if state != 'done' else 'in_progress'
            nxt = ot_status_dual(nxt_state)
            vals = list(self.ot_out_tree.item(item, 'values'))
            vals[1] = nxt
            self.ot_out_tree.item(item, values=tuple(vals))
            if item in self.ot_out_full_data:
                self.ot_out_full_data[item]['status'] = nxt


    def load_ot_in_tasks_to_out(self):
        candidates=[]
        for fn in os.listdir(RESOURCES_DIR):
            if fn.startswith('ot_in_') and fn.endswith('.json'):
                try:
                    d=datetime.datetime.strptime(fn.replace('ot_in_','').replace('.json',''), "%Y-%m-%d").date()
                    candidates.append((d, os.path.join(RESOURCES_DIR, fn)))
                except Exception:
                    continue
        if not candidates:
            messagebox.showinfo('Info','No OT In file found.')
            return
        candidates.sort(key=lambda x:x[0])
        d, fpath = candidates[-1]
        with open(fpath,'r') as f:
            payload=json.load(f)
        meta=payload.get('meta',{}) if isinstance(payload,dict) else {}
        tasks=payload.get('tasks',[]) if isinstance(payload,dict) else []
        self.loaded_ot_in_meta = meta
        self.loaded_ot_in_tasks = tasks
        try:
            self.loaded_ot_shift_date = datetime.date.fromisoformat(meta.get('shift_date'))
        except Exception:
            self.loaded_ot_shift_date = d
        ot_to = meta.get('ot_to') or ''
        if ot_to:
            try:
                t=datetime.datetime.strptime(ot_to, "%I:%M%p")
                self.ot_out_h.set(t.strftime("%I")); self.ot_out_m.set(t.strftime("%M")); self.ot_out_a.set(t.strftime("%p"))
            except Exception:
                pass
        self.ot_out_tree.delete(*self.ot_out_tree.get_children())
        self.ot_out_full_data.clear()
        for i,row in enumerate(tasks):
            task = row[3] if len(row)>3 else ''
            status_raw = row[9] if len(row)>9 else ''
            status = ot_status_dual(_ot_status_state_from_any(status_raw))
            iid=f'OTOUT{i:03}'
            self.ot_out_tree.insert('', tk.END, iid=iid, values=(task, status, '❌'))
            self.ot_out_full_data[iid]={'task_row': list(row), 'status': status}
        messagebox.showinfo('Loaded', f'Loaded {len(tasks)} OT task(s) from {d.isoformat()}')

    def prepare_ot_out(self):
        if not self.loaded_ot_in_tasks:
            messagebox.showerror('Error','No OT In tasks loaded. Click Load OT In Tasks first.')
            return
        ot_to = self._ot_time_str(self.ot_out_h, self.ot_out_m, self.ot_out_a)
        if not ot_to:
            messagebox.showerror('Error','Actual End Time is required for OT Out.')
            return
        ot_from = self.loaded_ot_in_meta.get('ot_from') or ''
        if not ot_from:
            messagebox.showerror('Error','OT From time not found from OT In file.')
            return
        shift_date = self.loaded_ot_shift_date or infer_shift_date_from_config(datetime.datetime.now(), self.config)
        total = calculate_total_hours(ot_from, ot_to)
        justification = self.loaded_ot_in_meta.get('justification','')
        updated_tasks=[]
        for iid in self.ot_out_tree.get_children():
            base = self.ot_out_full_data.get(iid, {}).get('task_row', [])
            status_disp = self.ot_out_tree.item(iid,'values')[1]
            status_single = ot_status_single(_ot_status_state_from_any(status_disp))
            if isinstance(base, list) and len(base) >= 10:
                base[9] = status_single
            updated_tasks.append(tuple(base))
        subject = f"OT Notification | {self.config.get('YOUR_NAME','')} | {self.config.get('FIXED_STREAM','')} | ({shift_date.strftime('%d.%m.%Y')})"
        body = create_ot_out_html_body(updated_tasks, self.config, shift_date, ot_from, ot_to, total, justification)
        payload={'meta':{'shift_date':shift_date.isoformat(),'ot_from':ot_from,'ot_to':ot_to,'justification':justification,'created_time':now_display_time(),'created_iso':datetime.datetime.now().isoformat(timespec='seconds')},'tasks':updated_tasks}
        try:
            with open(f"{OT_OUT_FILE_PREFIX}{shift_date.strftime('%Y-%m-%d')}.json",'w') as f:
                json.dump(payload,f,indent=4)
        except Exception:
            pass
        if generate_email_draft(subject, body, self.config):
            messagebox.showinfo('Success','OT Out email draft created!')
    # PRESETS UI (same behavior)
    # -----------------------------
    def create_presets_widgets(self):
        self.presets_frame = ttk.Frame(self)

        selection_frame = ttk.Frame(self.presets_frame)
        selection_frame.pack(pady=10, fill="x", padx=10)

        ttk.Label(selection_frame, text="Preset Type:").grid(row=0, column=0, padx=(0, 5), sticky="w")
        self.preset_type_var = tk.StringVar()
        self.preset_type_combo = ttk.Combobox(
            selection_frame,
            textvariable=self.preset_type_var,
            values=["Daily", "Weekday", "Monthly", "Task Dropdown Options"],
            state="readonly",
            width=20
        )
        self.preset_type_combo.grid(row=0, column=1, padx=(0, 20), sticky="w")
        self.preset_type_combo.bind("<<ComboboxSelected>>", self.on_preset_type_change)

        ttk.Label(selection_frame, text="Select Key:").grid(row=0, column=2, padx=(0, 5), sticky="w")
        self.preset_key_var = tk.StringVar()
        self.preset_key_combo = ttk.Combobox(selection_frame, textvariable=self.preset_key_var, state="disabled", width=15)
        self.preset_key_combo.grid(row=0, column=3, sticky="w")
        self.preset_key_combo.bind("<<ComboboxSelected>>", self.on_preset_key_change)

        self.preset_editor_frame = ttk.LabelFrame(self.presets_frame, text="Preset Editor")

        preset_input_frame = ttk.LabelFrame(self.preset_editor_frame, text="Add New Preset Task")
        preset_input_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        preset_input_frame.columnconfigure(1, weight=1)

        ttk.Label(preset_input_frame, text="Tasklist:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.preset_tasklist_entry = ttk.Entry(preset_input_frame)
        self.preset_tasklist_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        ttk.Label(preset_input_frame, text='Stream:').grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.preset_stream_var = tk.StringVar()
        ttk.Entry(preset_input_frame, textvariable=self.preset_stream_var).grid(row=1, column=1, sticky='ew', padx=5, pady=2)

        self.preset_freq_label = ttk.Label(preset_input_frame, text="Frequency:")
        self.preset_freq_label.grid(row=2, column=0, sticky="w", padx=5, pady=2)

        self.preset_frequency_var = tk.StringVar()
        self.preset_frequency_combo = ttk.Combobox(
            preset_input_frame,
            textvariable=self.preset_frequency_var,
            values=['', 'Daily', 'Weekly', 'Monthly', 'Weekly, Monthly'],
            width=15
        )
        self.preset_frequency_combo.grid(row=2, column=1, sticky="w", padx=5, pady=2)
        # Start/End time pickers for presets
        time_hours = [str(i).zfill(2) for i in range(1, 13)]
        time_minutes = [str(i).zfill(2) for i in range(60)]

        ttk.Label(preset_input_frame, text="Start Time:").grid(row=3, column=0, sticky="w", padx=5, pady=2)
        self.preset_start_hour_var, self.preset_start_minute_var, self.preset_start_ampm_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        pst = ttk.Frame(preset_input_frame)
        pst.grid(row=3, column=1, sticky="w", padx=5, pady=2)
        ttk.Combobox(pst, textvariable=self.preset_start_hour_var, values=time_hours, width=3, state="readonly").pack(side="left")
        ttk.Combobox(pst, textvariable=self.preset_start_minute_var, values=time_minutes, width=3, state="readonly").pack(side="left", padx=(5,0))
        ttk.Combobox(pst, textvariable=self.preset_start_ampm_var, values=['AM','PM'], width=3, state="readonly").pack(side="left", padx=(5,0))

        ttk.Label(preset_input_frame, text="End Time:").grid(row=4, column=0, sticky="w", padx=5, pady=2)
        self.preset_end_hour_var, self.preset_end_minute_var, self.preset_end_ampm_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        pet = ttk.Frame(preset_input_frame)
        pet.grid(row=4, column=1, sticky="w", padx=5, pady=2)
        ttk.Combobox(pet, textvariable=self.preset_end_hour_var, values=time_hours, width=3, state="readonly").pack(side="left")
        ttk.Combobox(pet, textvariable=self.preset_end_minute_var, values=time_minutes, width=3, state="readonly").pack(side="left", padx=(5,0))
        ttk.Combobox(pet, textvariable=self.preset_end_ampm_var, values=['AM','PM'], width=3, state="readonly").pack(side="left", padx=(5,0))

        ttk.Button(preset_input_frame, text="Add/Update", command=self.add_preset_task_to_preset_editor).grid(row=5, column=1, pady=10, sticky="w")


        

        preset_list_frame = ttk.LabelFrame(self.preset_editor_frame, text="Preset Task List")
        preset_list_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        self.preset_editor_frame.rowconfigure(1, weight=1)
        self.preset_editor_frame.columnconfigure(0, weight=1)

        self.preset_display_cols = ["Tasklist", "Stream", "Frequency", "Start Time", "End Time", "Action"]
        self.preset_tree = ttk.Treeview(preset_list_frame, columns=self.preset_display_cols, show='headings')
        self.preset_tree.heading('Tasklist', text='Tasklist')
        self.preset_tree.column('Tasklist', width=320)
        self.preset_tree.heading('Stream', text='Stream')
        self.preset_tree.column('Stream', width=90, minwidth=60, anchor='w', stretch=True)
        self.preset_tree.heading('Frequency', text='Frequency')
        self.preset_tree.column('Frequency', width=150)
        self.preset_tree.heading('Start Time', text='Start Time')
        self.preset_tree.column('Start Time', width=90)
        self.preset_tree.heading('End Time', text='End Time')
        self.preset_tree.column('End Time', width=90)
        self.preset_tree.heading('Action', text='Action')
        self.preset_tree.column('Action', width=70, minwidth=70, anchor='center', stretch=False)

        self.preset_tree.pack(side="left", fill="both", expand=True)
        self.preset_tree.bind("<ButtonRelease-1>", self.handle_preset_tree_click)
        self.preset_tree.bind("<Double-1>", self.edit_preset_task)

        preset_reorder_frame = ttk.Frame(preset_list_frame)
        preset_reorder_frame.pack(side="right", fill="y", padx=(0, 5))
        ttk.Button(preset_reorder_frame, text="▲", command=lambda: self.move_task(self.preset_tree, -1), width=3).pack(pady=2)
        ttk.Button(preset_reorder_frame, text="▼", command=lambda: self.move_task(self.preset_tree, 1), width=3).pack(pady=2)

        ttk.Button(self.preset_editor_frame, text="Save Current Presets", command=self.save_presets).grid(row=2, column=0, pady=10)
# Navigation handled via header buttons
    def on_preset_type_change(self, event=None):
        preset_type = self.preset_type_var.get()

        # Frequency is automatic for Daily/Weekday/Monthly presets.
        # Frequency is configurable for Task Dropdown Options.
        # Frequency field behavior:
        # - Daily presets: frequency is forced to Daily (hide field)
        # - Weekday/Monthly presets: allow selecting Weekly/Monthly/Weekly, Monthly
        # - Task Dropdown Options: allow selecting any supported frequency
        if preset_type in ("Weekday", "Monthly", "Task Dropdown Options"):
            self.preset_freq_label.grid()
            self.preset_frequency_combo.grid()
        else:
            self.preset_freq_label.grid_remove()
            self.preset_frequency_combo.grid_remove()
        if preset_type == "Daily":
            self.preset_key_var.set("")
            self.preset_key_combo.config(state="disabled", values=[])
            self.show_preset_editor("Daily")
        elif preset_type == "Weekday":
            weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            self.preset_key_combo.config(state="readonly", values=weekdays)
            self.preset_key_var.set(weekdays[0])
            self.show_preset_editor(weekdays[0])
        elif preset_type == "Monthly":
            days = [str(i) for i in range(1, 32)]
            self.preset_key_combo.config(state="readonly", values=days)
            self.preset_key_var.set(days[0])
            self.show_preset_editor(days[0])
        elif preset_type == "Task Dropdown Options":
            self.preset_key_var.set("")
            self.preset_key_combo.config(state="disabled", values=[])
            self.show_preset_editor("TaskDropdown")

    def on_preset_key_change(self, event=None):
        key = self.preset_key_var.get()
        if key:
            self.show_preset_editor(key)

    def show_preset_editor(self, key):
        preset_type = self.preset_type_var.get()
        tasks = []

        if preset_type == "Daily":
            self.preset_editor_frame.config(text="Editing 'Daily' Presets")
            tasks = self.presets.get("Daily", [])
        elif preset_type == "Weekday":
            self.preset_editor_frame.config(text=f"Editing Presets for '{key}'")
            tasks = self.presets.get("Weekdays", {}).get(key, [])
        elif preset_type == "Monthly":
            self.preset_editor_frame.config(text=f"Editing Presets for Day '{key}' of the Month")
            tasks = self.presets.get("Monthly", {}).get(key, [])
        elif preset_type == "Task Dropdown Options":
            self.preset_editor_frame.config(text="Editing Task Dropdown Options")
            # TaskDropdown supports per-task frequency (backward compatible with older string-only lists)
            self._normalize_taskdropdown_items()
            tasks = self.presets.get("TaskDropdown", [])

        self.preset_editor_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.preset_tree.delete(*self.preset_tree.get_children())
        for task_values in tasks:
            # Enforce automatic frequency for Daily/Weekday/Monthly presets
            try:
                ptype = self.preset_type_var.get()
                forced_freq = None
                # Only Daily presets force frequency; Weekday/Monthly can be mixed
                if ptype == 'Daily':
                    forced_freq = 'Daily'
                if forced_freq is not None:
                    if isinstance(task_values, (list, tuple)) and len(task_values) >= 1:
                        task_values = [task_values[0], forced_freq]
                    elif isinstance(task_values, str):
                        task_values = [task_values, forced_freq]
            except Exception:
                pass
            if isinstance(task_values, list) and len(task_values) >= 2:
                st = task_values[3] if len(task_values) > 3 else ''
                et = task_values[4] if len(task_values) > 4 else ''
                # task_values can be old: [task, freq, '', st, et] or new: [task, stream, freq, '', st, et]
                def _is_freq(x):
                    s = (str(x) if x is not None else '').strip().lower()
                    return s in ('daily','weekly','monthly') or 'weekly' in s or 'monthly' in s or 'daily' in s
                if len(task_values) > 1 and _is_freq(task_values[1]):
                    _stream = ''
                    _freq = task_values[1]
                    st = task_values[3] if len(task_values) > 3 else ''
                    et = task_values[4] if len(task_values) > 4 else ''
                else:
                    _stream = task_values[1] if len(task_values) > 1 else ''
                    _freq = task_values[2] if len(task_values) > 2 else ''
                    st = task_values[4] if len(task_values) > 4 else ''
                    et = task_values[5] if len(task_values) > 5 else ''
                self.preset_tree.insert('', tk.END, values=(task_values[0], _stream, _freq, st, et, '❌'))

    def handle_preset_tree_click(self, event):
        if self.preset_tree.identify_region(event.x, event.y) == "cell":
            if self.preset_tree.identify_column(event.x) == f'#{len(self.preset_display_cols)}':
                selected_item_id = self.preset_tree.focus()
                if selected_item_id:
                    self.preset_tree.delete(selected_item_id)
    def add_preset_task_to_preset_editor(self):
        tasklist = self.preset_tasklist_entry.get().strip()
        preset_type = self.preset_type_var.get()

        # Frequency rules for presets:
        # - Daily preset forces 'Daily'
        # - Weekday/Monthly presets allow selecting Weekly/Monthly/Weekly, Monthly
        # - Task Dropdown Options allow selecting any supported frequency
        if preset_type == 'Daily':
            frequency = 'Daily'
        else:
            frequency = self.preset_frequency_var.get().strip()
            _fl, _fd = normalize_frequency_string(frequency)
            frequency = _fd or frequency
            # sensible defaults if blank
            if not frequency and preset_type == 'Weekday':
                frequency = 'Weekly'
            if not frequency and preset_type == 'Monthly':
                frequency = 'Monthly'

        if not tasklist:
            messagebox.showwarning('Warning', 'Tasklist field cannot be empty.')
            return

        # Build start/end strings
        sh, sm, sa = self.preset_start_hour_var.get(), self.preset_start_minute_var.get(), self.preset_start_ampm_var.get()
        eh, em, ea = self.preset_end_hour_var.get(), self.preset_end_minute_var.get(), self.preset_end_ampm_var.get()
        start_time = f"{sh}:{sm} {sa}" if (sh and sm and sa) else ''
        end_time = f"{eh}:{em} {ea}" if (eh and em and ea) else ''

        stream = (getattr(self, 'preset_stream_var', tk.StringVar()).get() or '').strip()
        self.preset_tree.insert('', tk.END, values=(tasklist, stream, frequency, start_time, end_time, '❌'))

        self.preset_tasklist_entry.delete(0, tk.END)
        try:
            self.preset_stream_var.set('')
        except Exception:
            pass
        self.preset_frequency_combo.set('')
        try:
            self.preset_start_hour_var.set(''); self.preset_start_minute_var.set(''); self.preset_start_ampm_var.set('')
            self.preset_end_hour_var.set(''); self.preset_end_minute_var.set(''); self.preset_end_ampm_var.set('')
        except Exception:
            pass

    def edit_preset_task(self, event):
        # Action column is the last column (❌)
        if self.preset_tree.identify_column(event.x) == f'#{len(self.preset_display_cols)}':
            return

        selected_item_id = self.preset_tree.focus()
        if not selected_item_id:
            return

        edit_window = tk.Toplevel(self)
        edit_window.title("Edit Preset Task")

        values = list(self.preset_tree.item(selected_item_id, 'values'))
        tasklist = values[0] if len(values) > 0 else ''
        stream_val = values[1] if len(values) > 1 else ''
        frequency = values[2] if len(values) > 2 else ''
        start_time = values[3] if len(values) > 2 else ''
        end_time = values[4] if len(values) > 3 else ''

        ttk.Label(edit_window, text="Tasklist:").pack(padx=10, pady=(10, 0))
        task_entry = ttk.Entry(edit_window, width=60)
        task_entry.pack(padx=10, pady=5)
        task_entry.insert(0, tasklist)

        ttk.Label(edit_window, text='Stream:').pack(padx=10, pady=(10, 0))
        stream_var = tk.StringVar(value=stream_val)
        ttk.Entry(edit_window, width=30, textvariable=stream_var).pack(padx=10, pady=5, anchor='w')

        preset_type = self.preset_type_var.get()
        is_freq_editable = preset_type in ("Weekday", "Monthly", "Task Dropdown Options")

        forced_freq = None
        # Only Daily presets force frequency; Weekday/Monthly can be mixed
        if preset_type == 'Daily':
            forced_freq = 'Daily'
        freq_var = tk.StringVar(value=(forced_freq if forced_freq is not None else frequency))
        if is_freq_editable:
            ttk.Label(edit_window, text="Frequency:").pack(padx=10, pady=(10, 0))
            ttk.Combobox(edit_window, textvariable=freq_var, values=['', 'Daily', 'Weekly', 'Monthly', 'Weekly, Monthly']).pack(padx=10, pady=5)

        time_hours = [str(i).zfill(2) for i in range(1, 13)]
        time_minutes = [str(i).zfill(2) for i in range(60)]

        def split_time(t):
            try:
                dt = datetime.datetime.strptime(str(t).strip(), '%I:%M %p')
                return dt.strftime('%I'), dt.strftime('%M'), dt.strftime('%p')
            except Exception:
                return '', '', ''

        sh, sm, sa = split_time(start_time)
        eh, em, ea = split_time(end_time)

        ttk.Label(edit_window, text="Start Time:").pack(padx=10, pady=(10, 0))
        st_frame = ttk.Frame(edit_window)
        st_frame.pack(padx=10, pady=5, anchor='w')
        st_h = tk.StringVar(value=sh)
        st_m = tk.StringVar(value=sm)
        st_a = tk.StringVar(value=sa)
        ttk.Combobox(st_frame, textvariable=st_h, values=time_hours, width=3, state='readonly').pack(side='left')
        ttk.Combobox(st_frame, textvariable=st_m, values=time_minutes, width=3, state='readonly').pack(side='left', padx=(5, 0))
        ttk.Combobox(st_frame, textvariable=st_a, values=['AM', 'PM'], width=3, state='readonly').pack(side='left', padx=(5, 0))

        ttk.Label(edit_window, text="End Time:").pack(padx=10, pady=(10, 0))
        et_frame = ttk.Frame(edit_window)
        et_frame.pack(padx=10, pady=5, anchor='w')
        et_h = tk.StringVar(value=eh)
        et_m = tk.StringVar(value=em)
        et_a = tk.StringVar(value=ea)
        ttk.Combobox(et_frame, textvariable=et_h, values=time_hours, width=3, state='readonly').pack(side='left')
        ttk.Combobox(et_frame, textvariable=et_m, values=time_minutes, width=3, state='readonly').pack(side='left', padx=(5, 0))
        ttk.Combobox(et_frame, textvariable=et_a, values=['AM', 'PM'], width=3, state='readonly').pack(side='left', padx=(5, 0))

        def save_changes():
            new_task = task_entry.get().strip()
            new_stream = (stream_var.get() if 'stream_var' in locals() else '').strip()
            if not new_task:
                messagebox.showwarning('Warning', 'Tasklist field cannot be empty.')
                return

            new_freq = forced_freq if forced_freq is not None else freq_var.get().strip()
            new_start = f"{st_h.get()}:{st_m.get()} {st_a.get()}" if (st_h.get() and st_m.get() and st_a.get()) else ''
            new_end = f"{et_h.get()}:{et_m.get()} {et_a.get()}" if (et_h.get() and et_m.get() and et_a.get()) else ''

            self.preset_tree.item(selected_item_id, values=(new_task, new_stream, new_freq, new_start, new_end, '❌'))
            edit_window.destroy()

        ttk.Button(edit_window, text="Save Changes", command=save_changes).pack(pady=10)



# --- Main Execution ---
if __name__ == "__main__":
    import traceback

    print("run_app.py is running as __main__")

    try:
        print("Creating App()...")
        app = App()
        print("Entering mainloop()...")
        app.mainloop()
        print("mainloop() ended (window closed)")
    except Exception:
        print("CRASH:")
        traceback.print_exc()
        input("Press Enter to exit...")