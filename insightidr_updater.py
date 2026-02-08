#!/usr/bin/env python3
# InsightIDR Investigation Updater (CustomTkinter)
# - First-run API key setup (env var or key file), persisted in settings (cross-platform)
# - Configurable assignees (no hardcoded names/emails)
# - Bulk status/disposition/assignee updates via v2
# - Comments via v1: POST /idr/v1/comments  {"target": <RRN>, "body": <text>}
# - Default sort: Newest → Oldest (toggle supported)
# - Non-blocking operations with progress dialogs

import os
import re
import json
import time
import platform
import webbrowser
import threading
from urllib.parse import quote
from datetime import datetime, timezone
from typing import Dict, List, Tuple, Optional, Set

import requests
import customtkinter as ctk
import tkinter.messagebox as messagebox
import tkinter.filedialog as filedialog

# ----------------------------
# Config
# ----------------------------
OPENLIKE_STATUSES = ["OPEN", "INVESTIGATING", "WAITING"]
PAGE_SIZE = 100
START_DATE = "2024-01-01"
END_DATE = None

# API endpoints will be set dynamically based on region in settings
BASE_V2 = ""
BASE_V1 = ""
URL_V1_CREATE_COMMENT = ""

# These headers are filled after API key resolution
V2_HEADERS = {}
V1_HEADERS = {}

IDR_INVESTIGATION_TAIL_RE = re.compile(r":investigation:([^:]+)\s*$")

# ----------------------------
# Settings (cross-platform)
# ----------------------------
APP_DIR_NAME = "InsightIDRUpdater"
SETTINGS_FILENAME = "config.json"

def _app_support_dir() -> str:
    system = platform.system()
    if system == "Windows":
        base = os.getenv("APPDATA", os.path.expanduser("~"))
        return os.path.join(base, APP_DIR_NAME)
    elif system == "Darwin":  # macOS
        base = os.path.join(os.path.expanduser("~"), "Library", "Application Support")
        return os.path.join(base, APP_DIR_NAME)
    else:  # Linux/Other
        base = os.path.join(os.path.expanduser("~"), ".config")
        return os.path.join(base, APP_DIR_NAME)

def _settings_path() -> str:
    d = _app_support_dir()
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, SETTINGS_FILENAME)

def load_settings() -> Dict:
    p = _settings_path()
    if os.path.exists(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_settings(cfg: Dict):
    p = _settings_path()
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass

def resolve_api_key_from_settings(cfg: Dict) -> Optional[str]:
    """
    Settings schema:
    {
      "api_key_source": "env" | "file",
      "api_key_env_var": "R7_IDR_API_KEY",
      "api_key_file_path": "C:\\path\\to\\key.txt"
    }
    """
    src = cfg.get("api_key_source")
    if src == "env":
        var = cfg.get("api_key_env_var") or "R7_IDR_API_KEY"
        val = os.getenv(var)
        if val:
            return val.strip().strip('"').strip("'")
        return None
    if src == "file":
        path = cfg.get("api_key_file_path")
        if path and os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    val = f.read()
                return (val or "").strip().strip('"').strip("'")
            except Exception:
                return None
        return None
    return None

def get_assignees_from_settings(cfg: Dict) -> List[Tuple[str, str]]:
    """
    Get assignees from settings.
    Schema: {"assignees": [{"name": "Full Name", "email": "email@example.com"}, ...]}
    Returns: [("Full Name", "email@example.com"), ...]
    """
    assignees = cfg.get("assignees", [])
    if not assignees:
        return [("Unassigned / No Change", "")]
    result = [("Unassigned / No Change", "")]
    for a in assignees:
        name = a.get("name", "").strip()
        email = a.get("email", "").strip()
        if name and email:
            result.append((name, email))
    return result

def save_assignees_to_settings(cfg: Dict, assignees: List[Tuple[str, str]]):
    """Save assignees to settings (skip the first 'Unassigned' entry)."""
    cfg["assignees"] = [
        {"name": name, "email": email}
        for name, email in assignees
        if email  # Skip unassigned entry
    ]
    save_settings(cfg)

def get_region_from_settings(cfg: Dict) -> str:
    """Get region from settings, default to 'us' if not set."""
    return cfg.get("region", "us").strip()

def get_org_id_from_settings(cfg: Dict) -> str:
    """Get organization ID from settings."""
    return cfg.get("org_id", "").strip()

def save_region_org_to_settings(cfg: Dict, region: str, org_id: str):
    """Save region and org_id to settings."""
    cfg["region"] = region.strip()
    cfg["org_id"] = org_id.strip()
    save_settings(cfg)

def set_api_endpoints(region: str):
    """Set global API endpoints based on region."""
    global BASE_V2, BASE_V1, URL_V1_CREATE_COMMENT
    BASE_V2 = f"https://{region}.api.insight.rapid7.com/idr/v2/investigations"
    BASE_V1 = f"https://{region}.api.insight.rapid7.com/idr/v1"
    URL_V1_CREATE_COMMENT = f"{BASE_V1}/comments"

def console_link(rrn: str, region: str, org_id: str) -> str:
    """Generate console link for an investigation."""
    if not rrn or not org_id:
        return ""
    return f"https://{region}.idr.insight.rapid7.com/op/{org_id}#/investigations/{rrn}"

# ----------------------------
# Helpers
# ----------------------------
def iso_boundary(date_str: str, end_of_day: bool = False) -> str:
    if not date_str:
        return ""
    t = "23:59:59Z" if end_of_day else "00:00:00Z"
    return f"{date_str}T{t}"

def parse_iso_to_local(dt: Optional[str]) -> str:
    """Render API time to local 'YYYY/MM/DD HH:MM'."""
    if not dt:
        return ""
    try:
        if dt.endswith("Z"):
            dt_obj = datetime.fromisoformat(dt.replace("Z", "+00:00"))
        else:
            dt_obj = datetime.fromisoformat(dt)
        local = dt_obj.astimezone()  # system tz
        return local.strftime("%Y/%m/%d %H:%M")
    except Exception:
        return dt

def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# ----------------------------
# API calls
# ----------------------------
def list_investigations() -> List[Dict]:
    """GET v2 investigations with open-like statuses (oldest->newest; UI will sort)."""
    index = 0
    all_items: List[Dict] = []
    statuses_str = ",".join(OPENLIKE_STATUSES)
    params_base = {
        "size": PAGE_SIZE,
        "statuses": statuses_str,
        "sort": "created_time,ASC",
        "start_time": iso_boundary(START_DATE, end_of_day=False),
    }
    if END_DATE:
        params_base["end_time"] = iso_boundary(END_DATE, end_of_day=True)

    while True:
        params = dict(params_base)
        params["index"] = index
        r = requests.get(BASE_V2, headers=V2_HEADERS, params=params, timeout=60)
        r.raise_for_status()
        payload = r.json() or {}
        items = payload.get("data", []) or []
        meta = payload.get("metadata", {}) or {}

        if not items:
            break

        all_items.extend(items)
        total_pages = int(meta.get("total_pages", 0) or 0)
        current_index = int(meta.get("index", index))
        if total_pages == 0 or current_index >= (total_pages - 1):
            break

        index += 1
        time.sleep(0.04)

    return all_items

def set_status(id_or_rrn: str, new_status: str):
    url = f"{BASE_V2}/{quote(id_or_rrn, safe='')}/status/{quote(new_status, safe='')}"
    r = requests.put(url, headers=V2_HEADERS, timeout=60)
    r.raise_for_status()

def set_disposition(id_or_rrn: str, disposition: str):
    url = f"{BASE_V2}/{quote(id_or_rrn, safe='')}/disposition/{quote(disposition, safe='')}"
    r = requests.put(url, headers=V2_HEADERS, timeout=60)
    r.raise_for_status()

def assign_user(id_or_rrn: str, assignee_email: str):
    url_patch = f"{BASE_V2}/{quote(id_or_rrn, safe='')}"
    body = {"assignee": {"email": assignee_email}}
    r = requests.patch(url_patch, headers=V2_HEADERS, json=body, timeout=60)
    if r.status_code in (200, 204):
        return
    url_put = f"{BASE_V2}/{quote(id_or_rrn, safe='')}/assignee/{quote(assignee_email, safe='')}"
    r2 = requests.put(url_put, headers=V2_HEADERS, timeout=60)
    r2.raise_for_status()

def get_rrn(id_or_rrn: str) -> str:
    """Return an investigation RRN. If given an ID, fetch v2 record and return its rrn."""
    if not id_or_rrn:
        raise ValueError("Missing investigation id/rrn")
    if id_or_rrn.startswith("rrn:"):
        return id_or_rrn
    # v2 GET by ID (endpoint accepts id or rrn)
    url = f"{BASE_V2}/{quote(id_or_rrn, safe='')}"
    r = requests.get(url, headers=V2_HEADERS, timeout=60)
    r.raise_for_status()
    j = r.json() or {}
    inv = j.get("data") if isinstance(j, dict) and "data" in j else j
    if not isinstance(inv, dict):
        raise RuntimeError("Unexpected v2 GET response shape")
    rrn = inv.get("rrn")
    if not rrn:
        raise RuntimeError("Could not find RRN in v2 response")
    return rrn

def create_comment_v1(target_rrn: str, text: str) -> dict:
    """
    POST /idr/v1/comments
    body = {"target": <RRN>, "body": <text>}
    Returns dict with ok/status/text/url/body.
    """
    if not text:
        return {"ok": True, "status": 204, "url": URL_V1_CREATE_COMMENT, "text": "", "body": {}}
    payload = {"target": target_rrn, "body": text}
    r = requests.post(URL_V1_CREATE_COMMENT, headers=V1_HEADERS, json=payload, timeout=60)
    return {
        "ok": r.status_code in (200, 201),
        "status": r.status_code,
        "url": URL_V1_CREATE_COMMENT,
        "text": (r.text or "").strip(),
        "body": payload,
    }

def get_comments_v1(target_rrn: str) -> List[Dict]:
    """
    GET /idr/v1/comments?target=<RRN>
    Returns list of comments for the investigation.
    """
    if not target_rrn:
        return []
    try:
        url = f"{BASE_V1}/comments"
        params = {"target": target_rrn}
        r = requests.get(url, headers=V1_HEADERS, params=params, timeout=60)
        r.raise_for_status()
        data = r.json()
        # API returns {"data": [...]} 
        if isinstance(data, dict):
            return data.get("data", [])
        return data if isinstance(data, list) else []
    except Exception:
        return []

# ----------------------------
# Progress Dialog
# ----------------------------
class ProgressDialog(ctk.CTkToplevel):
    def __init__(self, parent, title="Processing..."):
        super().__init__(parent)
        self.title(title)
        self.geometry("500x200")
        self.resizable(False, False)
        
        self.label = ctk.CTkLabel(self, text="Please wait...", font=ctk.CTkFont(size=14))
        self.label.pack(pady=(20, 10))
        
        self.progress = ctk.CTkProgressBar(self, width=400, mode="indeterminate")
        self.progress.pack(pady=10)
        self.progress.start()
        
        self.detail_label = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=12))
        self.detail_label.pack(pady=10)
        
        # Center on parent
        self.transient(parent)
        self.grab_set()
        
    def update_message(self, message: str):
        self.label.configure(text=message)
        self.update()
        
    def update_detail(self, detail: str):
        self.detail_label.configure(text=detail)
        self.update()

# ----------------------------
# Assignee Configuration Dialog
# ----------------------------
class AssigneeConfigDialog(ctk.CTkToplevel):
    def __init__(self, parent, current_assignees: List[Tuple[str, str]]):
        super().__init__(parent)
        self.title("Configure Assignees")
        self.geometry("650x500")
        
        self.result = None
        # Skip first entry (Unassigned)
        self.assignees = [list(a) for a in current_assignees[1:]]
        
        # Header
        ctk.CTkLabel(
            self, 
            text="Configure Team Members", 
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=(15, 5))
        
        ctk.CTkLabel(
            self,
            text="Add team members who can be assigned to investigations",
            font=ctk.CTkFont(size=12)
        ).pack(pady=(0, 15))
        
        # Scrollable frame for assignees
        self.scroll_frame = ctk.CTkScrollableFrame(self, height=300)
        self.scroll_frame.pack(fill="both", expand=True, padx=15, pady=(0, 10))
        
        self.entry_rows = []
        self.rebuild_list()
        
        # Buttons
        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(fill="x", padx=15, pady=(5, 15))
        
        ctk.CTkButton(
            btn_frame, 
            text="➕ Add Person",
            command=self.add_row,
            width=120
        ).pack(side="left", padx=(0, 10))
        
        ctk.CTkButton(
            btn_frame,
            text="Save",
            command=self.save,
            width=100
        ).pack(side="right", padx=(10, 0))
        
        ctk.CTkButton(
            btn_frame,
            text="Cancel",
            command=self.destroy,
            width=100
        ).pack(side="right")
        
        self.transient(parent)
        self.grab_set()
        
    def rebuild_list(self):
        # Clear existing
        for row in self.entry_rows:
            for widget in row:
                widget.destroy()
        self.entry_rows.clear()
        
        # Header row
        header_frame = ctk.CTkFrame(self.scroll_frame)
        header_frame.pack(fill="x", padx=5, pady=(5, 10))
        ctk.CTkLabel(header_frame, text="Full Name", width=200, anchor="w").pack(side="left", padx=5)
        ctk.CTkLabel(header_frame, text="Email Address", width=280, anchor="w").pack(side="left", padx=5)
        
        # Add rows
        for idx, (name, email) in enumerate(self.assignees):
            self.add_row_ui(name, email, idx)
    
    def add_row_ui(self, name="", email="", idx=None):
        row_frame = ctk.CTkFrame(self.scroll_frame)
        row_frame.pack(fill="x", padx=5, pady=3)
        
        name_entry = ctk.CTkEntry(row_frame, width=200, placeholder_text="John Doe")
        name_entry.pack(side="left", padx=5)
        name_entry.insert(0, name)
        
        email_entry = ctk.CTkEntry(row_frame, width=280, placeholder_text="jdoe@company.com")
        email_entry.pack(side="left", padx=5)
        email_entry.insert(0, email)
        
        if idx is not None:
            def remove_this():
                self.assignees.pop(idx)
                self.rebuild_list()
            remove_btn = ctk.CTkButton(
                row_frame, 
                text="✕", 
                width=40,
                command=remove_this,
                fg_color="red",
                hover_color="darkred"
            )
            remove_btn.pack(side="left", padx=5)
            self.entry_rows.append([row_frame, name_entry, email_entry, remove_btn])
        else:
            self.entry_rows.append([row_frame, name_entry, email_entry])
    
    def add_row(self):
        self.assignees.append(["", ""])
        self.rebuild_list()
    
    def save(self):
        # Collect all entries
        collected = []
        for row in self.entry_rows:
            if len(row) >= 3:
                name = row[1].get().strip()
                email = row[2].get().strip()
                if name and email:
                    collected.append((name, email))
        
        if not collected:
            messagebox.showwarning("No Assignees", "Please add at least one team member.")
            return
        
        # Add back the unassigned option
        self.result = [("Unassigned / No Change", "")] + collected
        self.destroy()

# ----------------------------
# UI
# ----------------------------
STATUSES = ["OPEN", "INVESTIGATING", "WAITING", "CLOSED"]
DISPOSITIONS = ["", "BENIGN", "MALICIOUS", "NOT_APPLICABLE"]  # "" = don't send

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
        self.title("InsightIDR Investigation Updater")
        self.geometry("1600x900")  # Wider for 3-column layout

        # sort state — default NEWEST first
        self.sort_oldest_first = False   # default newest→oldest

        # assignee filter state
        self.assignee_filter_var = ctk.StringVar(value="All")

        # fonts
        self.font_title_bold = ctk.CTkFont(weight="bold", size=18)   # bigger title
        self.font_bold_red = ctk.CTkFont(weight="bold")

        # ----- Top bar -----
        top = ctk.CTkFrame(self)
        top.pack(fill="x", padx=12, pady=(12, 6))

        self.refresh_btn = ctk.CTkButton(top, text="Refresh", command=self.refresh_async)
        self.refresh_btn.pack(side="left", padx=(6, 12), pady=8)

        self.select_all_var = ctk.BooleanVar(value=False)
        self.select_all_box = ctk.CTkCheckBox(
            top, text="Select All", variable=self.select_all_var,
            command=self.toggle_select_all
        )
        self.select_all_box.pack(side="left", padx=6, pady=8)

        self.sort_switch_var = ctk.BooleanVar(value=self.sort_oldest_first)
        self.sort_switch = ctk.CTkSwitch(
            top,
            text="Sort by Created Time (Newest → Oldest)" if not self.sort_oldest_first else "Sort by Created Time (Oldest → Newest)",
            variable=self.sort_switch_var,
            onvalue=True,
            offvalue=False,
            command=self.on_sort_toggle
        )
        self.sort_switch.pack(side="left", padx=12)

        # Assignee Filter
        ctk.CTkLabel(top, text="Filter by Assignee:").pack(side="left", padx=(20,6))
        self.assignee_filter = ctk.CTkOptionMenu(
            top, values=["All"], variable=self.assignee_filter_var, command=lambda *_: self.rebuild_list()
        )
        self.assignee_filter.pack(side="left", padx=(0, 8))

        # ----- Main middle layout (3-column) -----
        mid = ctk.CTkFrame(self)
        mid.pack(fill="both", expand=True, padx=12, pady=6)

        # Left column: Investigation list
        left = ctk.CTkFrame(mid)
        left.pack(side="left", fill="both", expand=True, padx=(0, 8), pady=8)

        header = ctk.CTkLabel(
            left,
            text="Investigations (OPEN / INVESTIGATING / WAITING)",
            anchor="w"
        )
        header.pack(fill="x", padx=8, pady=(8, 4))

        self.scroll = ctk.CTkScrollableFrame(left, height=680)
        self.scroll.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        # Middle column: Action controls
        middle = ctk.CTkFrame(mid, width=380)
        middle.pack(side="left", fill="y", padx=(0, 8), pady=8)

        form = ctk.CTkFrame(middle)
        form.pack(fill="x", padx=10, pady=10)

        ctk.CTkLabel(form, text="New Status").grid(row=0, column=0, sticky="w", padx=6, pady=(8,4))
        self.status_choice = ctk.CTkOptionMenu(form, values=STATUSES)
        self.status_choice.set("OPEN")
        self.status_choice.grid(row=0, column=1, sticky="ew", padx=6, pady=(8,4))

        ctk.CTkLabel(form, text="Disposition").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        self.dispo_choice = ctk.CTkOptionMenu(form, values=DISPOSITIONS)
        self.dispo_choice.set("")
        self.dispo_choice.grid(row=1, column=1, sticky="ew", padx=6, pady=4)

        ctk.CTkLabel(form, text="Assignee").grid(row=2, column=0, sticky="w", padx=6, pady=4)
        self.assignee_choice = ctk.CTkOptionMenu(form, values=["Unassigned / No Change"])
        self.assignee_choice.set("Unassigned / No Change")
        self.assignee_choice.grid(row=2, column=1, sticky="ew", padx=6, pady=4)

        form.grid_columnconfigure(1, weight=1)

        # Configure Assignees button
        ctk.CTkButton(
            middle, 
            text="⚙️ Configure Team Members", 
            command=self.configure_assignees
        ).pack(fill="x", padx=16, pady=(0, 10))

        # Comment box + Clear
        ctk.CTkLabel(middle, text="Comment:").pack(anchor="w", padx=16, pady=(12, 4))
        self.comment_box = ctk.CTkTextbox(middle, height=120)
        self.comment_box.pack(fill="x", padx=16, pady=(0, 8))

        # Update button
        self.update_btn = ctk.CTkButton(middle, text="Update Selected", command=self.update_selected_async)
        self.update_btn.pack(fill="x", padx=16, pady=(4, 6))

        def clear_comment():
            self.comment_box.delete("1.0", "end")
            self._log_status("Comment box cleared.")
        ctk.CTkButton(middle, text="Clear Comment Box", command=clear_comment).pack(fill="x", padx=16, pady=(0, 10))

        # Troubleshoot button
        self.test_btn = ctk.CTkButton(middle, text="Test Comment (Troubleshoot)", command=self.test_comment_popup)
        self.test_btn.pack(fill="x", padx=16, pady=(0, 10))

        # Settings box (API key source + Region/Org)
        settings_frame = ctk.CTkFrame(middle)
        settings_frame.pack(fill="x", padx=10, pady=(0,10))
        
        # Header
        ctk.CTkLabel(
            settings_frame, 
            text="Settings", 
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=10, pady=(10, 8))
        
        # Info display area
        info_frame = ctk.CTkFrame(settings_frame)
        info_frame.pack(fill="x", padx=10, pady=(0,10))
        
        self.settings_info_label = ctk.CTkLabel(
            info_frame, 
            text="", 
            anchor="w", 
            justify="left", 
            font=ctk.CTkFont(size=11)
        )
        self.settings_info_label.pack(fill="x", padx=8, pady=8)
        
        # Buttons in grid layout for uniformity
        btn_container = ctk.CTkFrame(settings_frame)
        btn_container.pack(fill="x", padx=10, pady=(0,10))
        
        ctk.CTkButton(
            btn_container, 
            text="🔑 API Key Settings", 
            command=self.open_api_key_settings_dialog,
            height=32
        ).pack(fill="x", padx=6, pady=(0,6))
        
        ctk.CTkButton(
            btn_container,
            text="🌐 Region & Organization",
            command=self.open_region_org_dialog,
            height=32
        ).pack(fill="x", padx=6, pady=(0,6))

        # Right column: Status Log, Comments, and Comment History tabs
        right = ctk.CTkFrame(mid, width=500)
        right.pack(side="right", fill="both", expand=False, padx=(0, 0), pady=8)
        
        # Tab buttons
        tab_btn_frame = ctk.CTkFrame(right)
        tab_btn_frame.pack(fill="x", padx=8, pady=(8,0))
        
        self.status_tab_btn = ctk.CTkButton(
            tab_btn_frame, 
            text="📋 Status Log",
            command=lambda: self._switch_tab("status"),
            width=120
        )
        self.status_tab_btn.pack(side="left", padx=(0,6))
        
        self.comments_tab_btn = ctk.CTkButton(
            tab_btn_frame,
            text="💬 Comments", 
            command=lambda: self._switch_tab("comments"),
            width=120
        )
        self.comments_tab_btn.pack(side="left", padx=(0,6))
        
        self.history_tab_btn = ctk.CTkButton(
            tab_btn_frame,
            text="📝 History",
            command=lambda: self._switch_tab("history"),
            width=100
        )
        self.history_tab_btn.pack(side="left")
        
        # Tab content frames
        self.status_content = ctk.CTkFrame(right)
        self.comments_content = ctk.CTkFrame(right)
        self.history_content = ctk.CTkFrame(right)
        
        # Status tab content
        self.status_box = ctk.CTkTextbox(self.status_content, height=600, wrap="word")
        self.status_box.pack(fill="both", expand=True, padx=8, pady=8)
        
        # Comments tab content
        comments_header = ctk.CTkFrame(self.comments_content)
        comments_header.pack(fill="x", padx=8, pady=(8,4))
        
        ctk.CTkLabel(
            comments_header, 
            text="Select investigation to view comments",
            font=ctk.CTkFont(size=11, slant="italic")
        ).pack(side="left", padx=(0,8))
        
        refresh_comments_btn = ctk.CTkButton(
            comments_header,
            text="🔄 Refresh",
            command=self.refresh_selected_comments,
            width=100
        )
        refresh_comments_btn.pack(side="right")
        
        self.comments_box = ctk.CTkTextbox(self.comments_content, height=600, wrap="word")
        self.comments_box.pack(fill="both", expand=True, padx=8, pady=(0,8))
        
        # Comment History tab content
        history_header = ctk.CTkFrame(self.history_content)
        history_header.pack(fill="x", padx=8, pady=(8,4))
        
        ctk.CTkLabel(
            history_header,
            text="Recent comment history (click to copy)",
            font=ctk.CTkFont(size=11, slant="italic")
        ).pack(side="left", padx=(0,8))
        
        clear_history_btn = ctk.CTkButton(
            history_header,
            text="🗑️ Clear",
            command=self.clear_comment_history,
            width=80
        )
        clear_history_btn.pack(side="right")
        
        self.history_box = ctk.CTkTextbox(self.history_content, height=600, wrap="word")
        self.history_box.pack(fill="both", expand=True, padx=8, pady=(0,8))
        
        # Initialize comment history
        self.comment_history = []
        self._load_comment_history()
        self._display_comment_history()
        
        # Show status tab by default
        self.current_tab = "status"
        self._switch_tab("status")

        # data holders
        self.rows: List[Dict] = []
        self.row_vars: List[ctk.BooleanVar] = []
        self.row_frames: List[ctk.CTkFrame] = []

        # ----- Resolve API key and assignees -----
        self.cfg = load_settings()
        
        # Set region and API endpoints
        region = get_region_from_settings(self.cfg)
        org_id = get_org_id_from_settings(self.cfg)
        
        # Prompt for region/org if not set
        if not region or not org_id:
            self.first_run_region_org_setup()
            self.cfg = load_settings()
            region = get_region_from_settings(self.cfg)
            org_id = get_org_id_from_settings(self.cfg)
        
        set_api_endpoints(region)
        self.region = region
        self.org_id = org_id
        
        # Check for assignees - if none, prompt for configuration
        assignees = get_assignees_from_settings(self.cfg)
        if len(assignees) <= 1:  # Only has "Unassigned"
            self.first_run_assignee_setup()
            self.cfg = load_settings()
        
        self._refresh_assignee_dropdown()
        
        # API key setup
        api_key = resolve_api_key_from_settings(self.cfg)
        if not api_key:
            self.first_run_api_key_setup()
            self.cfg = load_settings()
            api_key = resolve_api_key_from_settings(self.cfg)

        self._refresh_settings_label()

        if not api_key:
            messagebox.showerror("API Key Missing", "Could not resolve API key. Use the API Key Settings section.")
        else:
            self._set_headers(api_key)
            self.refresh_async()

    # ------- Tab switching -------
    def _switch_tab(self, tab_name: str):
        """Switch between Status, Comments, and History tabs."""
        self.current_tab = tab_name
        
        # Hide all tabs
        self.status_content.pack_forget()
        self.comments_content.pack_forget()
        self.history_content.pack_forget()
        
        # Reset all button colors
        self.status_tab_btn.configure(fg_color=["gray70", "gray30"])
        self.comments_tab_btn.configure(fg_color=["gray70", "gray30"])
        self.history_tab_btn.configure(fg_color=["gray70", "gray30"])
        
        # Show selected tab and highlight button
        if tab_name == "status":
            self.status_content.pack(fill="both", expand=True, padx=0, pady=(6,6))
            self.status_tab_btn.configure(fg_color=["#3B8ED0", "#1F6AA5"])
        elif tab_name == "comments":
            self.comments_content.pack(fill="both", expand=True, padx=0, pady=(6,6))
            self.comments_tab_btn.configure(fg_color=["#3B8ED0", "#1F6AA5"])
        else:  # history
            self.history_content.pack(fill="both", expand=True, padx=0, pady=(6,6))
            self.history_tab_btn.configure(fg_color=["#3B8ED0", "#1F6AA5"])
    
    # ------- Comment History Management -------
    def _load_comment_history(self):
        """Load comment history from settings."""
        cfg = load_settings()
        self.comment_history = cfg.get("comment_history", [])
        # Keep only last 50 comments
        if len(self.comment_history) > 50:
            self.comment_history = self.comment_history[-50:]
    
    def _save_comment_history(self):
        """Save comment history to settings."""
        cfg = load_settings()
        cfg["comment_history"] = self.comment_history[-50:]  # Keep only last 50
        save_settings(cfg)
    
    def _add_to_comment_history(self, comment_text: str):
        """Add a comment to history with timestamp."""
        if not comment_text or not comment_text.strip():
            return
        
        timestamp = now_str()
        entry = {
            "timestamp": timestamp,
            "text": comment_text.strip()
        }
        self.comment_history.append(entry)
        self._save_comment_history()
        self._display_comment_history()
        self._log_status("Comment added to history")
    
    def _display_comment_history(self):
        """Display comment history in the history tab."""
        self.history_box.delete("1.0", "end")
        
        if not self.comment_history:
            self.history_box.insert("1.0", "No comment history yet.\n\nComments you send will appear here for easy reuse.")
            return
        
        header = f"Comment History ({len(self.comment_history)} entries)\n"
        header += "Click on any comment text to select and copy it.\n"
        header += "=" * 60 + "\n\n"
        self.history_box.insert("1.0", header)
        
        # Show newest first
        for idx, entry in enumerate(reversed(self.comment_history), 1):
            timestamp = entry.get("timestamp", "")
            text = entry.get("text", "")
            
            comment_entry = f"[{idx}] {timestamp}\n"
            comment_entry += f"{text}\n"
            comment_entry += "-" * 60 + "\n\n"
            
            self.history_box.insert("end", comment_entry)
    
    def clear_comment_history(self):
        """Clear all comment history."""
        result = messagebox.askyesno(
            "Clear History",
            "Are you sure you want to clear all comment history?\n\nThis cannot be undone."
        )
        if result:
            self.comment_history = []
            self._save_comment_history()
            self._display_comment_history()
            self._log_status("Comment history cleared")
    
    def refresh_selected_comments(self):
        """Fetch and display comments for the first selected investigation."""
        rows = self._selected_rows()
        if not rows:
            self.comments_box.delete("1.0", "end")
            self.comments_box.insert("1.0", "No investigation selected.\n\nPlease select an investigation from the list to view its comments.")
            return
        
        rec = rows[0]
        title = rec.get("title", "")
        rrn = rec.get("rrn") or ""
        inv_id = rec.get("id") or ""
        
        if not rrn and inv_id:
            try:
                rrn = get_rrn(inv_id)
            except Exception:
                rrn = ""
        
        if not rrn:
            self.comments_box.delete("1.0", "end")
            self.comments_box.insert("1.0", "Could not resolve RRN for this investigation.")
            return
        
        # Fetch comments in background
        def worker():
            try:
                comments = get_comments_v1(rrn)
                self.after(0, lambda: self._display_comments(title, comments))
            except Exception as e:
                self.after(0, lambda: self._display_comments_error(str(e)))
        
        self.comments_box.delete("1.0", "end")
        self.comments_box.insert("1.0", f"Loading comments for:\n{title}\n\nPlease wait...")
        
        thread = threading.Thread(target=worker, daemon=True)
        thread.start()
    
    def _display_comments(self, title: str, comments: List[Dict]):
        """Display fetched comments in the comments box."""
        self.comments_box.delete("1.0", "end")
        
        header = f"Investigation: {title}\n"
        header += f"Total Comments: {len(comments)}\n"
        header += "=" * 60 + "\n\n"
        self.comments_box.insert("1.0", header)
        
        if not comments:
            self.comments_box.insert("end", "No comments found for this investigation.")
            return
        
        # Sort by timestamp (newest first)
        sorted_comments = sorted(
            comments, 
            key=lambda c: c.get("created_time", ""),
            reverse=True
        )
        
        for idx, comment in enumerate(sorted_comments, 1):
            creator = comment.get("creator", {})
            creator_name = creator.get("name", "Unknown")
            creator_email = creator.get("email", "")
            created = parse_iso_to_local(comment.get("created_time", ""))
            body = comment.get("body", "")
            
            comment_text = f"[{idx}] {creator_name}"
            if creator_email:
                comment_text += f" <{creator_email}>"
            comment_text += f"\n    Time: {created}\n"
            comment_text += f"    {body}\n\n"
            
            self.comments_box.insert("end", comment_text)
        
        self._log_status(f"Loaded {len(comments)} comment(s) for investigation")
    
    def _display_comments_error(self, error: str):
        """Display error when fetching comments fails."""
        self.comments_box.delete("1.0", "end")
        self.comments_box.insert("1.0", f"Error fetching comments:\n\n{error}")

    # ------- Region/Org configuration -------
    def first_run_region_org_setup(self):
        """First-run region and org ID setup."""
        self._log_status("First-run region/org setup started.")
        
        msg = "Welcome! Please configure your Rapid7 InsightIDR region and organization ID.\n\n" \
              "Common regions: us, us2, us3, eu, ca, au, ap\n" \
              "You can find your Org ID in the InsightIDR console URL."
        messagebox.showinfo("Region & Organization Setup", msg)
        
        self.open_region_org_dialog()
    
    def open_region_org_dialog(self):
        """Open dialog to configure region and org ID."""
        def save_and_close():
            region = region_entry.get().strip()
            org_id = org_entry.get().strip()
            
            if not region:
                messagebox.showwarning("Required", "Region is required (e.g., us, us2, us3, eu)")
                return
            
            cfg = load_settings()
            save_region_org_to_settings(cfg, region, org_id)
            set_api_endpoints(region)
            self.region = region
            self.org_id = org_id
            self._refresh_settings_label()
            dialog.destroy()
            self._log_status(f"Region/Org updated: {region} / {org_id or '(none)'}")
        
        dialog = ctk.CTkToplevel(self)
        dialog.title("Configure Region & Organization")
        dialog.geometry("500x280")
        frame = ctk.CTkFrame(dialog)
        frame.pack(fill="both", expand=True, padx=12, pady=12)
        
        ctk.CTkLabel(
            frame, 
            text="InsightIDR Region & Organization", 
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=10, pady=(8,12))
        
        # Region
        ctk.CTkLabel(frame, text="Region (e.g., us, us2, us3, eu, ca, au, ap):").pack(anchor="w", padx=10, pady=(0,4))
        region_entry = ctk.CTkEntry(frame, width=400, placeholder_text="us3")
        region_entry.pack(padx=10, pady=(0,12))
        current_region = get_region_from_settings(load_settings())
        if current_region:
            region_entry.insert(0, current_region)
        
        # Org ID
        ctk.CTkLabel(frame, text="Organization ID (optional, for console links):").pack(anchor="w", padx=10, pady=(0,4))
        org_entry = ctk.CTkEntry(frame, width=400, placeholder_text="Find in your InsightIDR console URL")
        org_entry.pack(padx=10, pady=(0,12))
        current_org = get_org_id_from_settings(load_settings())
        if current_org:
            org_entry.insert(0, current_org)
        
        # Buttons
        btn_frame = ctk.CTkFrame(frame)
        btn_frame.pack(fill="x", padx=10, pady=(12,8))
        ctk.CTkButton(btn_frame, text="Save", command=save_and_close, width=100).pack(side="left", padx=(0,8))
        ctk.CTkButton(btn_frame, text="Cancel", command=dialog.destroy, width=100).pack(side="left")
        
        dialog.transient(self)
        dialog.grab_set()
        self.wait_window(dialog)

    # ------- Assignee configuration -------
    def first_run_assignee_setup(self):
        """First-run assignee configuration."""
        self._log_status("First-run assignee setup started.")
        
        msg = "Welcome! Before you start, please configure your team members.\n\n" \
              "Add the full names and email addresses of people who can be\n" \
              "assigned to investigations."
        messagebox.showinfo("Team Configuration", msg)
        
        self.configure_assignees()
    
    def configure_assignees(self):
        """Open assignee configuration dialog."""
        current = get_assignees_from_settings(self.cfg)
        dialog = AssigneeConfigDialog(self, current)
        self.wait_window(dialog)
        
        if dialog.result:
            save_assignees_to_settings(self.cfg, dialog.result)
            self.cfg = load_settings()
            self._refresh_assignee_dropdown()
            self._log_status("Team members updated.")
    
    def _refresh_assignee_dropdown(self):
        """Refresh the assignee dropdown with current settings."""
        assignees = get_assignees_from_settings(self.cfg)
        labels = [f"{n} <{e}>" if e else n for n, e in assignees]
        self.assignee_choice.configure(values=labels)
        if labels:
            self.assignee_choice.set(labels[0])

    # ------- API key setup -------
    def _set_headers(self, api_key: str):
        global V1_HEADERS, V2_HEADERS
        V2_HEADERS = {
            "X-Api-Key": api_key,
            "Accept": "application/json",
            "Content-Type": "application/json",
            "Accept-version": "investigations-preview",
        }
        V1_HEADERS = {
            "X-Api-Key": api_key,
            "Accept": "application/json",
            "Content-Type": "application/json",
        }

    def _refresh_settings_label(self):
        cfg = load_settings()
        src = cfg.get("api_key_source")
        region = get_region_from_settings(cfg)
        org_id = get_org_id_from_settings(cfg)
        
        # Build clean, aligned text
        lines = []
        
        # API Key info
        if src == "env":
            var = cfg.get("api_key_env_var", "R7_IDR_API_KEY")
            lines.append(f"API Key: Environment Variable")
            lines.append(f"  Variable: {var}")
        elif src == "file":
            path = cfg.get("api_key_file_path", "")
            lines.append(f"API Key: File")
            lines.append(f"  Path: {path or '(not set)'}")
        else:
            lines.append("API Key: Not configured")
        
        lines.append("")  # Blank line
        lines.append(f"Region: {region}")
        lines.append(f"Org ID: {org_id or '(not set)'}")
        
        self.settings_info_label.configure(text="\n".join(lines))

    def first_run_api_key_setup(self):
        """Prompt for API key source (env var or file). Save to settings."""
        self._log_status("First-run API key setup started.")

        def save_and_close(src, env_var, path):
            cfg = load_settings()
            cfg["api_key_source"] = src
            if src == "env":
                cfg["api_key_env_var"] = env_var or "R7_IDR_API_KEY"
                cfg.pop("api_key_file_path", None)
            else:
                cfg["api_key_file_path"] = path or ""
                cfg.pop("api_key_env_var", None)
            save_settings(cfg)
            dialog.destroy()
            # Try load right away
            api_key = resolve_api_key_from_settings(cfg)
            self._refresh_settings_label()
            if api_key:
                self._set_headers(api_key)
                self._log_status("API key saved.")
            else:
                messagebox.showwarning("API Key", "Saved, but key could not be read. Check your selection.")

        dialog = ctk.CTkToplevel(self)
        dialog.title("API Key Setup")
        dialog.geometry("560x320")
        frame = ctk.CTkFrame(dialog)
        frame.pack(fill="both", expand=True, padx=12, pady=12)

        mode_var = ctk.StringVar(value="env")

        # Mode select
        ctk.CTkLabel(frame, text="Choose API key source:").pack(anchor="w", padx=10, pady=(8,4))
        opt_row = ctk.CTkFrame(frame); opt_row.pack(fill="x", padx=10, pady=(0,8))
        ctk.CTkRadioButton(opt_row, text="Environment Variable", variable=mode_var, value="env").pack(side="left", padx=(0,10))
        ctk.CTkRadioButton(opt_row, text="Key File (text file with only the key)", variable=mode_var, value="file").pack(side="left", padx=(10,0))

        # Env var row
        env_frame = ctk.CTkFrame(frame); env_frame.pack(fill="x", padx=10, pady=(4,8))
        ctk.CTkLabel(env_frame, text="Env var name:").pack(side="left", padx=(6,6))
        env_entry = ctk.CTkEntry(env_frame, width=240)
        env_entry.insert(0, "R7_IDR_API_KEY")
        env_entry.pack(side="left", padx=(0,8))

        # File row
        file_frame = ctk.CTkFrame(frame); file_frame.pack(fill="x", padx=10, pady=(4,8))
        file_path_var = ctk.StringVar(value="")
        def browse_file():
            p = filedialog.askopenfilename(title="Select key file", filetypes=[("Text files","*.txt *.key *.conf *.cfg *.json"),("All files","*.*")])
            if p:
                file_path_var.set(p)
        ctk.CTkLabel(file_frame, text="Key file path:").pack(side="left", padx=(6,6))
        file_entry = ctk.CTkEntry(file_frame, width=320, textvariable=file_path_var)
        file_entry.pack(side="left", padx=(0,8))
        ctk.CTkButton(file_frame, text="Browse…", width=80, command=browse_file).pack(side="left")

        # Buttons
        btns = ctk.CTkFrame(frame); btns.pack(fill="x", padx=10, pady=(8,8))
        def on_save():
            src = mode_var.get()
            if src == "env":
                save_and_close("env", env_entry.get().strip(), "")
            else:
                save_and_close("file", "", file_path_var.get().strip())
        ctk.CTkButton(btns, text="Save", command=on_save).pack(side="left", padx=(0,8))
        ctk.CTkButton(btns, text="Cancel", command=dialog.destroy).pack(side="left", padx=(8,0))

        dialog.transient(self)
        dialog.grab_set()
        self.wait_window(dialog)

    def open_api_key_settings_dialog(self):
        """Allow changing API key source anytime."""
        self.first_run_api_key_setup()
        # Try to refresh after change
        api_key = resolve_api_key_from_settings(load_settings())
        if api_key:
            self._set_headers(api_key)
            self.refresh_async()

    # ------- UI helpers -------
    def clear_list(self):
        for f in self.row_frames:
            f.destroy()
        self.row_frames.clear()
        self.row_vars.clear()

    def _assignee_from_rec(self, rec: Dict) -> str:
        return ((rec.get("assignee") or {}).get("email") or "").strip()

    def _apply_assignee_filter(self, rows: List[Dict]) -> List[Dict]:
        choice = self.assignee_filter_var.get()
        if choice == "All":
            return rows
        if choice == "Unassigned":
            return [r for r in rows if not self._assignee_from_rec(r)]
        return [r for r in rows if self._assignee_from_rec(r) == choice]

    def add_row(self, rec: Dict):
        rrn = rec.get("rrn", "")
        title = rec.get("title", "") or "(no title)"
        status = rec.get("status", "")
        priority = rec.get("priority", "")
        source = rec.get("source", "")
        created_time = parse_iso_to_local(rec.get("created_time", ""))
        assignee_email = self._assignee_from_rec(rec)
        link = console_link(rrn, self.region, self.org_id)

        # Card frame with border & padding for clear separation
        row = ctk.CTkFrame(self.scroll, border_width=1, corner_radius=10)
        row.pack(fill="x", padx=6, pady=6)

        var = ctk.BooleanVar(value=False)
        chk = ctk.CTkCheckBox(row, text="", variable=var, width=22)
        chk.grid(row=0, column=0, rowspan=3, padx=(8, 12), pady=8, sticky="nw")

        # line 1: title (bigger bold) + open button
        title_lbl = ctk.CTkLabel(row, text=title, anchor="w", font=self.font_title_bold)
        title_lbl.grid(row=0, column=1, sticky="w", padx=4, pady=(10, 2))

        if link:
            def open_link(url=link):
                webbrowser.open(url, new=2)
                self._log_status("Opened investigation in browser.")
            link_btn = ctk.CTkButton(row, text="Open", width=90, command=open_link)
            link_btn.grid(row=0, column=2, padx=(10, 10), pady=(10,2), sticky="ne")

        # line 2: meta (Created Time + Status/Priority/Source)
        meta = f"Created: {created_time}    Status: {status}    Priority: {priority}    Source: {source}"
        meta_lbl = ctk.CTkLabel(row, text=meta, anchor="w")
        meta_lbl.grid(row=1, column=1, sticky="w", padx=4, pady=(0,2), columnspan=2)

        # line 3: assignee (red bold 'EMPTY' if none)
        if assignee_email:
            assignee_lbl = ctk.CTkLabel(row, text=f"Assignee: {assignee_email}", anchor="w")
        else:
            assignee_lbl = ctk.CTkLabel(
                row,
                text="Assignee: EMPTY",
                anchor="w",
                font=self.font_bold_red,
                text_color="red"
            )
        assignee_lbl.grid(row=2, column=1, sticky="w", padx=4, pady=(0,10), columnspan=2)

        row.grid_columnconfigure(1, weight=1)

        self.row_vars.append(var)
        self.row_frames.append(row)

    def rebuild_list(self):
        self.clear_list()

        def key_fn(r):
            dt = r.get("created_time") or ""
            try:
                if dt.endswith("Z"):
                    d = datetime.fromisoformat(dt.replace("Z", "+00:00"))
                else:
                    d = datetime.fromisoformat(dt)
            except Exception:
                d = datetime(1970, 1, 1, tzinfo=timezone.utc)
            return d

        rows_sorted = sorted(self.rows, key=key_fn, reverse=not self.sort_oldest_first)
        rows_sorted = self._apply_assignee_filter(rows_sorted)

        for rec in rows_sorted:
            self.add_row(rec)

        direction = "Oldest → Newest" if self.sort_oldest_first else "Newest → Oldest"
        self._log_status(f"Loaded {len(rows_sorted)} (filtered) · {direction}")

    # ------- logging helper -------
    def _log_status(self, msg: str):
        line = f"[{now_str()}] {msg}\n"
        self.status_box.insert("end", line)
        self.status_box.see("end")

    # ------- async actions -------
    def refresh_async(self):
        """Refresh in background thread with progress dialog."""
        def worker():
            try:
                data = list_investigations()
                self.after(0, lambda: self._refresh_complete(data))
            except Exception as e:
                self.after(0, lambda: self._refresh_error(e))
        
        progress = ProgressDialog(self, "Refreshing Investigations")
        progress.update_message("Fetching investigations from API...")
        
        def start_work():
            thread = threading.Thread(target=worker, daemon=True)
            thread.start()
            self._check_thread_progress(progress, thread)
        
        self.after(100, start_work)
    
    def _check_thread_progress(self, progress_dialog, thread):
        """Check if thread is still alive, close dialog when done."""
        if thread.is_alive():
            self.after(100, lambda: self._check_thread_progress(progress_dialog, thread))
        else:
            progress_dialog.destroy()
    
    def _refresh_complete(self, data: List[Dict]):
        """Called when refresh completes successfully."""
        self.rows = data
        
        # Populate assignee filter options (All, Unassigned, unique emails)
        emails: Set[str] = set()
        for r in self.rows:
            e = self._assignee_from_rec(r)
            if e:
                emails.add(e)
        options = ["All", "Unassigned"] + sorted(emails, key=lambda x: x.lower())
        self.assignee_filter.configure(values=options)
        if self.assignee_filter_var.get() not in options:
            self.assignee_filter_var.set("All")
        
        self.rebuild_list()
        self.select_all_var.set(False)
        self._log_status(f"Refresh complete - {len(data)} total investigations loaded")
    
    def _refresh_error(self, error: Exception):
        """Called when refresh fails."""
        self._log_status(f"Refresh failed: {error}")
        messagebox.showerror("Error", str(error))

    def toggle_select_all(self):
        new_val = self.select_all_var.get()
        for v in self.row_vars:
            v.set(new_val)

    def on_sort_toggle(self):
        # Toggle: True -> Oldest→Newest, False -> Newest→Oldest
        self.sort_oldest_first = bool(self.sort_switch_var.get())
        self.sort_switch.configure(
            text="Sort by Created Time (Oldest → Newest)" if self.sort_oldest_first
                 else "Sort by Created Time (Newest → Oldest)"
        )
        self.rebuild_list()

    def _sorted_rows_current_view(self) -> List[Dict]:
        def key_fn(r):
            dt = r.get("created_time") or ""
            try:
                if dt.endswith("Z"):
                    d = datetime.fromisoformat(dt.replace("Z", "+00:00"))
                else:
                    d = datetime.fromisoformat(dt)
            except Exception:
                d = datetime(1970, 1, 1, tzinfo=timezone.utc)
            return d
        rows_sorted = sorted(self.rows, key=key_fn, reverse=not self.sort_oldest_first)
        return self._apply_assignee_filter(rows_sorted)

    def _selected_rows(self) -> List[Dict]:
        rows_sorted = self._sorted_rows_current_view()
        selected_rows: List[Dict] = []
        for i, v in enumerate(self.row_vars):
            if v.get() and i < len(rows_sorted):
                selected_rows.append(rows_sorted[i])
        return selected_rows

    def update_selected_async(self):
        """Update selected investigations in background thread."""
        rows = self._selected_rows()
        if not rows:
            self._log_status("Select at least one investigation.")
            return

        chosen_status = self.status_choice.get()
        chosen_dispo = self.dispo_choice.get() or None
        chosen_assignee_label = self.assignee_choice.get()
        comment = self.comment_box.get("1.0", "end").strip() or None

        # map label -> email
        assignees = get_assignees_from_settings(self.cfg)
        chosen_email = None
        for (n, e) in assignees:
            label = f"{n} <{e}>" if e else n
            if label == chosen_assignee_label:
                chosen_email = e or None
                break

        total = len(rows)
        results = {"successes": 0, "fails": []}
        
        progress = ProgressDialog(self, "Updating Investigations")
        
        def worker():
            for idx, rec in enumerate(rows, start=1):
                inv_rrn = rec.get("rrn") or ""
                inv_id = rec.get("id") or ""
                inv_key_v2 = inv_id or inv_rrn
                title = rec.get("title", "")
                
                # Update progress
                self.after(0, lambda i=idx, t=title: progress.update_detail(f"Processing {i}/{total}: {t[:50]}..."))
                
                try:
                    if chosen_status:
                        set_status(inv_key_v2, chosen_status)
                    if chosen_dispo:
                        set_disposition(inv_key_v2, chosen_dispo)
                    if chosen_email:
                        assign_user(inv_key_v2, chosen_email)
                    if comment:
                        target_rrn = inv_rrn or get_rrn(inv_id)
                        info = create_comment_v1(target_rrn, comment)
                        if not info["ok"]:
                            raise RuntimeError(f"Comment HTTP {info['status']} {info['text']}")
                        # Add to history on first successful comment (only once per bulk update)
                        if idx == 1 and results["successes"] == 0:
                            self.after(0, lambda c=comment: self._add_to_comment_history(c))
                    
                    results["successes"] += 1
                    self.after(0, lambda i=idx, t=total, ti=title: 
                              self._log_status(f"✓ Updated {i}/{t}: {ti}"))
                except Exception as e:
                    results["fails"].append(f"{title}: {e}")
                    self.after(0, lambda i=idx, t=total, ti=title, er=str(e): 
                              self._log_status(f"✗ Failed {i}/{t}: {ti} — {er}"))
            
            self.after(0, lambda: self._update_complete(results, total, progress))
        
        def start_work():
            thread = threading.Thread(target=worker, daemon=True)
            thread.start()
        
        self.after(100, start_work)
    
    def _update_complete(self, results: dict, total: int, progress_dialog):
        """Called when updates complete."""
        progress_dialog.destroy()
        
        successes = results["successes"]
        fails = results["fails"]
        
        if fails:
            self._log_status(f"Done. Updated {successes}/{total}; {len(fails)} failed.")
            messagebox.showwarning("Some updates failed", "\n".join(fails)[:4000])
        else:
            self._log_status(f"Done. Updated {successes}/{total} investigation(s).")
        
        self.refresh_async()

    def test_comment_popup(self):
        """Preview and send v1 /comments with the RRN as target."""
        rows = self._selected_rows()
        if not rows:
            messagebox.showinfo("Test Comment", "Select at least one investigation to test.")
            return

        rec = rows[0]
        title = rec.get("title", "")
        rrn = rec.get("rrn") or ""
        inv_id = rec.get("id") or ""
        comment_text = self.comment_box.get("1.0", "end").strip() or "Test comment from Troubleshoot button"

        if not rrn and inv_id:
            try:
                rrn = get_rrn(inv_id)
            except Exception:
                rrn = "(could not resolve RRN from ID)"

        win = ctk.CTkToplevel(self)
        win.title("Test Comment – Review & Send (/v1/comments)")
        win.geometry("820x560")

        ctk.CTkLabel(win, text=f"Title: {title}", anchor="w").pack(fill="x", padx=10, pady=(8,4))

        rrn_frame = ctk.CTkFrame(win); rrn_frame.pack(fill="x", padx=10, pady=(2,8))
        ctk.CTkLabel(rrn_frame, text="Target RRN (copyable):").pack(anchor="w", padx=6, pady=(4,2))
        rrn_box = ctk.CTkTextbox(rrn_frame, height=28); rrn_box.pack(fill="x", padx=6, pady=(0,4))
        rrn_box.insert("1.0", rrn or "(none)"); rrn_box.configure(state="disabled")

        ctk.CTkLabel(win, text="Endpoint:").pack(anchor="w", padx=10)
        url_box = ctk.CTkTextbox(win, height=40); url_box.pack(fill="x", padx=10, pady=(0,8))
        url_box.insert("1.0", URL_V1_CREATE_COMMENT); url_box.configure(state="disabled")

        ctk.CTkLabel(win, text="JSON Body:").pack(anchor="w", padx=10)
        body_box = ctk.CTkTextbox(win, height=90); body_box.pack(fill="x", padx=10, pady=(0,8))
        body_box.insert("1.0", str({"target": rrn, "body": comment_text}))
        body_box.configure(state="disabled")

        def send_now():
            try:
                info = create_comment_v1(rrn, comment_text)
                ok = info["ok"]; status = info["status"]; text = info["text"]
                if ok:
                    self._add_to_comment_history(comment_text)
                messagebox.showinfo("Result (/v1/comments)", f"Status: {status} ({'SUCCESS' if ok else 'FAILED'})\n\nResponse:\n{text[:4000]}")
                self._log_status(f"/v1/comments send: HTTP {status}")
            except Exception as e:
                messagebox.showerror("Result (/v1/comments)", str(e))

        ctk.CTkButton(win, text="Send Test Comment", command=send_now).pack(pady=(6, 12))

def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
