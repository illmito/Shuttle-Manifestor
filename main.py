import os
import json
import tkinter as tk
from tkinter import ttk, messagebox
import string
import re
import sys
from PIL import Image, ImageTk

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────

CONFIG_PATH = "config.json"

default_config = {
    "__instructions": [
        "=== Passenger Formatter Configuration ===",
        "",
        "This file controls how the app reads and formats data pasted from Excel.",
        "",
        "'headers': Maps internal field names to Excel column headers.",
        "'group_label': Heading shown before the formatted seat list. Supports {placeholders}.",
        "'output_format': Format for each passenger row. Supports {seat_number:02} and {placeholders}.",
        "",
        "All placeholder names must match keys in 'headers'. Case-sensitive.",
        "---- RECOMMENDED TO NOT CHANGE PROFILE SETTINGS ----"
    ],
    "profiles": {
        "Defence Travel": {
            "headers": {
                "TRN": "TRN",
                "Rank/Title": "Rank/Title",
                "First Name": "First Name",
                "Last Name": "Last Name",
                "Employee ID": "Employee ID",
                "Contact Details": "Contact Details",
                "Transport Pick Up Date": "Transport Pick Up Date",
                "Transport Pick Up Time": "Transport Pick Up Time",
                "Transport Pick Up Location": "Transport Pick Up Location",
                "Flight Date": "Flight Date",
                "Flight Time": "Flight Time",
                "Flight Route": "Flight Route",
                "No of Bags": "No of Bags",
                "Destination Address": "Destination Address",
                "Cost Centre": "Cost Centre",
                "GL Account": "GL Account",
                "WBS": "WBS",
                "Fund Code": "Fund Code",
                "ACMS Number": "ACMS Number"
            },
            "group_label": "{Transport Pick Up Location} > {Destination Address} @ {Transport Pick Up Time} on {Transport Pick Up Date}",
            "output_format": "Seat Number {seat_number:02} {First Name} {Last Name} {Contact Details} {Flight Route} @ {Flight Time} {No of Bags}x Bags"
        },
        "Toll": {
            "headers": {
                "First Name": "FIRST NAME",
                "Last Name": "SURNAME",
                "Contact Details": "MOBILE NUMBER",
                "Pickup Info": "DATE & TIME OF COLLECTION (BUS STOP NUMBER - BELOW)",
                "Transport Pick Up Location": "FROM LOCATION",
                "Destination Address": "TO LOCATION",
                "No of Bags": "BAGS REQ.",
                "Employee ID": "PMKEYS"
            },
            "group_label": "{Transport Pick Up Location} > {Destination Address} @ {Pickup Info}",
            "output_format": "Seat Number {seat_number:02} {First Name} {Last Name} {Contact Details} {No of Bags}x Bags"
        },
                        "Toll 2": {
            "headers": {
                "First Name": "FIRST NAME",
                "Last Name": "SURNAME",
                "Contact Details": "MOBILE NUMBER",
                "Pickup Info": "DATE & TIME OF COLLECTION",
                "Transport Pick Up Location": "FROM LOCATION",
                "Destination Address": "TO LOCATION",
                "No of Bags": "BAGS REQ.",
                "Employee ID": "PMKEYS"
            },
            "group_label": "{Transport Pick Up Location} > {Destination Address} @ {Pickup Info}",
            "output_format": "Seat Number {seat_number:02} {First Name} {Last Name} {Contact Details} {No of Bags}x Bags"
        }
    }
}

if not os.path.exists(CONFIG_PATH):
    with open(CONFIG_PATH, "w") as f:
        json.dump(default_config, f, indent=2)

with open(CONFIG_PATH, "r") as f:
    config = json.load(f)

profile_key = next(iter(config["profiles"]))
profile = config["profiles"][profile_key]


# ─────────────────────────────────────────────
# FORMATTING LOGIC
# ─────────────────────────────────────────────

def match_headers(actual_headers, expected_headers):
    matched = {}
    for i, col in enumerate(actual_headers):
        col_clean = col.strip()
        for key, expected in expected_headers.items():
            if col_clean == expected:
                matched[key] = i
    return matched


def format_pasted_data(raw_data, seat_start, profile):
    lines = raw_data.strip().split('\n')
    if len(lines) <= 1:
        return ["[!] Not enough data to format."]

    headers = lines[0].strip().split('\t')
    rows = lines[1:]
    header_map = match_headers(headers, profile["headers"])

    group_label_fields = [f for _, f, _, _ in string.Formatter().parse(profile["group_label"]) if f is not None]
    output_format_fields = [f for _, f, _, _ in string.Formatter().parse(profile["output_format"]) if f not in (None, "seat_number")]
    required_fields = set(group_label_fields + output_format_fields)

    if not required_fields.issubset(header_map.keys()):
        missing = required_fields - set(header_map.keys())
        return [f"[!] Missing required fields: {', '.join(missing)}"]

    results = []
    try:
        first_row = rows[0].split('\t')
        group_info = {}
        for key in group_label_fields:
            if key in header_map:
                raw_value = first_row[header_map[key]].strip()
                if key == "Pickup Info":
                    match = re.search(r"\d{2}/\d{2}/\d{3,4}-?\s*\d{1,2}:\d{2}\s*[ap]m", raw_value, re.IGNORECASE)
                    group_info[key] = match.group(0) if match else "[Invalid Pickup Info]"
                else:
                    group_info[key] = raw_value
        results.append(profile["group_label"].format(**group_info))
    except Exception as e:
        results.append(f"[!] Could not generate group label: {e}")

    seat_number = seat_start
    for row in rows:
        cols = row.strip().split('\t')
        try:
            row_data = {"seat_number": seat_number}
            for key in output_format_fields:
                if key in header_map and header_map[key] < len(cols):
                    raw_value = cols[header_map[key]].strip()
                    if key == "Pickup Info":
                        match = re.search(r"\d{2}/\d{2}/\d{3,4}-?\s*\d{1,2}:\d{2}\s*[ap]m", raw_value, re.IGNORECASE)
                        row_data[key] = match.group(0) if match else "[Invalid Pickup Info]"
                    else:
                        row_data[key] = raw_value
            results.append(profile["output_format"].format(**row_data))
            seat_number += 1
        except Exception as e:
            results.append(f"[!] Error processing row: {e}")
    return results


# ─────────────────────────────────────────────
# DUPLICATE DETECTION LOGIC
# ─────────────────────────────────────────────

def normalise_name(name):
    name = name.lower().strip()
    name = re.sub(r"[^a-z0-9\s]", "", name)
    name = re.sub(r"\s+", " ", name)
    return name


def normalise_phone(phone):
    return re.sub(r"\D", "", phone)


def extract_passenger_fields(line):
    result = {
        "seat": None, "first_name": "", "last_name": "",
        "contact": "", "flight": "", "flight_time": "",
        "bags": "", "raw": line.strip()
    }

    seat_match = re.search(r"Seat\s+Number\s+(\d+)", line, re.IGNORECASE)
    if seat_match:
        result["seat"] = int(seat_match.group(1))

    flight_match = re.search(r"([A-Z]{2}\s*\d{3,4})\s*@\s*(\d{1,2}:?\d{2})", line)
    if flight_match:
        result["flight"] = re.sub(r"\s+", "", flight_match.group(1)).upper()
        raw_time = flight_match.group(2)
        if ":" not in raw_time and len(raw_time) == 4:
            raw_time = raw_time[:2] + ":" + raw_time[2:]
        result["flight_time"] = raw_time

    bags_match = re.search(r"(\d+)\s*(?:x|KG'?s?x?)\s*Bags?", line, re.IGNORECASE)
    if bags_match:
        result["bags"] = bags_match.group(1)

    phone_line = re.sub(r"Seat\s+Number\s+\d+", "", line, flags=re.IGNORECASE)
    phone_line = re.sub(r"[A-Z]{2}\s*\d{3,4}\s*@\s*\d{1,2}:\d{2}", "", phone_line)
    phone_line = re.sub(r"\d+\s*(?:x|KG'?s?x?)\s*Bags?", "", phone_line, flags=re.IGNORECASE)
    for m in re.finditer(r"(\+?[\d][\d\s\-]{6,14}[\d])", phone_line):
        candidate = normalise_phone(m.group(1))
        if len(candidate) >= 8:
            result["contact"] = candidate
            break

    name_line = re.sub(r"^Seat\s+Number\s+\d+\s*", "", line, flags=re.IGNORECASE).strip()
    name_line = re.sub(r"[A-Z]{2}\s*\d{3,4}\s*@\s*\d{1,2}:\d{2}", "", name_line)
    name_line = re.sub(r"\d+\s*(?:x|KG'?s?x?)\s*Bags?", "", name_line, flags=re.IGNORECASE)
    name_line = re.split(r"\d", name_line)[0]
    name_parts = [p for p in name_line.strip().split() if p.isalpha()]

    if len(name_parts) >= 2:
        result["first_name"] = name_parts[0]
        result["last_name"] = name_parts[1]
    elif len(name_parts) == 1:
        result["first_name"] = name_parts[0]

    return result


def parse_manifest_lines(text):
    passengers = []
    for line in text.strip().splitlines():
        line = line.strip()
        if not line:
            continue
        if re.search(r"Seat\s+Number\s+\d+", line, re.IGNORECASE):
            passengers.append(extract_passenger_fields(line))
    return passengers


def compare_passengers(new_p, existing_p):
    new_first = normalise_name(new_p["first_name"])
    new_last  = normalise_name(new_p["last_name"])
    ex_first  = normalise_name(existing_p["first_name"])
    ex_last   = normalise_name(existing_p["last_name"])
    new_phone = normalise_phone(new_p["contact"])
    ex_phone  = normalise_phone(existing_p["contact"])
    new_flight = re.sub(r"\s+", "", new_p["flight"]).upper()
    ex_flight  = re.sub(r"\s+", "", existing_p["flight"]).upper()

    # Names are the primary duplicate signal.
    # A shared phone number is common for group bookings/admin contacts,
    # so phone_match should support a name match, not create a duplicate by itself.
    name_exact = (
        bool(new_first and ex_first) and
        new_first == ex_first and
        new_last == ex_last
    )

    name_partial = (
        # Same first name, one side is missing the surname: "JAMES" vs "JAMES REDIY"
        (new_first == ex_first and new_first != "" and (new_last == "" or ex_last == "")) or
        # Same first name, surname looks like an abbreviation/partial: "HOW" vs "HOWETT"
        (new_first == ex_first and new_first != "" and new_last != "" and ex_last != "" and
         (new_last in ex_last or ex_last in new_last))
    )

    phone_match        = bool(new_phone and ex_phone and new_phone == ex_phone)
    phone_missing_new  = not new_phone
    phone_missing_ex   = not ex_phone
    flight_match       = bool(new_flight and ex_flight and new_flight == ex_flight)
    flight_missing_new = not new_flight
    flight_missing_ex  = not ex_flight
    flight_differs     = bool(new_flight and ex_flight and new_flight != ex_flight)

    # Important fix:
    # Same phone number with a different name is not enough to be a duplicate.
    # This prevents shared contact numbers from creating false duplicate cards.
    if not name_exact and not name_partial:
        return None, None

    def build_merge_suggestion():
        seat = new_p["seat"] or 0
        return (
            f"Seat Number {seat:02d} {new_p['first_name'] or existing_p['first_name']} "
            f"{new_p['last_name'] or existing_p['last_name']} "
            f"{new_phone or ex_phone} {new_flight or ex_flight} @ "
            f"{new_p['flight_time'] or existing_p['flight_time']} "
            f"{new_p['bags'] or existing_p['bags']}x Bags"
        ).strip()

    details = {
        "new": new_p,
        "existing": existing_p,
        "name_exact": name_exact,
        "phone_match": phone_match,
        "flight_match": flight_match,
        "suggested_merge": None
    }

    # Merge only when the names already look like the same passenger.
    if (name_exact or name_partial) and phone_match and (flight_missing_new ^ flight_missing_ex):
        details["suggested_merge"] = build_merge_suggestion()
        return "merge", details

    if name_exact and (
        (flight_missing_new and not phone_missing_new and phone_missing_ex) or
        (flight_missing_ex and not phone_missing_ex and phone_missing_new)
    ):
        details["suggested_merge"] = build_merge_suggestion()
        return "merge", details

    if name_exact and phone_match and (flight_match or (flight_missing_new and flight_missing_ex)):
        return "exact", details

    if name_exact and flight_differs:
        return "possible", details

    if name_exact and (phone_missing_new or phone_missing_ex):
        return "possible", details

    if name_partial:
        return "possible", details

    return None, None

def find_duplicates(new_passengers, existing_passengers):
    raw = []
    for new_p in new_passengers:
        best_match = best_type = best_ex = None
        for ex_p in existing_passengers:
            match_type, details = compare_passengers(new_p, ex_p)
            if match_type == "exact":
                best_match, best_type, best_ex = details, "exact", ex_p["raw"]; break
            elif match_type == "merge" and best_type not in ("exact",):
                best_match, best_type, best_ex = details, "merge", ex_p["raw"]
            elif match_type == "possible" and best_type not in ("exact", "merge"):
                best_match, best_type, best_ex = details, "possible", ex_p["raw"]
        if best_type:
            raw.append((best_type, best_match, new_p, best_ex))

    groups = {}
    for match_type, details, new_p, ex_raw in raw:
        groups.setdefault(ex_raw, []).append((match_type, details, new_p))

    flat_results, group_results = [], []
    for ex_raw, entries in groups.items():
        if len(entries) == 1:
            flat_results.append((entries[0][0], entries[0][1]))
        else:
            types = [e[0] for e in entries]
            group_type = "exact" if "exact" in types else ("merge" if "merge" in types else "possible")

            def completeness(entry):
                p = entry[2]
                return sum([bool(p["first_name"]), bool(p["last_name"]),
                            bool(p["contact"]), bool(p["flight"]), bool(p["flight_time"])])

            entries_sorted = sorted(entries, key=completeness, reverse=True)
            group_results.append({
                "type": group_type,
                "existing": entries[0][1]["existing"],
                "entries": entries,
                "best": entries_sorted[0][2],
                "all_new": [e[2] for e in entries],
            })

    return flat_results, group_results


# ─────────────────────────────────────────────
# DUPLICATE CHECKER WINDOW
# ─────────────────────────────────────────────

def open_duplicate_checker(parent, current_output):
    win = tk.Toplevel(parent)
    win.title("Duplicate Checker")
    win.geometry("1300x860")
    win.minsize(900, 680)

    t           = get_theme()
    BG          = t["BG"]
    PANEL_BG    = t["PANEL_BG"]
    HEADER_BG   = t["HEADER_BG"]
    ACCENT      = t["ACCENT"]
    TEXT        = t["TEXT"]
    SUBTEXT     = t["SUBTEXT"]
    EXACT_BG    = t["EXACT_BG"]
    EXACT_FG    = t["EXACT_FG"]
    POSSIBLE_BG = t["POSSIBLE_BG"]
    POSSIBLE_FG = t["POSSIBLE_FG"]
    MERGE_BG    = t["MERGE_BG"]
    MERGE_FG    = t["MERGE_FG"]
    KEEP_FG     = t["KEEP_FG"]
    TAB_ACTIVE  = t["TAB_ACTIVE"]
    TAB_INACTIVE= t["TAB_INACTIVE"]
    LOCKED_BG   = t["LOCKED_BG"]
    LOCKED_FG   = t["LOCKED_FG"]

    win.configure(bg=BG)
    win.lift()
    win.focus_force()
    win.attributes("-topmost", True)
    win.after(200, lambda: win.attributes("-topmost", False))

    FONT_MONO  = ("Consolas", 10)
    FONT_UI    = ("Segoe UI", 10)
    FONT_HEAD  = ("Segoe UI", 11, "bold")
    FONT_TITLE = ("Segoe UI", 13, "bold")

    # ── State ──
    duplicate_results = []
    group_results     = []

    check_has_run = [False]
    active_tab = tk.StringVar(value="confirmed")

    # ── Top bar ──
    topbar = tk.Frame(win, bg=HEADER_BG, pady=8)
    topbar.pack(fill="x")
    tk.Label(topbar, text="⟳  Duplicate Checker", bg=HEADER_BG, fg="#ffffff",
             font=("Segoe UI", 14, "bold")).pack(side="left", padx=16)
    tk.Label(topbar, text="Cross-reference new entries against an existing manifest",
             bg=HEADER_BG, fg="#adb5bd", font=FONT_UI).pack(side="left", padx=4)

    body = tk.Frame(win, bg=BG)
    body.pack(fill="both", expand=True, padx=12, pady=(8, 0))
    body.columnconfigure(0, weight=1, uniform="main_cols")
    body.columnconfigure(1, weight=0)
    body.columnconfigure(2, weight=1, uniform="main_cols")
    body.rowconfigure(1, weight=1)
    body.rowconfigure(3, weight=3)

    def make_panel_label(parent, text, col, clear_cmd=None):
        f = tk.Frame(parent, bg=HEADER_BG, pady=6)
        f.grid(row=0, column=col, sticky="ew", padx=(0, 4) if col < 2 else 0, pady=(0, 4))
        tk.Label(f, text=text, bg=HEADER_BG, fg="#ffffff", font=FONT_HEAD).pack(side="left", padx=10)
        if clear_cmd:
            tk.Button(f, text="✕ Clear", command=clear_cmd,
                      bg=TAB_INACTIVE, fg="#ffffff",
                      font=("Segoe UI", 8, "bold"), relief="flat",
                      activebackground=HEADER_BG, activeforeground="#ffffff",
                      padx=8, pady=2, cursor="hand2").pack(side="right", padx=8)
        return f

    make_panel_label(body, "📋  New Entries", 0,
                     clear_cmd=lambda: new_text.delete("1.0", "end"))
    divider_top = tk.Frame(body, bg=ACCENT, width=2)
    divider_top.grid(row=0, column=1, rowspan=2, sticky="ns", padx=6)
    make_panel_label(body, "📁  Existing Manifest", 2,
                     clear_cmd=lambda: [exist_text.delete("1.0", "end"),
                                        exist_text.insert("1.0", "← Paste existing manifest here..."),
                                        exist_text.config(fg=SUBTEXT)])

    def make_text_panel(parent, col, prefill=""):
        frame = tk.Frame(parent, bg=PANEL_BG, bd=1, relief="solid",
                         highlightbackground=t["PANEL_BDR"], highlightthickness=1)
        frame.grid(row=1, column=col, sticky="nsew", padx=(0, 4) if col < 2 else 0)
        sb = tk.Scrollbar(frame, bg=PANEL_BG, troughcolor=BG)
        sb.pack(side="right", fill="y")
        txt = tk.Text(frame, bg=PANEL_BG, fg=TEXT, insertbackground=ACCENT,
                      font=FONT_MONO, wrap="none", yscrollcommand=sb.set,
                      relief="flat", padx=10, pady=10, selectbackground=ACCENT,
                      selectforeground="#fff")
        txt.pack(fill="both", expand=True)
        sb.config(command=txt.yview)
        if prefill:
            txt.insert("1.0", prefill)
        return txt

    new_text   = make_text_panel(body, 0, prefill=current_output.strip())
    exist_text = make_text_panel(body, 2)
    exist_text.insert("1.0", "← Paste existing manifest here...")
    exist_text.config(fg=SUBTEXT)

    def on_exist_focus_in(e):
        if exist_text.get("1.0", "end").strip() == "← Paste existing manifest here...":
            exist_text.delete("1.0", "end")
            exist_text.config(fg=TEXT)
    def on_exist_focus_out(e):
        if not exist_text.get("1.0", "end").strip():
            exist_text.insert("1.0", "← Paste existing manifest here...")
            exist_text.config(fg=SUBTEXT)
    exist_text.bind("<FocusIn>", on_exist_focus_in)
    exist_text.bind("<FocusOut>", on_exist_focus_out)

    # ── Action + tab bar ──
    action_tab_row = tk.Frame(body, bg=BG)
    action_tab_row.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8, 0))
    action_tab_row.columnconfigure(1, weight=1)

    tab_frame = tk.Frame(action_tab_row, bg=BG)
    tab_frame.grid(row=0, column=0, sticky="w")

    confirmed_tab_label = tk.StringVar(value="● Confirmed Duplicates (0)")
    review_tab_label    = tk.StringVar(value="◉ Needs Review (0 remaining)")

    tab_confirmed = tk.Button(tab_frame, textvariable=confirmed_tab_label,
                              bg=TAB_ACTIVE, fg="#ffffff",
                              font=("Segoe UI", 10, "bold"), relief="flat",
                              padx=16, pady=8, cursor="hand2",
                              command=lambda: switch_tab("confirmed"))
    tab_confirmed.pack(side="left", padx=(0, 2))

    tab_review = tk.Button(tab_frame, textvariable=review_tab_label,
                           bg=TAB_INACTIVE, fg="#ffffff",
                           font=("Segoe UI", 10, "bold"), relief="flat",
                           padx=16, pady=8, cursor="hand2",
                           command=lambda: switch_tab("review"))
    tab_review.pack(side="left")

    status_var = tk.StringVar(value="")

    btn_frame_right = tk.Frame(action_tab_row, bg=BG)
    btn_frame_right.grid(row=0, column=2, sticky="e")

    check_btn = tk.Button(btn_frame_right, text="🔍  Check for Duplicates",
                          command=lambda: run_check(), bg=ACCENT, fg="#fff",
                          font=("Segoe UI", 10, "bold"), relief="flat",
                          padx=14, pady=8, cursor="hand2")
    check_btn.pack(side="left", padx=(0, 6))

    generate_btn = tk.Button(btn_frame_right,
                             text="🔒  Generate Clean List — Run a check first",
                             bg=LOCKED_BG, fg=LOCKED_FG,
                             font=("Segoe UI", 10, "bold"), relief="flat",
                             padx=14, pady=8, state="disabled")
    generate_btn.pack(side="left")

    def switch_tab(name):
        active_tab.set(name)
        if name == "confirmed":
            tab_confirmed.config(bg=TAB_ACTIVE, fg="#ffffff")
            tab_review.config(bg=TAB_INACTIVE, fg="#ffffff")
        else:
            tab_confirmed.config(bg=TAB_INACTIVE, fg="#ffffff")
            tab_review.config(bg=TAB_ACTIVE, fg="#ffffff")
        rebuild_visible_cards()

    # ── Status bar ──
    status_bar = tk.Frame(win, bg=t["STATUS_BG"], pady=5)
    status_bar.pack(fill="x", side="bottom")
    tk.Frame(status_bar, bg=t.get("PANEL_BDR", HEADER_BG), height=1).pack(fill="x", side="top")
    tk.Label(status_bar, textvariable=status_var, bg=t["STATUS_BG"], fg=t["STATUS_FG"],
             font=("Segoe UI", 9), anchor="w").pack(side="left", padx=12)

    # ── Results area ──
    results_outer = tk.Frame(body, bg=BG)
    results_outer.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(4, 0))
    results_outer.columnconfigure(0, weight=1)
    results_outer.rowconfigure(0, weight=1)

    canvas = tk.Canvas(results_outer, bg=BG, highlightthickness=0, bd=0)
    canvas.grid(row=0, column=0, sticky="nsew")
    results_sb = tk.Scrollbar(results_outer, orient="vertical", command=canvas.yview)
    results_sb.grid(row=0, column=1, sticky="ns")
    canvas.configure(yscrollcommand=results_sb.set)

    results_inner = tk.Frame(canvas, bg=BG)
    canvas_window = canvas.create_window((0, 0), window=results_inner, anchor="nw")

    def on_results_configure(e):
        canvas.configure(scrollregion=canvas.bbox("all"))
    results_inner.bind("<Configure>", on_results_configure)

    def on_canvas_resize(e):
        canvas.itemconfig(canvas_window, width=e.width)
    canvas.bind("<Configure>", on_canvas_resize)

    def on_mousewheel(e):
        canvas.yview_scroll(int(-1*(e.delta/120)), "units")
    canvas.bind_all("<MouseWheel>", on_mousewheel)

    # ── Generate clean list ──
    def generate_clean_list():
        new_raw   = new_text.get("1.0", "end").strip()
        ex_raw    = exist_text.get("1.0", "end").strip()
        new_lines = new_raw.splitlines()
        ex_lines  = [l for l in ex_raw.splitlines()
                     if l.strip() and l.strip() != "← Paste existing manifest here..."]

        existing_to_drop   = set()
        new_to_drop        = set()
        merge_replacements = {}

        for match_type, details, dv in duplicate_results:
            decision     = dv.get()
            raw_new      = details["new"]["raw"]
            raw_existing = details["existing"]["raw"]
            if match_type == "exact":
                existing_to_drop.add(raw_existing)
            elif decision == "keep_new":
                existing_to_drop.add(raw_existing)
            elif decision in ("remove_new", "keep_existing"):
                new_to_drop.add(raw_new)
            elif decision == "keep_both":
                pass
            elif decision == "merge" and details.get("suggested_merge"):
                merge_replacements[raw_new] = details["suggested_merge"]
                existing_to_drop.add(raw_existing)
            else:
                existing_to_drop.add(raw_existing)

        for grp in group_results:
            decision     = grp["decision_var"].get()
            all_new      = grp["all_new"]
            best_new     = grp["best"]
            raw_existing = grp["existing"]["raw"]
            if decision in ("keep_best", "keep_new"):
                for np_ in all_new:
                    if np_["raw"] != best_new["raw"]:
                        new_to_drop.add(np_["raw"])
                existing_to_drop.add(raw_existing)
            elif decision == "keep_existing":
                for np_ in all_new:
                    new_to_drop.add(np_["raw"])
            elif decision == "remove_all":
                for np_ in all_new:
                    new_to_drop.add(np_["raw"])
                existing_to_drop.add(raw_existing)
            else:
                for np_ in all_new:
                    if np_["raw"] != best_new["raw"]:
                        new_to_drop.add(np_["raw"])
                existing_to_drop.add(raw_existing)

        output_lines = []
        for line in new_lines:
            stripped = line.strip()
            if stripped in new_to_drop:
                continue
            elif stripped in merge_replacements:
                output_lines.append(merge_replacements[stripped])
            else:
                output_lines.append(line)

        for line in ex_lines:
            stripped = line.strip()
            if stripped not in existing_to_drop:
                if re.search(r"Seat\s+Number\s+\d+", stripped, re.IGNORECASE):
                    output_lines.append(line)

        seat_num    = 1
        final_lines = []
        for line in output_lines:
            if re.search(r"Seat\s+Number\s+\d+", line, re.IGNORECASE):
                line = re.sub(r"(Seat\s+Number\s+)\d+",
                              lambda m: f"{m.group(1)}{seat_num:02d}", line, flags=re.IGNORECASE)
                seat_num += 1
            final_lines.append(line)

        removed_count = len(new_to_drop) + len(existing_to_drop)
        clean_output  = "\n".join(final_lines)

        out_win = tk.Toplevel(win)
        out_win.title("Clean Manifest Output")
        out_win.geometry("900x650")
        out_win.configure(bg=BG)

        header = tk.Frame(out_win, bg=HEADER_BG, pady=10)
        header.pack(fill="x")
        tk.Label(header, text="Clean Manifest — Ready to Copy",
                 bg=HEADER_BG, fg="#ffffff", font=FONT_TITLE).pack(side="left", padx=16)
        tk.Label(out_win,
                 text=f"{seat_num-1} passenger(s)  ·  {removed_count} resolved  ·  {len(merge_replacements)} merged",
                 bg=BG, fg=SUBTEXT, font=FONT_UI).pack(padx=12, pady=(8, 2), anchor="w")

        out_frame = tk.Frame(out_win, bg=PANEL_BG)
        out_frame.pack(fill="both", expand=True, padx=12, pady=8)
        out_sb = tk.Scrollbar(out_frame)
        out_sb.pack(side="right", fill="y")
        out_t = tk.Text(out_frame, bg=PANEL_BG, fg=TEXT, font=FONT_MONO,
                        wrap="none", yscrollcommand=out_sb.set,
                        relief="flat", padx=8, pady=8)
        out_t.pack(fill="both", expand=True)
        out_sb.config(command=out_t.yview)
        out_t.insert("1.0", clean_output)
        out_t.config(state="disabled")

        def copy_clean():
            out_win.clipboard_clear()
            out_win.clipboard_append(clean_output)
            out_win.update()
            messagebox.showinfo("Copied", "Clean manifest copied to clipboard.", parent=out_win)

        tk.Button(out_win, text="Copy to Clipboard", command=copy_clean,
                  bg=ACCENT, fg="#fff", font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=14, pady=6, cursor="hand2").pack(pady=(0, 12))

    # ── Generate lock ──
    def refresh_generate_lock():
        review_flat   = [(tp, d, dv) for tp, d, dv in duplicate_results if tp in ("possible", "merge")]
        review_groups = [g for g in group_results if g["type"] in ("possible", "merge")]
        remaining  = sum(1 for _, _, dv in review_flat   if dv.get() == "— undecided —")
        remaining += sum(1 for g in review_groups if g["decision_var"].get() == "— undecided —")

        if check_has_run[0] and remaining == 0:
            generate_btn.config(text="✓  Generate Clean List",
                                bg=KEEP_FG, fg="#000", state="normal",
                                cursor="hand2", command=generate_clean_list)
        else:
            generate_btn.config(
                text=f"🔒  Generate Clean List — {remaining} item(s) still need review" if remaining > 0
                     else "🔒  Generate Clean List — Run a check first",
                bg=LOCKED_BG, fg=LOCKED_FG, state="disabled", cursor="")
        review_tab_label.set(f"◉ Needs Review ({remaining} remaining)")

    # ── Card helpers ──
    def make_card(parent, card_bg, border_col):
        wrap = tk.Frame(parent, bg=BG)
        wrap.pack(fill="x", padx=8, pady=4)
        tk.Frame(wrap, bg=border_col, width=4).pack(side="left", fill="y")
        body = tk.Frame(wrap, bg=card_bg, pady=10, padx=14)
        body.pack(side="left", fill="both", expand=True)
        return body

    def make_entry_row(parent, label, line, bg):
        row = tk.Frame(parent, bg=bg)
        row.pack(fill="x", pady=1)
        tk.Label(row, text=label, bg=bg, fg=SUBTEXT,
                 font=("Segoe UI", 9, "bold"), width=10, anchor="w").pack(side="left")
        tk.Label(row, text=line, bg=bg, fg=TEXT, font=FONT_MONO, anchor="w").pack(side="left")

    def make_decision_btn(parent, label, value, dv, bg_c, fg_c):
        def on_click():
            dv.set(value)
            refresh_generate_lock()
        b = tk.Button(parent, text=label, command=on_click, bg=bg_c, fg=fg_c,
                      font=("Segoe UI", 9, "bold"), relief="flat",
                      activebackground=bg_c, activeforeground=fg_c,
                      padx=10, pady=3, cursor="hand2")
        b.pack(side="left", padx=(0, 6))
        return b

    # ── Card builder ──
    def rebuild_visible_cards():
        for w in results_inner.winfo_children():
            w.destroy()

        tab = active_tab.get()
        has_any = bool(duplicate_results or group_results)

        if not has_any:
            if check_has_run[0]:
                tk.Label(results_inner,
                         text="✓  No duplicates found — all entries are unique.",
                         bg=BG, fg=KEEP_FG,
                         font=("Segoe UI", 12, "bold")).pack(pady=16)
                tk.Label(results_inner,
                         text="Click Generate Clean List to produce the final output.",
                         bg=BG, fg=SUBTEXT, font=FONT_UI).pack()
            else:
                tk.Label(results_inner, text="Run a check to see results here.",
                         bg=BG, fg=SUBTEXT, font=FONT_UI).pack(pady=20)
            return

        flat_items = [(tp, d, dv) for tp, d, dv in duplicate_results
                      if (tab == "confirmed" and tp == "exact") or
                         (tab == "review"    and tp in ("possible", "merge"))]
        grp_items  = [g for g in group_results] if tab == "review" else []

        if not flat_items and not grp_items:
            msg = "No confirmed duplicates found." if tab == "confirmed" else "No items need review."
            tk.Label(results_inner, text=f"✓  {msg}",
                     bg=BG, fg=KEEP_FG, font=("Segoe UI", 11, "bold")).pack(pady=20)
            return

        for match_type, details, decision_var in flat_items:
            new_p = details["new"]
            ex_p  = details["existing"]

            if match_type == "exact":
                card_bg = EXACT_BG; badge_fg = EXACT_FG
                border_col = t["CARD_BORDER_EXACT"]; badge_text = "● CONFIRMED DUPLICATE"
            elif match_type == "merge":
                card_bg = MERGE_BG; badge_fg = MERGE_FG
                border_col = t["CARD_BORDER_MERGE"]; badge_text = "◈ MERGE SUGGESTED"
            else:
                card_bg = POSSIBLE_BG; badge_fg = POSSIBLE_FG
                border_col = t["CARD_BORDER_POSSIBLE"]; badge_text = "◉ POSSIBLE DUPLICATE"

            card = make_card(results_inner, card_bg, border_col)
            top_row = tk.Frame(card, bg=card_bg)
            top_row.pack(fill="x", pady=(0, 4))
            tk.Label(top_row, text=badge_text, bg=card_bg, fg=badge_fg,
                     font=("Segoe UI", 9, "bold")).pack(side="left")
            seat_str = f"  →  New entry Seat {new_p['seat']:02d}" if new_p['seat'] else ""
            tk.Label(top_row, text=seat_str, bg=card_bg, fg=TEXT, font=FONT_UI).pack(side="left")

            make_entry_row(card, "NEW:",      new_p["raw"], card_bg)
            make_entry_row(card, "EXISTING:", ex_p["raw"],  card_bg)

            if match_type == "merge" and details.get("suggested_merge"):
                mr = tk.Frame(card, bg=card_bg)
                mr.pack(fill="x", pady=(2, 2))
                tk.Label(mr, text="MERGE →", bg=card_bg, fg=MERGE_FG,
                         font=("Segoe UI", 9, "bold"), width=10, anchor="w").pack(side="left")
                tk.Label(mr, text=details["suggested_merge"], bg=card_bg,
                         fg=MERGE_FG, font=FONT_MONO, anchor="w").pack(side="left")

            btn_frame = tk.Frame(card, bg=card_bg)
            btn_frame.pack(fill="x", pady=(6, 0))

            if match_type == "exact":
                tk.Label(btn_frame, text="✓  Auto: Keep New  (existing dropped)",
                         bg=card_bg, fg=KEEP_FG,
                         font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 12))
                make_decision_btn(btn_frame, "Keep Both", "keep_both",  decision_var, TAB_INACTIVE, "#ffffff")
                make_decision_btn(btn_frame, "✕ Remove",  "remove_new", decision_var, EXACT_FG,     "#ffffff")
            elif match_type == "merge":
                make_decision_btn(btn_frame, "Keep New",      "keep_new",      decision_var, KEEP_FG,      "#ffffff")
                make_decision_btn(btn_frame, "Keep Existing", "keep_existing", decision_var, HEADER_BG,    "#ffffff")
                make_decision_btn(btn_frame, "Keep Both",     "keep_both",     decision_var, TAB_INACTIVE, "#ffffff")
                make_decision_btn(btn_frame, "Accept Merge",  "merge",         decision_var, MERGE_FG,     "#ffffff")
                make_decision_btn(btn_frame, "✕ Remove",      "remove_new",    decision_var, EXACT_FG,     "#ffffff")
            else:
                make_decision_btn(btn_frame, "Confirm Keep New", "keep_new",      decision_var, KEEP_FG,      "#ffffff")
                make_decision_btn(btn_frame, "Keep Existing",    "keep_existing", decision_var, HEADER_BG,    "#ffffff")
                make_decision_btn(btn_frame, "Keep Both",        "keep_both",     decision_var, TAB_INACTIVE, "#ffffff")
                make_decision_btn(btn_frame, "✕ Remove",         "remove_new",    decision_var, EXACT_FG,     "#ffffff")

            tk.Label(btn_frame, textvariable=decision_var,
                     bg=card_bg, fg=SUBTEXT,
                     font=("Segoe UI", 9, "italic")).pack(side="left", padx=8)

        for grp in grp_items:
            grp_type     = grp["type"]
            ex_p         = grp["existing"]
            all_new      = grp["all_new"]
            best_new     = grp["best"]
            decision_var = grp["decision_var"]

            card_bg    = POSSIBLE_BG if grp_type != "exact" else EXACT_BG
            badge_fg   = EXACT_FG if grp_type == "exact" else POSSIBLE_FG
            border_col = t["CARD_BORDER_EXACT"] if grp_type == "exact" else t["CARD_BORDER_POSSIBLE"]

            card = make_card(results_inner, card_bg, border_col)
            top_row = tk.Frame(card, bg=card_bg)
            top_row.pack(fill="x", pady=(0, 6))
            tk.Label(top_row,
                     text=f"⚠ MULTIPLE DUPLICATES — {len(all_new)} entries matched to same record",
                     bg=card_bg, fg=badge_fg, font=("Segoe UI", 9, "bold")).pack(side="left")

            for np_ in all_new:
                is_best = np_["raw"] == best_new["raw"]
                label   = "BEST →" if is_best else f"SEAT {np_['seat']:02d}:"
                color   = KEEP_FG  if is_best else TEXT
                row = tk.Frame(card, bg=card_bg)
                row.pack(fill="x", pady=1)
                tk.Label(row, text=label, bg=card_bg, fg=color,
                         font=("Segoe UI", 9, "bold"), width=10, anchor="w").pack(side="left")
                tk.Label(row, text=np_["raw"], bg=card_bg, fg=TEXT,
                         font=FONT_MONO, anchor="w").pack(side="left")

            make_entry_row(card, "EXISTING:", ex_p["raw"], card_bg)
            tk.Frame(card, bg=SUBTEXT, height=1).pack(fill="x", pady=(6, 4))

            btn_frame = tk.Frame(card, bg=card_bg)
            btn_frame.pack(fill="x", pady=(2, 0))
            make_decision_btn(btn_frame, "✓ Keep Best Only", "keep_best",    decision_var, KEEP_FG,   "#ffffff")
            make_decision_btn(btn_frame, "Keep Existing",    "keep_existing", decision_var, HEADER_BG, "#ffffff")
            make_decision_btn(btn_frame, "✕ Remove All",     "remove_all",    decision_var, EXACT_FG,  "#ffffff")
            tk.Label(btn_frame, textvariable=decision_var,
                     bg=card_bg, fg=SUBTEXT,
                     font=("Segoe UI", 9, "italic")).pack(side="left", padx=8)

    # ── Run check ──
    def run_check():
        nonlocal duplicate_results, group_results
        duplicate_results.clear()
        group_results.clear()
        check_has_run[0] = False

        new_raw = new_text.get("1.0", "end").strip()
        ex_raw  = exist_text.get("1.0", "end").strip()

        if not new_raw:
            status_var.set("⚠  No new entries to check.")
            return
        if not ex_raw or ex_raw == "← Paste existing manifest here...":
            status_var.set("⚠  Please paste the existing manifest on the right.")
            return

        new_passengers = parse_manifest_lines(new_raw)
        ex_passengers  = parse_manifest_lines(ex_raw)

        if not new_passengers:
            status_var.set("⚠  No seat entries found in new entries.")
            return
        if not ex_passengers:
            status_var.set("⚠  No seat entries found in existing manifest.")
            return

        flat_res, grp_res = find_duplicates(new_passengers, ex_passengers)

        for match_type, details in flat_res:
            initial = "keep_new" if match_type == "exact" else "— undecided —"
            dv = tk.StringVar(value=initial)
            duplicate_results.append((match_type, details, dv))

        group_results.clear()
        for grp in grp_res:
            initial = "keep_best" if grp["type"] == "exact" else "— undecided —"
            dv = tk.StringVar(value=initial)
            grp["decision_var"] = dv
            group_results.append(grp)

        exact_count    = sum(1 for tp, _, _ in duplicate_results if tp == "exact")
        possible_count = sum(1 for tp, _, _ in duplicate_results if tp == "possible")
        merge_count    = sum(1 for tp, _, _ in duplicate_results if tp == "merge")
        group_count    = len(group_results)

        confirmed_tab_label.set(f"● Confirmed Duplicates ({exact_count})")
        status_var.set(
            f"✓ Check complete — {exact_count} confirmed, {possible_count} possible, "
            f"{merge_count} merge suggested, {group_count} grouped  ·  {len(new_passengers)} new entries scanned."
        )

        check_has_run[0] = True

        if possible_count + merge_count + group_count > 0:
            switch_tab("review")
        else:
            switch_tab("confirmed")

        refresh_generate_lock()


# ─────────────────────────────────────────────
# THEME REGISTRY
# ─────────────────────────────────────────────

THEMES = {


        "Default": {
        "BG": "#f0f2f5", "PANEL_BG": "#ffffff", "HEADER_BG": "#00583D",
        "ACCENT": "#00A85A", "TEXT": "#111827", "SUBTEXT": "#6c757d",
        "KEEP_FG": "#00A859", "EXACT_FG": "#f44336", "POSSIBLE_FG": "#f59e0b",
        "MERGE_FG": "#04a9f5", "WARN_BG": "#fff8e1", "WARN_FG": "#f59e0b",
        "LABEL_FG": "#d9fbe8", "PANEL_BDR": "#e5e7eb", "GROUP_BG": "#073B2B",
        "SPIN_BTN": "#073B2B", "BTN_CLEAR": "#6c757d", "BTN_NAVY": "#007A4D",
        "BTN_OUTLN": "#00A859", "TAB_ACTIVE": "#073B2B", "TAB_INACTIVE": "#6c757d",
        "LOCKED_BG": "#e9ecef", "LOCKED_FG": "#adb5bd",
        "EXACT_BG": "#ffffff", "POSSIBLE_BG": "#ffffff", "MERGE_BG": "#ffffff",
        "STATUS_BG": "#e9f7ef", "STATUS_FG": "#00583D",
        "CARD_BORDER_EXACT": "#f44336", "CARD_BORDER_POSSIBLE": "#f59e0b",
        "CARD_BORDER_MERGE": "#04a9f5",
    },


    "Default Blue": {
        "BG": "#f0f2f5", "PANEL_BG": "#ffffff", "HEADER_BG": "#0b3155",
        "ACCENT": "#04a9f5", "TEXT": "#111827", "SUBTEXT": "#6c757d",
        "KEEP_FG": "#00c48c", "EXACT_FG": "#f44336", "POSSIBLE_FG": "#f59e0b",
        "MERGE_FG": "#04a9f5", "WARN_BG": "#fff8e1", "WARN_FG": "#f59e0b",
        "LABEL_FG": "#a8bbd4", "PANEL_BDR": "#e5e7eb", "GROUP_BG": "#1c2b4a",
        "SPIN_BTN": "#1c2b4a", "BTN_CLEAR": "#6c757d", "BTN_NAVY": "#3a4864",
        "BTN_OUTLN": "#04a9f5", "TAB_ACTIVE": "#1c2b4a", "TAB_INACTIVE": "#6c757d",
        "LOCKED_BG": "#e9ecef", "LOCKED_FG": "#adb5bd",
        "EXACT_BG": "#ffffff", "POSSIBLE_BG": "#ffffff", "MERGE_BG": "#ffffff",
        "STATUS_BG": "#e9ecef", "STATUS_FG": "#495057",
        "CARD_BORDER_EXACT": "#f44336", "CARD_BORDER_POSSIBLE": "#f59e0b",
        "CARD_BORDER_MERGE": "#04a9f5",
    },

    "Dark Mode": {
        "BG": "#1a1f2e", "PANEL_BG": "#232a3b", "HEADER_BG": "#2c3550",
        "ACCENT": "#4e9af1", "TEXT": "#dce8f5", "SUBTEXT": "#8a9bb5",
        "KEEP_FG": "#6bff9e", "EXACT_FG": "#ff6b6b", "POSSIBLE_FG": "#f1a84e",
        "MERGE_FG": "#4ec9f1", "WARN_BG": "#3b2e1a", "WARN_FG": "#f1a84e",
        "LABEL_FG": "#8a9bb5", "PANEL_BDR": "#2c3550", "GROUP_BG": "#2c3550",
        "SPIN_BTN": "#2c3550", "BTN_CLEAR": "#3a4a5c", "BTN_NAVY": "#3a4864",
        "BTN_OUTLN": "#4e9af1", "TAB_ACTIVE": "#2c3550", "TAB_INACTIVE": "#1a1f2e",
        "LOCKED_BG": "#2a2a2a", "LOCKED_FG": "#555555",
        "EXACT_BG": "#3b1f1f", "POSSIBLE_BG": "#3b2e1a", "MERGE_BG": "#1a2e3b",
        "STATUS_BG": "#1a1f2e", "STATUS_FG": "#8a9bb5",
        "CARD_BORDER_EXACT": "#ff6b6b", "CARD_BORDER_POSSIBLE": "#f1a84e",
        "CARD_BORDER_MERGE": "#4ec9f1",
    },
    "OLED Black": {
        "BG": "#000000", "PANEL_BG": "#0d0d0d", "HEADER_BG": "#000000",
        "ACCENT": "#00d4ff", "TEXT": "#ffffff", "SUBTEXT": "#555555",
        "KEEP_FG": "#00ff88", "EXACT_FG": "#ff3333", "POSSIBLE_FG": "#ffaa00",
        "MERGE_FG": "#00d4ff", "WARN_BG": "#1a1000", "WARN_FG": "#ffaa00",
        "LABEL_FG": "#555555", "PANEL_BDR": "#1a1a1a", "GROUP_BG": "#0d0d0d",
        "SPIN_BTN": "#0d0d0d", "BTN_CLEAR": "#333333", "BTN_NAVY": "#0d0d0d",
        "BTN_OUTLN": "#00d4ff", "TAB_ACTIVE": "#1a1a1a", "TAB_INACTIVE": "#000000",
        "LOCKED_BG": "#0d0d0d", "LOCKED_FG": "#2a2a2a",
        "EXACT_BG": "#1a0000", "POSSIBLE_BG": "#1a1000", "MERGE_BG": "#001a20",
        "STATUS_BG": "#000000", "STATUS_FG": "#555555",
        "CARD_BORDER_EXACT": "#ff3333", "CARD_BORDER_POSSIBLE": "#ffaa00",
        "CARD_BORDER_MERGE": "#00d4ff",
    },

    "Rose Pine Dawn": {
        "BG": "#faf4ed", "PANEL_BG": "#fffaf3", "HEADER_BG": "#575279",
        "ACCENT": "#d7827a", "TEXT": "#575279", "SUBTEXT": "#9893a5",
        "KEEP_FG": "#56949f", "EXACT_FG": "#b4637a", "POSSIBLE_FG": "#ea9d34",
        "MERGE_FG": "#d7827a", "WARN_BG": "#fdf0e0", "WARN_FG": "#ea9d34",
        "LABEL_FG": "#9893a5", "PANEL_BDR": "#dfdad9", "GROUP_BG": "#575279",
        "SPIN_BTN": "#575279", "BTN_CLEAR": "#9893a5", "BTN_NAVY": "#575279",
        "BTN_OUTLN": "#d7827a", "TAB_ACTIVE": "#575279", "TAB_INACTIVE": "#9893a5",
        "LOCKED_BG": "#f2e9e1", "LOCKED_FG": "#cecacd",
        "EXACT_BG": "#f7dde3", "POSSIBLE_BG": "#fdf0e0", "MERGE_BG": "#e8f4f5",
        "STATUS_BG": "#f2e9e1", "STATUS_FG": "#797593",
        "CARD_BORDER_EXACT": "#b4637a", "CARD_BORDER_POSSIBLE": "#ea9d34",
        "CARD_BORDER_MERGE": "#d7827a",
    },

    "Solarized Light": {
        "BG": "#fdf6e3", "PANEL_BG": "#eee8d5", "HEADER_BG": "#073642",
        "ACCENT": "#268bd2", "TEXT": "#073642", "SUBTEXT": "#839496",
        "KEEP_FG": "#2aa198", "EXACT_FG": "#dc322f", "POSSIBLE_FG": "#cb4b16",
        "MERGE_FG": "#268bd2", "WARN_BG": "#fdf0d5", "WARN_FG": "#cb4b16",
        "LABEL_FG": "#93a1a1", "PANEL_BDR": "#d3cbb8", "GROUP_BG": "#073642",
        "SPIN_BTN": "#073642", "BTN_CLEAR": "#839496", "BTN_NAVY": "#586e75",
        "BTN_OUTLN": "#268bd2", "TAB_ACTIVE": "#073642", "TAB_INACTIVE": "#839496",
        "LOCKED_BG": "#eee8d5", "LOCKED_FG": "#93a1a1",
        "EXACT_BG": "#fce0df", "POSSIBLE_BG": "#fdf0d5", "MERGE_BG": "#dff0f8",
        "STATUS_BG": "#eee8d5", "STATUS_FG": "#657b83",
        "CARD_BORDER_EXACT": "#dc322f", "CARD_BORDER_POSSIBLE": "#cb4b16",
        "CARD_BORDER_MERGE": "#268bd2",
    },

}

_active_theme = {"name": "Default"}

def get_theme():
    return THEMES[_active_theme["name"]]


# ─────────────────────────────────────────────
# MAIN GUI
# ─────────────────────────────────────────────

def launch_gui():
    root = tk.Tk()

    try:
        icon_img = Image.open(resource_path("windowlogo.png"))
        icon_img = icon_img.resize((32, 32), Image.LANCZOS)
        icon_photo = ImageTk.PhotoImage(icon_img)
        root.iconphoto(True, icon_photo)
        root._icon_photo = icon_photo
    except Exception as e:
        print(f"Window icon error: {e}")

    root.title("Shuttle Manifester - ")
    root.geometry("1100x720")
    root.minsize(900, 600)
    seat_start_var = tk.StringVar(value="1")
    profile_var    = tk.StringVar(value=profile_key)

    def copy_group_label_with_seats():
        try:
            label = group_label_cache.get().strip()
            output_lines = output_text.get("2.0", tk.END).strip().splitlines()
            last_seat = None
            for line in reversed(output_lines):
                match = re.search(r"Seat\s+Number\s+(\d+)", line)
                if match:
                    last_seat = int(match.group(1))
                    break
            if not label:
                raise ValueError("There is no valid Shuttle Service")
            if last_seat is None:
                raise ValueError("Could not detect any seat number in the output.")
            final_label = f"{label} @ {last_seat} seats"
            root.clipboard_clear()
            root.clipboard_append(final_label)
            root.update()
            messagebox.showinfo("Copied", f"Copied:\n{final_label}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not copy label: {e}")

    def on_profile_change(event=None):
        global profile
        profile = config["profiles"].get(profile_var.get(), profile)
        on_format()

    group_label_cache = tk.StringVar()

    def on_format(*_):
        try:
            seat_start = int(seat_start_var.get())
        except ValueError:
            output_text.config(state="normal")
            output_text.delete("1.0", tk.END)
            return
        raw_data = input_text.get("1.0", tk.END)
        output_text.config(state="normal")
        output_text.delete("1.0", tk.END)
        if raw_data.strip():
            formatted = format_pasted_data(raw_data, seat_start, profile)
            for i, line in enumerate(formatted):
                if i == 0 and not line.startswith("[!]"):
                    output_text.insert(tk.END, line + "\n", "group_label")
                    group_label_cache.set(line)
                elif line.startswith("[!]"):
                    output_text.insert(tk.END, line + "\n", "error")
                else:
                    start_index = output_text.index("end-1c")
                    output_text.insert(tk.END, line + "\n")
                    match = re.search(r"Seat\s+Number\s+(\d+)", line)
                    if match:
                        seat_num = int(match.group(1))
                        if seat_num >= 22:
                            s = f"{start_index}+{match.start(1)}c"
                            e = f"{start_index}+{match.end(1)}c"
                            output_text.tag_add("seat_warning", s, e)
        output_text.config(state="normal")

    def copy_output():
        root.clipboard_clear()
        root.clipboard_append(output_text.get("2.0", tk.END))
        root.update()

    def on_clear_all():
        seat_start_var.set("1")
        input_text.delete("1.0", tk.END)
        output_text.delete("1.0", tk.END)
        group_label_cache.set("")

    def open_dup_checker():
        current_output = output_text.get("1.0", tk.END).strip()
        open_duplicate_checker(root, current_output)

    def create_context_menu(widget):
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(label="Undo",       command=lambda: widget.event_generate("<<Undo>>"))
        menu.add_separator()
        menu.add_command(label="Cut",        command=lambda: widget.event_generate("<<Cut>>"))
        menu.add_command(label="Copy",       command=lambda: widget.event_generate("<<Copy>>"))
        menu.add_command(label="Paste",      command=lambda: widget.event_generate("<<Paste>>"))
        menu.add_separator()
        menu.add_command(label="Select All", command=lambda: widget.event_generate("<<SelectAll>>"))
        def show_menu(event):
            widget.focus()
            menu.tk_popup(event.x_root, event.y_root)
        widget.bind("<Button-3>", show_menu)
        widget.bind("<Button-2>", show_menu)
        widget.bind("<Control-z>", lambda e: widget.event_generate("<<Undo>>"))
        widget.bind("<Command-z>", lambda e: widget.event_generate("<<Undo>>"))

    current_theme = tk.StringVar(value="Default")

    # ── Load saved theme ──
    saved_theme = config.get("theme", "Default")
    if saved_theme in THEMES:
        _active_theme["name"] = saved_theme
        current_theme.set(saved_theme)

    t         = get_theme()
    BG        = t["BG"];  PANEL_BG  = t["PANEL_BG"]; HEADER_BG = t["HEADER_BG"]
    ACCENT    = t["ACCENT"]; TEXT   = t["TEXT"];      SUBTEXT   = t["SUBTEXT"]
    KEEP_FG   = t["KEEP_FG"]; EXACT_FG = t["EXACT_FG"]
    FONT_MONO = ("Consolas", 10); FONT_UI = ("Segoe UI", 10); FONT_HEAD = ("Segoe UI", 11, "bold")

    root.configure(bg=BG)
    root.columnconfigure(0, weight=1)

    # ── Menu bar ──
    menubar = tk.Menu(root, bg=HEADER_BG, fg="#ffffff", activebackground=ACCENT,
                      activeforeground="#fff", relief="flat")
    root.config(menu=menubar)
    config_menu = tk.Menu(menubar, tearoff=0, bg=HEADER_BG, fg="#ffffff",
                          activebackground=ACCENT, activeforeground="#fff")
    menubar.add_cascade(label="Menu", menu=config_menu)
    theme_menu = tk.Menu(config_menu, tearoff=0, bg=HEADER_BG, fg="#ffffff",
                         activebackground=ACCENT, activeforeground="#fff")
    config_menu.add_cascade(label="Theme", menu=theme_menu)
    config_menu.add_separator()
    config_menu.add_command(label="Profile Manager",  command=lambda: open_config_editor(root))
    config_menu.add_command(label="Clean Manifest",   command=lambda: open_clean_manifest_window(root))
    config_menu.add_command(label="Help",             command=lambda: open_help_window(root))

    # ── Top bar ──

    topbar = tk.Frame(root, bg=HEADER_BG, pady=10)
    topbar.grid(row=0, column=0, sticky="ew")

    try:
        left_img = Image.open(resource_path("logo.png"))
        left_img.thumbnail((200, 40), Image.LANCZOS)  # max 200wide x 40tall, keeps ratio
        left_photo = ImageTk.PhotoImage(left_img)
        left_logo_label = tk.Label(topbar, image=left_photo, bg=HEADER_BG)
        left_logo_label.image = left_photo
        left_logo_label.pack(side="left", padx=(12, 6))
    except Exception as e:
        print(f"Left logo error: {e}")

    tk.Label(topbar, text="Shuttle Manifester", bg=HEADER_BG, fg="#ffffff",
             font=("Segoe UI", 16, "bold")).pack(side="left", padx=(0, 16))
    tk.Label(topbar, text="Profile:", bg=HEADER_BG, fg="#a8bbd4",
             font=FONT_UI).pack(side="left", padx=(20, 4))
    profile_combo = ttk.Combobox(topbar, textvariable=profile_var, state="readonly", width=24)
    profile_combo["values"] = list(config["profiles"].keys())
    profile_combo.pack(side="left")
    profile_combo.bind("<<ComboboxSelected>>", on_profile_change)

    try:
        from datetime import datetime
        today = datetime.now()
        right_file = "ventiaxmas.png" if (today.month == 12 and today.day >= 15) else "ventialogo.png"
        right_img = Image.open(resource_path(right_file))
        right_img.thumbnail((200, 50), Image.LANCZOS)  # max 200wide x 40tall, keeps ratio
        right_photo = ImageTk.PhotoImage(right_img)
        right_logo_label = tk.Label(topbar, image=right_photo, bg=HEADER_BG)
        right_logo_label.image = right_photo
        right_logo_label.pack(side="right", padx=(0, 16))
    except Exception as e:
        print(f"Right logo error: {e}")

    

    # ── Input label ──
    lbl_frame = tk.Frame(root, bg=BG)
    lbl_frame.grid(row=1, column=0, sticky="w", padx=12, pady=(10, 2))
    tk.Label(lbl_frame, text="Paste data from Excel (with headers):",
             bg=BG, fg=SUBTEXT, font=FONT_UI).pack(side="left")

    # ── Input panel ──
    input_frame = tk.Frame(root, bg=PANEL_BG, bd=1, relief="solid",
                           highlightbackground=t["PANEL_BDR"], highlightthickness=1)
    input_frame.grid(row=2, column=0, sticky="nsew", padx=12)
    root.rowconfigure(2, weight=1)
    input_scroll = tk.Scrollbar(input_frame, bg=PANEL_BG, troughcolor=BG)
    input_scroll.pack(side="right", fill="y")
    input_text = tk.Text(input_frame, height=12, undo=True, wrap="none",
                         yscrollcommand=input_scroll.set, bg=PANEL_BG, fg=TEXT,
                         insertbackground=ACCENT, font=FONT_MONO, relief="flat",
                         padx=10, pady=10, selectbackground=ACCENT, selectforeground="#fff")
    input_text.pack(fill="both", expand=True)
    input_scroll.config(command=input_text.yview)
    create_context_menu(input_text)

    # ── Controls bar ──
    form = tk.Frame(root, bg=HEADER_BG, pady=8)
    form.grid(row=3, column=0, sticky="ew", padx=0, pady=(6, 0))
    form.columnconfigure(3, weight=1)
    tk.Label(form, text="Start Seat:", bg=HEADER_BG, fg="#a8bbd4",
             font=FONT_UI).grid(row=0, column=0, padx=(12, 4))
    seat_spin = tk.Spinbox(form, from_=1, to=999, textvariable=seat_start_var, width=6,
                           bg=PANEL_BG, fg=TEXT, buttonbackground=HEADER_BG,
                           insertbackground=ACCENT, relief="flat", font=FONT_UI)
    seat_spin.grid(row=0, column=1, padx=(0, 16))

    def make_btn(parent, text, cmd, color="#e9ecef", fg="#212529", outline=None):
        return tk.Button(parent, text=text, command=cmd, bg=color, fg=fg,
                         font=("Segoe UI", 9, "bold"), relief="flat",
                         activebackground=color, activeforeground=fg,
                         padx=14, pady=6, cursor="hand2",
                         highlightthickness=2 if outline else 0,
                         highlightbackground=outline or color,
                         highlightcolor=outline or color)

    make_btn(form, "Clear All",          on_clear_all,
             color="#6c757d", fg="#ffffff").grid(row=0, column=4, padx=4)
    make_btn(form, "Copy Manifest",      copy_output,
             color="#3a4864", fg="#ffffff", outline="#04a9f5").grid(row=0, column=5, padx=4)
    make_btn(form, "Copy Service Seats", copy_group_label_with_seats,
             color="#3a4864", fg="#ffffff", outline="#04a9f5").grid(row=0, column=6, padx=4)
    make_btn(form, "⟳  Check Duplicates", open_dup_checker,
             color=ACCENT, fg="#ffffff").grid(row=0, column=7, padx=(16, 16), sticky="e")

    # ── Output label ──
    out_lbl_frame = tk.Frame(root, bg=BG)
    out_lbl_frame.grid(row=4, column=0, sticky="w", padx=12, pady=(10, 2))
    tk.Label(out_lbl_frame, text="Formatted Manifest Output:",
             bg=BG, fg=SUBTEXT, font=FONT_UI).pack(side="left")

    # ── Output panel ──
    output_frame = tk.Frame(root, bg=PANEL_BG, bd=1, relief="solid",
                            highlightbackground=t["PANEL_BDR"], highlightthickness=1)
    output_frame.grid(row=5, column=0, sticky="nsew", padx=12, pady=(0, 12))
    root.rowconfigure(5, weight=1)
    output_scroll = tk.Scrollbar(output_frame, bg=PANEL_BG, troughcolor=BG)
    output_scroll.pack(side="right", fill="y")
    output_text = tk.Text(output_frame, height=15, undo=True, wrap="none",
                          yscrollcommand=output_scroll.set, bg=PANEL_BG, fg=TEXT,
                          insertbackground=ACCENT, font=FONT_MONO, relief="flat",
                          padx=10, pady=10, selectbackground=ACCENT, selectforeground="#fff")
    output_text.pack(fill="both", expand=True)
    output_scroll.config(command=output_text.yview)
    create_context_menu(output_text)

    output_text.tag_config("group_label", foreground="#ffffff", background=t["GROUP_BG"],
                           selectbackground=ACCENT, font=("Segoe UI", 13, "bold italic"),
                           spacing1=8, spacing3=8, lmargin1=10, lmargin2=10, justify="center")
    output_text.tag_config("error",        foreground=EXACT_FG, font=("Segoe UI", 10, "italic"))
    output_text.tag_config("seat_warning", background="#fff8e1", foreground="#f59e0b")

    # ── Config Manager ──
    def open_config_editor(parent):
        ct      = get_theme()
        C_BG    = ct["BG"];       C_PANEL  = ct["PANEL_BG"]; C_HEADER = ct["HEADER_BG"]
        C_ACCENT= ct["ACCENT"];   C_TEXT   = ct["TEXT"];     C_SUBTEXT= ct["SUBTEXT"]
        C_DANGER= ct["EXACT_FG"]; C_SUCCESS= ct["KEEP_FG"]
        C_MONO  = ("Consolas", 10); C_UI = ("Segoe UI", 10)
        C_HEAD  = ("Segoe UI", 11, "bold"); C_BOLD = ("Segoe UI", 10, "bold")

        import tkinter.filedialog as fd
        mgr = tk.Toplevel(parent)
        mgr.title("Profile Manager")
        mgr.geometry("1050x680")
        mgr.minsize(860, 520)
        mgr.configure(bg=C_BG)

        try:
            with open(CONFIG_PATH, "r") as f:
                working_config = json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Could not load config: {e}", parent=mgr)
            mgr.destroy(); return

        profiles = working_config.get("profiles", {})

        topbar_mgr = tk.Frame(mgr, bg=C_HEADER, pady=10)
        topbar_mgr.pack(fill="x")
        tk.Label(topbar_mgr, text="⚙  Profile Manager", bg=C_HEADER, fg="#ffffff",
                 font=("Segoe UI", 13, "bold")).pack(side="left", padx=16)
        tk.Label(topbar_mgr, text="Create, edit and manage manifest profiles",
                 bg=C_HEADER, fg="#a8bbd4", font=C_UI).pack(side="left", padx=4)

        body_mgr = tk.Frame(mgr, bg=C_BG)
        body_mgr.pack(fill="both", expand=True, padx=12, pady=10)
        body_mgr.columnconfigure(0, weight=0)
        body_mgr.columnconfigure(1, weight=1)
        body_mgr.rowconfigure(0, weight=1)

        left = tk.Frame(body_mgr, bg=C_PANEL, width=200)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.pack_propagate(False)
        lh = tk.Frame(left, bg=C_HEADER, pady=8)
        lh.pack(fill="x")
        tk.Label(lh, text="Profiles", bg=C_HEADER, fg="#ffffff", font=C_HEAD).pack(side="left", padx=10)
        profile_listbox = tk.Listbox(left, bg=C_PANEL, fg=C_TEXT, font=C_UI, relief="flat",
                                     selectbackground=C_ACCENT, selectforeground="#fff",
                                     activestyle="none", bd=0, highlightthickness=0)
        profile_listbox.pack(fill="both", expand=True, padx=4, pady=4)
        list_btn_bar = tk.Frame(left, bg=C_PANEL, pady=6)
        list_btn_bar.pack(fill="x")

        def lb(p, text, cmd, color=C_ACCENT, fg="#fff"):
            return tk.Button(p, text=text, command=cmd, bg=color, fg=fg,
                             font=("Segoe UI", 8, "bold"), relief="flat",
                             activebackground=color, activeforeground=fg,
                             padx=8, pady=4, cursor="hand2")

        lb(list_btn_bar, "+ New",    lambda: new_profile(),    C_ACCENT).pack(side="left", padx=(4, 2))
        lb(list_btn_bar, "✕ Delete", lambda: delete_profile(), C_DANGER).pack(side="left", padx=2)

        right = tk.Frame(body_mgr, bg=C_PANEL)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        rh = tk.Frame(right, bg=C_HEADER, pady=8)
        rh.pack(fill="x")
        rh_title = tk.Label(rh, text="Select a profile to edit",
                            bg=C_HEADER, fg="#ffffff", font=C_HEAD)
        rh_title.pack(side="left", padx=10)

        editor_body = tk.Frame(right, bg=C_PANEL)
        editor_body.pack(fill="both", expand=True, padx=14, pady=10)
        editor_body.columnconfigure(1, weight=1)

        def field_label(row, text):
            tk.Label(editor_body, text=text, bg=C_PANEL, fg=C_SUBTEXT,
                     font=C_BOLD, anchor="w").grid(row=row, column=0,
                     sticky="nw", padx=(0, 12), pady=(10, 2))

        field_label(0, "Profile Name:")
        profile_name_var = tk.StringVar()
        tk.Entry(editor_body, textvariable=profile_name_var, font=C_UI,
                 bg=C_PANEL, fg=C_TEXT, relief="solid", bd=1,
                 insertbackground=C_ACCENT).grid(row=0, column=1, sticky="ew", pady=(10, 2))

        field_label(1, "Group Label Format:")
        group_label_var = tk.StringVar()
        tk.Entry(editor_body, textvariable=group_label_var, font=C_UI,
                 bg=C_PANEL, fg=C_TEXT, relief="solid", bd=1,
                 insertbackground=C_ACCENT).grid(row=1, column=1, sticky="ew", pady=(10, 2))
        group_preview_var = tk.StringVar(value="")
        tk.Label(editor_body, textvariable=group_preview_var, bg=C_PANEL, fg=C_ACCENT,
                 font=("Consolas", 9), anchor="w", wraplength=500,
                 justify="left").grid(row=2, column=1, sticky="ew", pady=(0, 4))

        field_label(3, "Output Format:")
        output_format_var = tk.StringVar()
        tk.Entry(editor_body, textvariable=output_format_var, font=C_UI,
                 bg=C_PANEL, fg=C_TEXT, relief="solid", bd=1,
                 insertbackground=C_ACCENT).grid(row=3, column=1, sticky="ew", pady=(10, 2))
        output_preview_var = tk.StringVar(value="")
        tk.Label(editor_body, textvariable=output_preview_var, bg=C_PANEL, fg=C_ACCENT,
                 font=("Consolas", 9), anchor="w", wraplength=500,
                 justify="left").grid(row=4, column=1, sticky="ew", pady=(0, 8))

        header_rows = []

        def update_previews(*_):
            sample = {"seat_number": 1}
            for iv, ev, _ in header_rows:
                key = iv.get().strip()
                if not key: continue
                kl = key.lower()
                if "first" in kl and "name" in kl:   sample[key] = "John"
                elif "last" in kl and "name" in kl:  sample[key] = "Smith"
                elif "name" in kl:                   sample[key] = "John Smith"
                elif any(x in kl for x in ("phone", "mobile", "contact")): sample[key] = "0400000000"
                elif "flight" in kl and "time" in kl: sample[key] = "06:30"
                elif "flight" in kl or "route" in kl: sample[key] = "QF401"
                elif "bag" in kl:                    sample[key] = "2"
                elif "date" in kl:                   sample[key] = "01/01/2025"
                elif "time" in kl:                   sample[key] = "05:00"
                elif any(x in kl for x in ("location","pickup","from")): sample[key] = "BASE"
                elif any(x in kl for x in ("destination","address","to")): sample[key] = "AIRPORT"
                else: sample[key] = f"[{key}]"
            for var, prev_var in [(group_label_var, group_preview_var),
                                  (output_format_var, output_preview_var)]:
                try:
                    prev_var.set("Preview: " + var.get().format(**sample))
                except KeyError as e:
                    prev_var.set(f"Preview: field {e} not in mappings yet")
                except Exception as e:
                    prev_var.set(f"Preview error: {e}")

        group_label_var.trace_add("write", update_previews)
        output_format_var.trace_add("write", update_previews)

        field_label(5, "Header Mappings:")
        headers_frame = tk.Frame(editor_body, bg=C_PANEL)
        headers_frame.grid(row=5, column=1, sticky="nsew", pady=(10, 4))
        editor_body.rowconfigure(5, weight=1)

        hdr_head = tk.Frame(headers_frame, bg=C_HEADER)
        hdr_head.pack(fill="x")
        tk.Label(hdr_head, text="Internal Field Name", bg=C_HEADER, fg="#ffffff",
                 font=C_BOLD, width=24, anchor="w").pack(side="left", padx=6, pady=4)
        tk.Label(hdr_head, text="Excel Column Header", bg=C_HEADER, fg="#ffffff",
                 font=C_BOLD, anchor="w").pack(side="left", padx=6, pady=4, fill="x", expand=True)

        hdr_canvas = tk.Canvas(headers_frame, bg=C_PANEL, highlightthickness=0, height=160)
        hdr_sb = tk.Scrollbar(headers_frame, orient="vertical", command=hdr_canvas.yview)
        hdr_canvas.configure(yscrollcommand=hdr_sb.set)
        hdr_sb.pack(side="right", fill="y")
        hdr_canvas.pack(fill="both", expand=True)
        hdr_inner = tk.Frame(hdr_canvas, bg=C_PANEL)
        hdr_win = hdr_canvas.create_window((0, 0), window=hdr_inner, anchor="nw")
        hdr_inner.bind("<Configure>", lambda e: hdr_canvas.configure(scrollregion=hdr_canvas.bbox("all")))
        hdr_canvas.bind("<Configure>", lambda e: hdr_canvas.itemconfig(hdr_win, width=e.width))

        def add_header_row(internal="", excel=""):
            rf = tk.Frame(hdr_inner, bg=C_PANEL)
            rf.pack(fill="x", pady=1)
            iv = tk.StringVar(value=internal)
            ev = tk.StringVar(value=excel)
            e1 = tk.Entry(rf, textvariable=iv, font=C_UI, width=24,
                          bg=C_PANEL, fg=C_TEXT, relief="solid", bd=1, insertbackground=C_ACCENT)
            e1.pack(side="left", padx=(0, 4))
            e1.bind("<KeyRelease>", lambda e: update_previews())
            e2 = tk.Entry(rf, textvariable=ev, font=C_UI, bg=C_PANEL, fg=C_TEXT,
                          relief="solid", bd=1, insertbackground=C_ACCENT)
            e2.pack(side="left", fill="x", expand=True, padx=(0, 4))
            def remove_row():
                header_rows[:] = [(i, e, f) for i, e, f in header_rows if f is not rf]
                rf.destroy()
            tk.Button(rf, text="✕", command=remove_row, bg=C_DANGER, fg="#fff",
                      font=("Segoe UI", 8, "bold"), relief="flat",
                      padx=6, pady=2, cursor="hand2").pack(side="left")
            header_rows.append((iv, ev, rf))

        tk.Button(headers_frame, text="+ Add Field", command=lambda: add_header_row(),
                  bg=C_ACCENT, fg="#fff", font=C_BOLD, relief="flat",
                  padx=10, pady=4, cursor="hand2").pack(anchor="w", pady=(4, 0))

        def load_profile(name):
            rh_title.config(text=f"Editing: {name}")
            profile_name_var.set(name)
            p = profiles[name]
            group_label_var.set(p.get("group_label", ""))
            output_format_var.set(p.get("output_format", ""))
            for _, _, rf in header_rows: rf.destroy()
            header_rows.clear()
            for k, v in p.get("headers", {}).items():
                add_header_row(k, v)
            update_previews()

        def on_listbox_select(e):
            sel = profile_listbox.curselection()
            if sel: load_profile(profile_listbox.get(sel[0]))

        profile_listbox.bind("<<ListboxSelect>>", on_listbox_select)

        def refresh_listbox(select=None):
            profile_listbox.delete(0, "end")
            for name in profiles: profile_listbox.insert("end", name)
            if select and select in profiles:
                idx = list(profiles.keys()).index(select)
                profile_listbox.selection_set(idx)
                profile_listbox.activate(idx)

        def save_profile():
            original_name = profile_listbox.get(profile_listbox.curselection()[0]) \
                if profile_listbox.curselection() else None
            new_name = profile_name_var.get().strip()
            if not new_name:
                messagebox.showwarning("Missing Name", "Please enter a profile name.", parent=mgr); return
            headers = {iv.get().strip(): ev.get().strip()
                       for iv, ev, _ in header_rows if iv.get().strip() and ev.get().strip()}
            if not headers:
                messagebox.showwarning("No Headers", "Please add at least one header mapping.", parent=mgr); return
            if original_name and original_name != new_name:
                del profiles[original_name]
            profiles[new_name] = {"headers": headers,
                                  "group_label":   group_label_var.get().strip(),
                                  "output_format": output_format_var.get().strip()}
            working_config["profiles"] = profiles
            with open(CONFIG_PATH, "w") as f: json.dump(working_config, f, indent=2)
            refresh_listbox(select=new_name)
            messagebox.showinfo("Saved", f"Profile '{new_name}' saved. Restart to apply changes.", parent=mgr)

        def new_profile():
            for _, _, rf in header_rows: rf.destroy()
            header_rows.clear()
            profile_name_var.set("New Profile")
            group_label_var.set(""); output_format_var.set("")
            rh_title.config(text="New Profile")
            profile_listbox.selection_clear(0, "end")
            add_header_row("First Name", "First Name")
            add_header_row("Last Name", "Last Name")
            add_header_row("Contact Details", "Contact Details")

        def delete_profile():
            sel = profile_listbox.curselection()
            if not sel:
                messagebox.showwarning("No Selection", "Select a profile to delete.", parent=mgr); return
            name = profile_listbox.get(sel[0])
            if len(profiles) <= 1:
                messagebox.showwarning("Cannot Delete", "You must keep at least one profile.", parent=mgr); return
            if messagebox.askyesno("Delete Profile", f"Delete '{name}'? This cannot be undone.", parent=mgr):
                del profiles[name]
                working_config["profiles"] = profiles
                with open(CONFIG_PATH, "w") as f: json.dump(working_config, f, indent=2)
                refresh_listbox()
                rh_title.config(text="Select a profile to edit")
                profile_name_var.set(""); group_label_var.set(""); output_format_var.set("")
                for _, _, rf in header_rows: rf.destroy()
                header_rows.clear()

        def export_config():
            import tkinter.filedialog as fd2
            path = fd2.asksaveasfilename(parent=mgr, title="Export Config",
                                         defaultextension=".json",
                                         filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                                         initialfile="shuttle_config.json")
            if not path: return
            with open(path, "w") as f: json.dump(working_config, f, indent=2)
            messagebox.showinfo("Exported", "Config exported to:\n" + path, parent=mgr)

        def import_config():
            import tkinter.filedialog as fd2
            path = fd2.askopenfilename(parent=mgr, title="Import Config",
                                       filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
            if not path: return
            try:
                with open(path, "r") as f: imported = json.load(f)
                imported_profiles = imported.get("profiles", {})
                if not imported_profiles:
                    messagebox.showwarning("No Profiles", "No profiles found in this file.", parent=mgr); return
                choice = messagebox.askyesnocancel(
                    "Import Profiles",
                    f"Found {len(imported_profiles)} profile(s).\n\n"
                    "Yes = Merge with existing\nNo = Replace all\nCancel = Abort",
                    parent=mgr)
                if choice is None: return
                if choice: profiles.update(imported_profiles)
                else: profiles.clear(); profiles.update(imported_profiles)
                working_config["profiles"] = profiles
                with open(CONFIG_PATH, "w") as f: json.dump(working_config, f, indent=2)
                refresh_listbox()
                action = "Merged" if choice else "Replaced"
                messagebox.showinfo("Imported",
                    f"{action} with {len(imported_profiles)} profile(s). Restart to apply changes.",
                    parent=mgr)
            except Exception as e:
                messagebox.showerror("Import Error", f"Could not import config: {e}", parent=mgr)

        def reset_defaults():
            if not messagebox.askyesno("Reset to Defaults",
                "This will replace ALL profiles with the default configuration. Are you sure?",
                parent=mgr): return
            with open(CONFIG_PATH, "w") as f: json.dump(default_config, f, indent=2)
            messagebox.showinfo("Reset", "Config reset to defaults. Restart to apply changes.", parent=mgr)
            mgr.destroy()

        action_bar = tk.Frame(mgr, bg=C_HEADER, pady=8)
        action_bar.pack(fill="x", side="bottom")

        def action_btn(text, cmd, color=C_ACCENT, fg="#fff"):
            tk.Button(action_bar, text=text, command=cmd, bg=color, fg=fg,
                      font=C_BOLD, relief="flat", activebackground=color, activeforeground=fg,
                      padx=14, pady=6, cursor="hand2").pack(side="left", padx=6)

        action_btn("💾  Save Profile",   save_profile,   C_SUCCESS)
        action_btn("📤  Export Config",  export_config,  "#1c2b4a")
        action_btn("📥  Import Config",  import_config,  "#1c2b4a")
        action_btn("↺  Reset Defaults", reset_defaults, C_DANGER)

        refresh_listbox()
        if profiles:
            profile_listbox.selection_set(0)
            load_profile(list(profiles.keys())[0])

    def trigger_on_paste(event):  root.after(10, on_format)
    def watch_input_changes(event): root.after(10, on_format)
    def watch_seat_change(*args): on_format()

    input_text.bind("<Control-v>", trigger_on_paste)
    input_text.bind("<Command-v>", trigger_on_paste)
    input_text.bind("<<Paste>>",   trigger_on_paste)
    input_text.bind("<KeyRelease>", watch_input_changes)
    seat_start_var.trace_add("write", watch_seat_change)

    # ── Apply Theme ──
    def apply_theme(theme_name):
        current_theme.set(theme_name)
        _active_theme["name"] = theme_name
        th = THEMES[theme_name]
        root.configure(bg=th["BG"])
        topbar.configure(bg=th["HEADER_BG"])
        lbl_frame.configure(bg=th["BG"])
        out_lbl_frame.configure(bg=th["BG"])
        form.configure(bg=th["HEADER_BG"])
        input_frame.configure(bg=th["PANEL_BG"], highlightbackground=th["PANEL_BDR"])
        output_frame.configure(bg=th["PANEL_BG"], highlightbackground=th["PANEL_BDR"])
        for w in topbar.winfo_children():
            if isinstance(w, tk.Label):
                w.configure(bg=th["HEADER_BG"],
                            fg="#ffffff" if "Shuttle" in str(w.cget("text")) else th["LABEL_FG"])
        for w in lbl_frame.winfo_children():
            if isinstance(w, tk.Label): w.configure(bg=th["BG"], fg=th["SUBTEXT"])
        for w in out_lbl_frame.winfo_children():
            if isinstance(w, tk.Label): w.configure(bg=th["BG"], fg=th["SUBTEXT"])
        for w in form.winfo_children():
            if isinstance(w, tk.Label): w.configure(bg=th["HEADER_BG"], fg=th["LABEL_FG"])
        seat_spin.configure(bg=th["PANEL_BG"], fg=th["TEXT"],
                            buttonbackground=th["SPIN_BTN"], insertbackground=th["ACCENT"])
        input_text.configure(bg=th["PANEL_BG"], fg=th["TEXT"],
                             insertbackground=th["ACCENT"], selectbackground=th["ACCENT"])
        input_scroll.configure(bg=th["PANEL_BG"], troughcolor=th["BG"])
        output_text.configure(bg=th["PANEL_BG"], fg=th["TEXT"],
                              insertbackground=th["ACCENT"], selectbackground=th["ACCENT"])
        output_scroll.configure(bg=th["PANEL_BG"], troughcolor=th["BG"])
        output_text.tag_config("group_label", foreground="#ffffff", background=th["GROUP_BG"],
                               selectbackground=th["ACCENT"])
        output_text.tag_config("error", foreground=th["EXACT_FG"])
        output_text.tag_config("seat_warning", background=th["WARN_BG"], foreground=th["WARN_FG"])
        btn_configs = {
            4: (th["BTN_CLEAR"], "#ffffff", None),
            5: (th["BTN_NAVY"],  "#ffffff", th["BTN_OUTLN"]),
            6: (th["BTN_NAVY"],  "#ffffff", th["BTN_OUTLN"]),
            7: (th["ACCENT"],    "#ffffff", None),
        }
        for w in form.winfo_children():
            if isinstance(w, tk.Button):
                col = w.grid_info().get("column")
                if col in btn_configs:
                    bg, fg, outline = btn_configs[col]
                    w.configure(bg=bg, fg=fg, activebackground=bg, activeforeground=fg,
                                highlightbackground=outline or bg,
                                highlightcolor=outline or bg,
                                highlightthickness=2 if outline else 0)
        menubar.configure(bg=th["HEADER_BG"], fg="#ffffff", activebackground=th["ACCENT"])
        config_menu.configure(bg=th["HEADER_BG"], fg="#ffffff", activebackground=th["ACCENT"])
        theme_menu.configure(bg=th["HEADER_BG"], fg="#ffffff", activebackground=th["ACCENT"])

        # ── Save theme to config ──
        try:
            with open(CONFIG_PATH, "r") as f:
                cfg = json.load(f)
            cfg["theme"] = theme_name
            with open(CONFIG_PATH, "w") as f:
                json.dump(cfg, f, indent=2)
        except Exception:
            pass

    for theme_name in THEMES:
        theme_menu.add_radiobutton(label=theme_name, variable=current_theme, value=theme_name,
                                   command=lambda n=theme_name: apply_theme(n))

    # ── Apply saved theme on startup ──
    if saved_theme != "Default":
        apply_theme(saved_theme)

    input_text.focus_set()
    root.mainloop()


# ─────────────────────────────────────────────
# CLEAN MANIFEST WINDOW
# ─────────────────────────────────────────────

def scan_manifest_for_duplicates(passengers):
    """
    Compare every passenger against every other in the same list.
    Returns a list of clusters — each cluster is a list of indices
    that all match each other (same person entered multiple times).
    Only reports each pair once and never self-matches.
    """
    reported = set()
    adj = {}  # index -> set of connected indices

    for i, p_a in enumerate(passengers):
        for j, p_b in enumerate(passengers):
            if j <= i:
                continue
            pair = (i, j)
            if pair in reported:
                continue
            match_type, details = compare_passengers(p_a, p_b)
            if not match_type:
                continue
            reported.add(pair)
            adj.setdefault(i, set()).add(j)
            adj.setdefault(j, set()).add(i)
            # Store best match type and details on the edge
            edge_key = (min(i,j), max(i,j))
            if not hasattr(scan_manifest_for_duplicates, '_edges'):
                scan_manifest_for_duplicates._edges = {}
            existing = scan_manifest_for_duplicates._edges.get(edge_key)
            priority = {"exact": 3, "merge": 2, "possible": 1}
            if not existing or priority.get(match_type, 0) > priority.get(existing[0], 0):
                scan_manifest_for_duplicates._edges[edge_key] = (match_type, details)

    # Find connected components
    visited = set()
    clusters = []
    for start in range(len(passengers)):
        if start in visited or start not in adj:
            continue
        cluster = []
        stack = [start]
        while stack:
            node = stack.pop()
            if node in visited:
                continue
            visited.add(node)
            cluster.append(node)
            for nb in adj.get(node, []):
                if nb not in visited:
                    stack.append(nb)
        if len(cluster) >= 2:
            clusters.append(sorted(cluster))

    # Build result: each cluster becomes one item
    results = []
    for cluster in clusters:
        # Determine overall match type for cluster
        cluster_type = "possible"
        for i in range(len(cluster)):
            for j in range(i+1, len(cluster)):
                edge_key = (cluster[i], cluster[j])
                edge = scan_manifest_for_duplicates._edges.get(edge_key)
                if edge:
                    mt = edge[0]
                    if mt == "exact":
                        cluster_type = "exact"
                    elif mt == "merge" and cluster_type != "exact":
                        cluster_type = "merge"

        # Get suggested merge if any edge has one
        suggested_merge = None
        for i in range(len(cluster)):
            for j in range(i+1, len(cluster)):
                edge_key = (cluster[i], cluster[j])
                edge = scan_manifest_for_duplicates._edges.get(edge_key)
                if edge and edge[1].get("suggested_merge"):
                    suggested_merge = edge[1]["suggested_merge"]
                    break

        results.append({
            "type":           cluster_type,
            "indices":        cluster,
            "passengers":     [passengers[i] for i in cluster],
            "suggested_merge": suggested_merge,
        })

    # Clean up class-level edge store
    if hasattr(scan_manifest_for_duplicates, '_edges'):
        del scan_manifest_for_duplicates._edges

    return results


def open_clean_manifest_window(parent):
    win = tk.Toplevel(parent)
    win.title("Clean Manifest")
    win.geometry("1200x820")
    win.minsize(900, 650)

    t            = get_theme()
    BG           = t["BG"];          PANEL_BG   = t["PANEL_BG"]
    HEADER_BG    = t["HEADER_BG"];   ACCENT     = t["ACCENT"]
    TEXT         = t["TEXT"];        SUBTEXT    = t["SUBTEXT"]
    EXACT_FG     = t["EXACT_FG"];    POSSIBLE_FG= t["POSSIBLE_FG"]
    MERGE_FG     = t["MERGE_FG"];    KEEP_FG    = t["KEEP_FG"]
    TAB_ACTIVE   = t["TAB_ACTIVE"];  TAB_INACTIVE=t["TAB_INACTIVE"]
    LOCKED_BG    = t["LOCKED_BG"];   LOCKED_FG  = t["LOCKED_FG"]
    EXACT_BG     = t["EXACT_BG"];    POSSIBLE_BG= t["POSSIBLE_BG"]
    MERGE_BG     = t["MERGE_BG"]

    win.configure(bg=BG)
    win.lift(); win.focus_force()
    win.attributes("-topmost", True)
    win.after(200, lambda: win.attributes("-topmost", False))

    FONT_MONO  = ("Consolas", 10)
    FONT_UI    = ("Segoe UI", 10)
    FONT_HEAD  = ("Segoe UI", 11, "bold")
    FONT_TITLE = ("Segoe UI", 13, "bold")

    # ── State ──
    clusters      = []
    decisions     = {}
    check_has_run = [False]

    # ── Top bar ──
    topbar = tk.Frame(win, bg=HEADER_BG, pady=8)
    topbar.pack(fill="x")
    tk.Label(topbar, text="🧹  Clean Manifest", bg=HEADER_BG, fg="#ffffff",
             font=("Segoe UI", 14, "bold")).pack(side="left", padx=16)
    tk.Label(topbar, text="Scan a single manifest and remove internal duplicates",
             bg=HEADER_BG, fg="#adb5bd", font=FONT_UI).pack(side="left", padx=4)

    # ── Status bar ──
    status_bar = tk.Frame(win, bg=t["STATUS_BG"], pady=5)
    status_bar.pack(fill="x", side="bottom")
    status_var = tk.StringVar(value="")
    tk.Frame(status_bar, bg=t.get("PANEL_BDR", HEADER_BG), height=1).pack(fill="x", side="top")
    tk.Label(status_bar, textvariable=status_var, bg=t["STATUS_BG"], fg=t["STATUS_FG"],
             font=("Segoe UI", 9), anchor="w").pack(side="left", padx=12)

    # ── Main body: LEFT | DIVIDER | RIGHT (equal width) ──
    body = tk.Frame(win, bg=BG)
    body.pack(fill="both", expand=True, padx=12, pady=8)

    # Force true 50/50 columns
    body.columnconfigure(0, weight=45, uniform="clean_cols")  # left paste manifest
    body.columnconfigure(1, weight=0)                         # divider
    body.columnconfigure(2, weight=55, uniform="clean_cols")  # right results/cards
    body.rowconfigure(0, weight=0)
    body.rowconfigure(1, weight=1)

    # ── LEFT: input ──
    lh = tk.Frame(body, bg=HEADER_BG, pady=6)
    lh.grid(row=0, column=0, sticky="ew", pady=(0, 4))
    tk.Label(lh, text="📋  Paste Manifest", bg=HEADER_BG, fg="#ffffff",
             font=FONT_HEAD).pack(side="left", padx=10)
    tk.Button(lh, text="✕ Clear", bg=TAB_INACTIVE, fg="#ffffff",
              font=("Segoe UI", 8, "bold"), relief="flat",
              activebackground=HEADER_BG, activeforeground="#ffffff",
              padx=8, pady=2, cursor="hand2",
              command=lambda: manifest_text.delete("1.0", "end")).pack(side="right", padx=8)

    inp = tk.Frame(body, bg=PANEL_BG, bd=1, relief="solid",
                   highlightbackground=t["PANEL_BDR"], highlightthickness=1)
    inp.grid(row=1, column=0, sticky="nsew")
    isb = tk.Scrollbar(inp, bg=PANEL_BG, troughcolor=BG)
    isb.pack(side="right", fill="y")

    manifest_text = tk.Text(inp, width=1,
                            bg=PANEL_BG, fg=TEXT, insertbackground=ACCENT,
                            font=FONT_MONO, wrap="none", yscrollcommand=isb.set,
                            relief="flat", padx=10, pady=10,
                            selectbackground=ACCENT, selectforeground="#fff")
    manifest_text.pack(fill="both", expand=True)
    isb.config(command=manifest_text.yview)

    # ── Divider ──
    tk.Frame(body, bg=ACCENT, width=2).grid(row=0, column=1, rowspan=2, sticky="ns", padx=8)

    # ── RIGHT: counter + buttons + cards ──
    rp = tk.Frame(body, bg=BG)
    rp.grid(row=0, column=2, rowspan=2, sticky="nsew")
    rp.columnconfigure(0, weight=1)
    rp.rowconfigure(1, weight=1)

    # Right header: counter label
    rh = tk.Frame(rp, bg=HEADER_BG, pady=6)
    rh.grid(row=0, column=0, sticky="ew", pady=(0, 4))
    action_counter = tk.StringVar(value="Scan to detect duplicates")
    tk.Label(rh, textvariable=action_counter, bg=HEADER_BG, fg="#ffffff",
             font=FONT_HEAD).pack(side="left", padx=10)

    # Scan + Generate row
    btn_row = tk.Frame(rp, bg=BG)
    btn_row.grid(row=0, column=0, sticky="ew", pady=(44, 0))  # offset below header

    # Rebuild rp layout — header row=0, btn_row=1, cards=2
    rp.rowconfigure(0, weight=0)
    rp.rowconfigure(1, weight=0)
    rp.rowconfigure(2, weight=1)

    rh.grid(row=0, column=0, sticky="ew", pady=(0, 0))
    btn_row.grid(row=1, column=0, sticky="ew", pady=(6, 6))

    tk.Button(btn_row, text="🔍  Scan for Duplicates",
              command=lambda: run_scan(), bg=ACCENT, fg="#fff",
              font=("Segoe UI", 9, "bold"), relief="flat",
              padx=12, pady=6, cursor="hand2").pack(side="left", padx=(0, 6))

    generate_btn = tk.Button(btn_row, text="🔒  Generate — Run a scan first",
                             bg=LOCKED_BG, fg=LOCKED_FG,
                             font=("Segoe UI", 9, "bold"), relief="flat",
                             padx=12, pady=6, state="disabled")
    generate_btn.pack(side="left")

    # Cards scrollable area
    results_outer = tk.Frame(rp, bg=BG)
    results_outer.grid(row=2, column=0, sticky="nsew")
    results_outer.columnconfigure(0, weight=1)
    results_outer.rowconfigure(0, weight=1)

    canvas = tk.Canvas(results_outer, bg=BG, highlightthickness=0, bd=0)
    canvas.grid(row=0, column=0, sticky="nsew")
    rsb = tk.Scrollbar(results_outer, orient="vertical", command=canvas.yview)
    rsb.grid(row=0, column=1, sticky="ns")
    canvas.configure(yscrollcommand=rsb.set)

    results_inner = tk.Frame(canvas, bg=BG)
    cwin = canvas.create_window((0, 0), window=results_inner, anchor="nw")
    results_inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>",        lambda e: canvas.itemconfig(cwin, width=e.width))
    canvas.bind_all("<MouseWheel>",   lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

    # ── Generate lock + counter ──
    def refresh_lock():
        total     = len(clusters)
        remaining = sum(1 for i in range(total)
                        if decisions.get(i, "— undecided —") == "— undecided —")
        actioned  = total - remaining
        if total > 0:
            action_counter.set(f"{actioned} / {total} actioned - ")
        else:
            action_counter.set("Scan to detect duplicates")
        if check_has_run[0] and remaining == 0:
            generate_btn.config(text="✓  Generate Clean List",
                                bg=KEEP_FG, fg=PANEL_BG, state="normal",
                                cursor="hand2", command=generate_clean)
        else:
            generate_btn.config(
                text=f"🔒  {remaining} still need action" if remaining > 0
                     else "🔒  Generate — Run a scan first",
                bg=LOCKED_BG, fg=LOCKED_FG, state="disabled", cursor="")

    # ── Generate clean ──
    def generate_clean():
        raw_lines     = manifest_text.get("1.0", "end").strip().splitlines()
        to_remove     = set()
        merge_replace = {}

        for i, cluster in enumerate(clusters):
            dec        = decisions.get(i, "keep_idx_0")
            passengers = cluster["passengers"]

            # Keep decision values internal/consistent.
            # Older buttons used display text like "Merge", "Keep A", "Keep B" and "keep both".
            # The clean-list generator expects lower-case/internal values, so normalize them here too.
            decision_aliases = {
                "Merge": "merge",
                "Keep A": "keep_idx_0",
                "Keep B": "keep_idx_1",
                "keep both": "keep_both",
                "keep_first": "keep_idx_0",
            }
            dec = decision_aliases.get(dec, dec)

            if dec == "keep_both":
                pass
            elif dec == "remove_all":
                for p in passengers:
                    to_remove.add(p["raw"])
            elif dec == "merge" and cluster.get("suggested_merge"):
                merge_replace[passengers[0]["raw"]] = cluster["suggested_merge"]
                for p in passengers[1:]:
                    to_remove.add(p["raw"])
            elif dec.startswith("keep_idx_"):
                keep_idx = int(dec.replace("keep_idx_", ""))
                for j, p in enumerate(passengers):
                    if j != keep_idx:
                        to_remove.add(p["raw"])
            else:
                # Fallback: keep the first entry and remove the rest.
                for p in passengers[1:]:
                    to_remove.add(p["raw"])

        output_lines = []
        for line in raw_lines:
            s = line.strip()
            if s in to_remove:
                continue
            elif s in merge_replace:
                output_lines.append(merge_replace[s])
            else:
                output_lines.append(line)

        seat_num = 1
        final_lines = []
        for line in output_lines:
            if re.search(r"Seat\s+Number\s+\d+", line, re.IGNORECASE):
                line = re.sub(r"(Seat\s+Number\s+)\d+",
                              lambda m: f"{m.group(1)}{seat_num:02d}", line, flags=re.IGNORECASE)
                seat_num += 1
            final_lines.append(line)

        clean_output = "\n".join(final_lines)

        out_win = tk.Toplevel(win)
        out_win.title("Clean Manifest Output")
        out_win.geometry("900x600")
        out_win.configure(bg=BG)
        hdr = tk.Frame(out_win, bg=HEADER_BG, pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Clean Manifest — Ready to Copy",
                 bg=HEADER_BG, fg="#ffffff", font=FONT_TITLE).pack(side="left", padx=16)
        tk.Label(out_win,
                 text=f"{seat_num-1} passenger(s)  ·  {len(to_remove)} removed  ·  {len(merge_replace)} merged",
                 bg=BG, fg=SUBTEXT, font=FONT_UI).pack(padx=12, pady=(8, 2), anchor="w")
        of = tk.Frame(out_win, bg=PANEL_BG)
        of.pack(fill="both", expand=True, padx=12, pady=8)
        osb = tk.Scrollbar(of)
        osb.pack(side="right", fill="y")
        ot = tk.Text(of, bg=PANEL_BG, fg=TEXT, font=FONT_MONO,
                     wrap="none", yscrollcommand=osb.set, relief="flat", padx=8, pady=8)
        ot.pack(fill="both", expand=True)
        osb.config(command=ot.yview)
        ot.insert("1.0", clean_output)
        ot.config(state="disabled")
        def copy_clean():
            out_win.clipboard_clear()
            out_win.clipboard_append(clean_output)
            out_win.update()
            messagebox.showinfo("Copied", "Clean manifest copied to clipboard.", parent=out_win)
        tk.Button(out_win, text="Copy to Clipboard", command=copy_clean,
                  bg=ACCENT, fg="#fff", font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=14, pady=6, cursor="hand2").pack(pady=(0, 12))

    # ── Card builder ──
    def rebuild_cards(preserve_scroll=False):
        scroll_pos = canvas.yview()[0] if preserve_scroll else 0.0
        for w in results_inner.winfo_children():
            w.destroy()

        if not clusters:
            msg = ("✓  No duplicates found — manifest is clean."
                   if check_has_run[0] else
                   "Paste a manifest and click Scan for Duplicates.")
            fg  = KEEP_FG if check_has_run[0] else SUBTEXT
            tk.Label(results_inner, text=msg, bg=BG, fg=fg,
                     font=("Segoe UI", 11, "bold" if check_has_run[0] else "normal")).pack(pady=20)
            if check_has_run[0]:
                tk.Label(results_inner,
                         text="Click Generate Clean List to produce the renumbered output.",
                         bg=BG, fg=SUBTEXT, font=FONT_UI).pack()
            win.after(10, lambda: canvas.yview_moveto(scroll_pos))
            return

        for cidx, cluster in enumerate(clusters):
            ctype      = cluster["type"]
            passengers = cluster["passengers"]
            dec        = decisions.get(cidx, "— undecided —")
            is_multi   = len(passengers) > 2
            is_actioned = dec != "— undecided —"

            # Border turns green when actioned
            if ctype == "exact":
                card_bg = EXACT_BG; badge_fg = EXACT_FG
                border_col = KEEP_FG if is_actioned else t["CARD_BORDER_EXACT"]
                badge_text = "● CONFIRMED DUPLICATE"
            elif ctype == "merge":
                card_bg = MERGE_BG; badge_fg = MERGE_FG
                border_col = KEEP_FG if is_actioned else t["CARD_BORDER_MERGE"]
                badge_text = "◈ MERGE SUGGESTED"
            else:
                card_bg = POSSIBLE_BG; badge_fg = POSSIBLE_FG
                border_col = KEEP_FG if is_actioned else t["CARD_BORDER_POSSIBLE"]
                badge_text = "◉ POSSIBLE DUPLICATE"

            if is_multi:
                badge_text = f"⚠ {len(passengers)} ENTRIES — same person detected"

            wrap = tk.Frame(results_inner, bg=BG)
            wrap.pack(fill="x", padx=4, pady=3)
            tk.Frame(wrap, bg=border_col, width=4).pack(side="left", fill="y")
            card = tk.Frame(wrap, bg=card_bg, pady=8, padx=12)
            card.pack(side="left", fill="both", expand=True)

            tk.Label(card, text=badge_text, bg=card_bg, fg=badge_fg,
                     font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 4))

            if not is_multi:
                p_a, p_b = passengers[0], passengers[1]
                for lbl, p in [("ENTRY A:", p_a), ("ENTRY B:", p_b)]:
                    row = tk.Frame(card, bg=card_bg)
                    row.pack(fill="x", pady=1)
                    tk.Label(row, text=lbl, bg=card_bg, fg=SUBTEXT,
                             font=("Segoe UI", 9, "bold"), width=10, anchor="w").pack(side="left")
                    tk.Label(row, text=p["raw"], bg=card_bg, fg=TEXT,
                             font=FONT_MONO, anchor="w").pack(side="left")

                if ctype == "merge" and cluster.get("suggested_merge"):
                    mr = tk.Frame(card, bg=card_bg)
                    mr.pack(fill="x", pady=(2, 4))
                    tk.Label(mr, text="MERGE →", bg=card_bg, fg=MERGE_FG,
                             font=("Segoe UI", 9, "bold"), width=10, anchor="w").pack(side="left")
                    tk.Label(mr, text=cluster["suggested_merge"], bg=card_bg,
                             fg=MERGE_FG, font=FONT_MONO, anchor="w").pack(side="left")

                tk.Frame(card, bg=t.get("PANEL_BDR", SUBTEXT), height=1).pack(fill="x", pady=(6, 4))
                bf = tk.Frame(card, bg=card_bg)
                bf.pack(fill="x")

                def mk(label, dv, bg_c, fg_c, ci=cidx):
                    def cb():
                        decisions[ci] = dv
                        refresh_lock()
                        rebuild_cards(preserve_scroll=True)
                    tk.Button(bf, text=label, command=cb, bg=bg_c, fg=fg_c,
                              font=("Segoe UI", 8, "bold"), relief="flat",
                              activebackground=bg_c, activeforeground=fg_c,
                              padx=8, pady=3, cursor="hand2").pack(side="left", padx=(0, 4))

                if ctype == "exact":
                    tk.Label(bf, text="✓ Auto: Keep A", bg=card_bg, fg=KEEP_FG,
                             font=("Segoe UI", 8, "bold")).pack(side="left", padx=(0, 8))
                    mk("Keep B",        "keep_idx_1", TAB_INACTIVE, "#ffffff")
                    mk("Keep Both",     "keep_both",  TAB_INACTIVE, "#ffffff")
                    mk("✕ Remove Both", "remove_all", EXACT_FG,     "#ffffff")
                elif ctype == "merge":
                    mk("Keep A",        "keep_idx_0", KEEP_FG,      "#ffffff")
                    mk("Keep B",        "keep_idx_1", HEADER_BG,    "#ffffff")
                    mk("Accept Merge",  "merge",       MERGE_FG,     "#ffffff")
                    mk("Keep Both",     "keep_both",   TAB_INACTIVE, "#ffffff")
                else:
                    mk("Keep A",        "keep_idx_0", KEEP_FG,      "#ffffff")
                    mk("Keep B",        "keep_idx_1", HEADER_BG,    "#ffffff")
                    mk("Keep Both",     "keep_both",  TAB_INACTIVE, "#ffffff")
                    mk("✕ Remove Both", "remove_all", EXACT_FG,     "#ffffff")

                dec_display = "auto: keep A" if (ctype == "exact" and dec == "— undecided —") else dec
                tk.Label(bf, text=dec_display, bg=card_bg, fg=SUBTEXT,
                         font=("Segoe UI", 8, "italic")).pack(side="left", padx=6)

            else:
                tk.Label(card, text="Click a row to select which entry to keep:",
                         bg=card_bg, fg=SUBTEXT, font=("Segoe UI", 8)).pack(anchor="w", pady=(0, 4))

                sel = [None]
                if dec.startswith("keep_idx_"):
                    sel[0] = int(dec.replace("keep_idx_", ""))

                for pi, p in enumerate(passengers):
                    is_sel = sel[0] == pi
                    rbg = KEEP_FG if is_sel else card_bg
                    rfg = "#000"  if is_sel else TEXT
                    row = tk.Frame(card, bg=rbg, cursor="hand2",
                                   highlightbackground=KEEP_FG if is_sel else t.get("PANEL_BDR", SUBTEXT),
                                   highlightthickness=2 if is_sel else 1)
                    row.pack(fill="x", pady=2, padx=2)
                    slbl = tk.Label(row, text=f"Seat {p['seat']:02d}:" if p["seat"] else "??:",
                                    bg=rbg, fg=rfg, font=("Segoe UI", 8, "bold"), width=8, anchor="w")
                    slbl.pack(side="left", padx=(4, 2), pady=3)
                    llbl = tk.Label(row, text=p["raw"], bg=rbg, fg=rfg,
                                    font=FONT_MONO, anchor="w")
                    llbl.pack(side="left", pady=3, fill="x", expand=True)
                    if is_sel:
                        tk.Label(row, text="✓", bg=rbg, fg="#000",
                                 font=("Segoe UI", 9, "bold")).pack(side="right", padx=6)

                    def on_row(idx=pi, ci=cidx):
                        decisions[ci] = f"keep_idx_{idx}"
                        refresh_lock()
                        rebuild_cards(preserve_scroll=True)
                    for ww in (row, slbl, llbl):
                        ww.bind("<Button-1>", lambda e, f=on_row: f())

                tk.Frame(card, bg=t.get("PANEL_BDR", SUBTEXT), height=1).pack(fill="x", pady=(6, 4))
                ef = tk.Frame(card, bg=card_bg)
                ef.pack(fill="x")

                def mk_e(label, dv, bg_c, fg_c, ci=cidx):
                    def cb():
                        decisions[ci] = dv
                        refresh_lock()
                        rebuild_cards(preserve_scroll=True)
                    tk.Button(ef, text=label, command=cb, bg=bg_c, fg=fg_c,
                              font=("Segoe UI", 8, "bold"), relief="flat",
                              activebackground=bg_c, activeforeground=fg_c,
                              padx=8, pady=3, cursor="hand2").pack(side="left", padx=(0, 4))

                mk_e("Keep All",     "keep_both",  TAB_INACTIVE, "#ffffff")
                mk_e("✕ Remove All", "remove_all", EXACT_FG,     "#ffffff")
                hint = "← click a row to select" if sel[0] is None else "✓ selection made"
                tk.Label(ef, text=hint, bg=card_bg, fg=SUBTEXT,
                         font=("Segoe UI", 8, "italic")).pack(side="left", padx=6)

        win.after(10, lambda sp=scroll_pos: canvas.yview_moveto(sp))

    # ── Run scan ──
    def run_scan():
        nonlocal clusters
        clusters = []
        decisions.clear()
        check_has_run[0] = False

        raw = manifest_text.get("1.0", "end").strip()
        if not raw:
            status_var.set("⚠  Paste a manifest first.")
            return
        passengers = parse_manifest_lines(raw)
        if len(passengers) < 2:
            status_var.set("⚠  Need at least 2 seat entries to scan.")
            return

        clusters = scan_manifest_for_duplicates(passengers)
        check_has_run[0] = True

        # Auto-set exact 2-entry clusters
        for i, c in enumerate(clusters):
            if c["type"] == "exact" and len(c["passengers"]) == 2:
                decisions[i] = "keep_first"

        exact   = sum(1 for c in clusters if c["type"] == "exact")
        possible= sum(1 for c in clusters if c["type"] == "possible")
        merge   = sum(1 for c in clusters if c["type"] == "merge")
        status_var.set(
            f"✓ Scan complete — {exact} confirmed, {possible} possible, "
            f"{merge} merge suggested  ·  {len(passengers)} entries scanned."
        )
        refresh_lock()
        rebuild_cards()


# ─────────────────────────────────────────────
# HELP WINDOW
# ─────────────────────────────────────────────

def open_help_window(parent):
    t = get_theme()
    BG = t["BG"]; PANEL_BG = t["PANEL_BG"]; HEADER_BG = t["HEADER_BG"]
    ACCENT = t["ACCENT"]; TEXT = t["TEXT"]; SUBTEXT = t["SUBTEXT"]
    PANEL_BDR = t["PANEL_BDR"]; TAB_ACTIVE = t["TAB_ACTIVE"]; TAB_INACTIVE = t["TAB_INACTIVE"]
    FONT_UI = ("Segoe UI", 10); FONT_HEAD = ("Segoe UI", 11, "bold")
    FONT_TITLE = ("Segoe UI", 13, "bold"); FONT_MONO = ("Consolas", 10)

    help_win = tk.Toplevel(parent)
    help_win.title("Help / Read")
    help_win.geometry("950x650")
    help_win.minsize(780, 520)
    help_win.configure(bg=BG)
    help_win.lift(); help_win.focus_force()

    topbar_h = tk.Frame(help_win, bg=HEADER_BG, pady=10)
    topbar_h.pack(fill="x")
    tk.Label(topbar_h, text="Help Knowledge", bg=HEADER_BG, fg="#ffffff",
             font=("Segoe UI", 14, "bold")).pack(side="left", padx=16)
    tk.Label(topbar_h, text="Quick Guides on how to use this application",
             bg=HEADER_BG, fg=t.get("LABEL_FG", "#a8bbd4"), font=FONT_UI).pack(side="left", padx=4)
    tk.Label(topbar_h, text="Created by Thomas Brayovic",
             bg=HEADER_BG, fg="#A2A2A2", font=FONT_UI).pack(side="right", padx=10)

    body_h = tk.Frame(help_win, bg=BG)
    body_h.pack(fill="both", expand=True, padx=12, pady=10)
    body_h.rowconfigure(1, weight=1)
    body_h.columnconfigure(0, weight=1)

    active_tab = tk.StringVar(value="duplicates")
    tab_row = tk.Frame(body_h, bg=BG)
    tab_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))

    content_outer = tk.Frame(body_h, bg=PANEL_BG,
                             highlightbackground=PANEL_BDR, highlightthickness=1)
    content_outer.grid(row=1, column=0, sticky="nsew")
    content_outer.rowconfigure(0, weight=1)
    content_outer.columnconfigure(0, weight=1)

    canvas_h = tk.Canvas(content_outer, bg=PANEL_BG, highlightthickness=0)
    canvas_h.grid(row=0, column=0, sticky="nsew")
    scrollbar_h = tk.Scrollbar(content_outer, orient="vertical", command=canvas_h.yview)
    scrollbar_h.grid(row=0, column=1, sticky="ns")
    canvas_h.configure(yscrollcommand=scrollbar_h.set)

    content_frame = tk.Frame(canvas_h, bg=PANEL_BG)
    canvas_win = canvas_h.create_window((0, 0), window=content_frame, anchor="nw")
    content_frame.bind("<Configure>", lambda e: canvas_h.configure(scrollregion=canvas_h.bbox("all")))
    canvas_h.bind("<Configure>", lambda e: canvas_h.itemconfig(canvas_win, width=e.width))

    def clear_content():
        for w in content_frame.winfo_children(): w.destroy()

    def section(title):
        tk.Label(content_frame, text=title, bg=PANEL_BG, fg=ACCENT,
                 font=FONT_TITLE, anchor="w").pack(fill="x", padx=18, pady=(18, 6))

    def paragraph(text):
        tk.Label(content_frame, text=text, bg=PANEL_BG, fg=TEXT, font=FONT_UI,
                 justify="left", anchor="w", wraplength=840).pack(fill="x", padx=18, pady=(0, 8))

    def help_card(title, body_text, accent_colour=None):
        border_colour = accent_colour or ACCENT
        wrap = tk.Frame(content_frame, bg=PANEL_BG)
        wrap.pack(fill="x", padx=18, pady=5)
        tk.Frame(wrap, bg=border_colour, width=4).pack(side="left", fill="y")
        card = tk.Frame(wrap, bg=PANEL_BG, highlightbackground=PANEL_BDR,
                        highlightthickness=1, padx=12, pady=8)
        card.pack(side="left", fill="x", expand=True)
        tk.Label(card, text=title, bg=PANEL_BG, fg=border_colour,
                 font=FONT_HEAD, anchor="w").pack(fill="x")
        tk.Label(card, text=body_text, bg=PANEL_BG, fg=TEXT, font=FONT_UI,
                 justify="left", anchor="w", wraplength=800).pack(fill="x", pady=(4, 0))

    def code_note(text):
        tk.Label(content_frame, text=text, bg=t.get("STATUS_BG", PANEL_BG),
                 fg=t.get("STATUS_FG", SUBTEXT), font=FONT_MONO, justify="left",
                 anchor="w", wraplength=840, padx=10, pady=8).pack(fill="x", padx=18, pady=(4, 10))

    def render_clean_manifest():
        clear_content()

        section("Clean Manifest")

        paragraph(
            "The Clean Manifest tool is used to review a pasted manifest, identify duplicate or repeated "
            "passenger entries, then generate a cleaned and renumbered manifest output."
        )

        section("How to use it")

        help_card(
            "1. Paste the manifest",
            "Paste the full manifest into the left side of the Clean Manifest window. "
            "Each passenger should be on its own line and should include a seat number."
        )

        help_card(
            "2. Scan for duplicates",
            "Click Scan for Duplicates. The app will group entries that appear to belong to the same person."
        )

        help_card(
            "3. Review the cards",
            "Each card shows the entries that may be duplicates. For simple duplicates, the app shows Entry A and Entry B. "
            "For three or more matches, the card shows all detected entries together."
        )

        help_card(
            "4. Choose what to keep",
            "Select which entry should remain in the final manifest. You can keep one entry, keep both/all entries, "
            "accept a merge where available, or remove entries depending on the card type."
        )

        help_card(
            "5. Generate the clean list",
            "Once all required review items have been actioned, Generate Clean List will unlock. "
            "The final output will remove selected duplicates and renumber the seats in order."
        )

        section("Card status meanings")

        help_card(
            "Confirmed Duplicate",
            "A high-confidence match. The app believes the entries are the same person based on matching details.",
            t["CARD_BORDER_EXACT"]
        )

        help_card(
            "Possible Duplicate",
            "A lower-confidence match. This may happen when names are similar, phone numbers match, "
            "or some details are missing.",
            t["CARD_BORDER_POSSIBLE"]
        )

        help_card(
            "Merge Suggested",
            "The app thinks both entries are the same person, but each line may contain information missing from the other. "
            "Accept Merge keeps the combined version.",
            t["CARD_BORDER_MERGE"]
        )

        help_card(
            "Same person detected",
            "This appears when three or more entries are grouped together for the same passenger. "
            "Click the row you want to keep, or choose Keep All / Remove All.",
            t["POSSIBLE_FG"]
        )

        section("Action buttons")

        help_card(
            "Keep A / Keep B",
            "Keeps the selected entry and removes the other duplicate entry from the cleaned output."
        )

        help_card(
            "Keep Both / Keep All",
            "Keeps all shown entries. Use this when the passengers are actually different people."
        )

        help_card(
            "Accept Merge",
            "Uses the suggested merged entry where the app has built a more complete passenger line."
        )

        help_card(
            "Remove Both / Remove All",
            "Removes the duplicate group from the final cleaned manifest."
        )

        section("Important notes")

        code_note(
            "Clean Manifest does not permanently change the original pasted text.\n"
            "It only generates a new cleaned output after you review the duplicate cards.\n\n"
            "The final output is renumbered from Seat Number 01 onwards."
        )

    def render_duplicates():
        clear_content()
        section("Duplicate Checker")
        paragraph("The Duplicate Checker compares the newly generated manifest against an existing manifest. "
                  "It looks for matching names, phone numbers, flight details and incomplete records.")
        section("Status meanings")
        help_card("Confirmed Duplicate",
                  "The app is confident these are the same person — name and phone both match. "
                  "Shown in the Confirmed tab and auto-resolved.", t["CARD_BORDER_EXACT"])
        help_card("Possible Duplicate",
                  "A likely match but not certain. Could be a partial name match, "
                  "phone match with different name, or same name with different flight details. "
                  "Requires your confirmation.", t["CARD_BORDER_POSSIBLE"])
        help_card("Merge Suggested",
                  "Both entries appear to be the same passenger but each has information the other is missing. "
                  "The app suggests a combined record.", t["CARD_BORDER_MERGE"])
        help_card("Multiple Duplicates",
                  "More than one new entry matched the same existing record. "
                  "The app picks the most complete entry as the best to keep.", t["POSSIBLE_FG"])
        section("Review tabs")
        help_card("Confirmed Duplicates tab",
                  "High-confidence duplicates. By default the app keeps the new entry and drops the existing one.")
        help_card("Needs Review tab",
                  "Items needing a user decision. Generate Clean List stays locked until all are resolved.")
        section("Action buttons")
        help_card("Confirm Keep New / Keep New", "Keeps the new entry, removes the matching existing record.")
        help_card("Keep Existing", "Keeps the existing record, removes the new entry.")
        help_card("Keep Both", "Keeps both entries. Use when they are actually different people.")
        help_card("Accept Merge", "Uses the suggested merged record combining both entries.")
        help_card("✕ Remove", "Removes the new entry from the final output entirely.")
        section("Clean Manifest Output")
        paragraph("Once every review item has a decision the Generate Clean List button unlocks. "
                  "The final manifest removes or merges selected duplicates and renumbers seats in order. "
                  "If no duplicates are found the button unlocks immediately after the check.")

    def render_shuttle_manifest():
        clear_content()

        section("Shuttle Manifester")

        paragraph(
            "The Shuttle Manifest screen is used to turn copied Excel passenger data into a clean, "
            "standardised manifest format. It reads the pasted data using the selected profile, applies "
            "the correct output template, assigns seat numbers, and produces a formatted manifest ready to copy."
        )

        section("How to use it")

        help_card(
            "1. Select the correct profile",
            "Use the Profile dropdown at the top of the app to choose the correct manifest format. "
            "Each profile controls which Excel headers are read and how the final passenger lines are formatted."
        )

        help_card(
            "2. Paste data from Excel",
            "Paste the copied Excel data into the input box. The pasted data should include the header row, "
            "because the app uses those headers to find passenger details such as name, contact number, pickup details, "
            "flight information, bags, and destination."
        )

        help_card(
            "3. Set the starting seat number",
            "Use the Start Seat field to choose which seat number the manifest should begin from. "
            "This is useful when adding more passengers to an existing manifest."
        )

        help_card(
            "4. Review the formatted output",
            "The app automatically formats the pasted data into the output box. "
            "The first line is the shuttle service label, followed by the passenger entries."
        )

        help_card(
            "5. Copy the manifest",
            "Click Copy Manifest to copy the formatted passenger list. "
            "This copies the passenger entries from the output box so they can be pasted into the required system or document."
        )

        help_card(
            "6. Copy the service seats label",
            "Click Copy Service Seats to copy the shuttle service label with the final seat count added. "
            "This is useful when you need a quick service summary showing the destination, pickup details, and number of seats."
        )

        section("Main buttons")

        help_card(
            "Clear All",
            "Clears the pasted input, formatted output, cached service label, and resets the starting seat number back to 1."
        )

        help_card(
            "Copy Manifest",
            "Copies the formatted passenger entries from the output box, excluding the shuttle service heading."
        )

        help_card(
            "Copy Service Seats",
            "Copies the shuttle service label with the total number of detected seats added to the end."
        )

        help_card(
            "Check Duplicates",
            "Opens the Duplicate Checker using the current formatted manifest output. "
            "This allows the current manifest to be compared against an existing manifest before finalising."
        )

        section("Profiles")

        paragraph(
            "Profiles control how the app reads and formats pasted Excel data. "
            "For example, one profile may expect Defence Travel column names, while another may expect Toll column names. "
            "If the wrong profile is selected, the app may show a missing fields message or produce incomplete output."
        )

        help_card(
            "Defence Travel",
            "Used for standard Defence Travel-style Excel data with fields such as First Name, Last Name, Contact Details, "
            "Flight Route, Flight Time, Bags, pickup location, destination, and cost details."
        )


        section("Output behaviour")

        help_card(
            "Service label",
            "The first output line is generated from the selected profile's group label format. "
            "It usually shows the pickup location, destination, pickup time, and pickup date."
        )

        help_card(
            "Passenger rows",
            "Each passenger row is generated from the selected profile's output format. "
            "The app fills in values such as seat number, passenger name, contact number, flight details, and bags."
        )

        help_card(
            "Seat numbering",
            "Seat numbers are assigned automatically from the Start Seat value. "
            "For example, if Start Seat is set to 5, the first passenger will be Seat Number 05."
        )

        help_card(
            "Seat warning",
            "Seat numbers from 22 onwards are highlighted as a warning. "
            "This helps flag when the manifest may be reaching or exceeding the expected shuttle capacity."
        )

        section("Common messages")

        help_card(
            "Missing required fields",
            "This means the pasted Excel data does not contain one or more column headers required by the selected profile. "
            "Check that the correct profile is selected and that the copied Excel data includes the header row."
        )

        help_card(
            "Not enough data to format",
            "This usually means the pasted data does not contain enough rows. "
            "The app expects a header row plus at least one passenger row."
        )

        help_card(
            "Invalid Pickup Info",
            "This appears when the app cannot detect a valid pickup date and time in the expected format for the selected profile."
        )

        section("Important notes")

        code_note(
            "Always paste the Excel header row with the passenger data.\n"
            "Make sure the correct profile is selected before reviewing the output.\n"
            "Use Clean Manifest or Duplicate Checker before finalising if the manifest may contain repeated passengers.\n"
            "The app does not change the original Excel file. It only formats the copied data inside the app."
        )


    def render_profiles():
        clear_content()
        section("Profile Manager")
        paragraph("Profiles control how the app reads pasted Excel data and formats the manifest output. "
                  "Each profile is designed for a different spreadsheet layout.")
        section("What each field does")
        help_card("Profile Name",
                  "The name shown in the profile dropdown. Choose a name that identifies the data source.")
        help_card("Group Label Format",
                  "The heading line above the seat list. Usually contains route, time and date. "
                  "Uses {placeholders} that match your header field names.")
        help_card("Output Format",
                  "Controls how each passenger line is generated. "
                  "Use {seat_number:02} for the seat number and {Field Name} for any mapped field.")
        help_card("Header Mappings",
                  "Maps the app's internal field names to the actual column headers in the spreadsheet.")
        section("Internal Field Name vs Excel Column Header")
        paragraph("Internal Field Name is what you use inside {placeholders} in the formats above.\n"
                  "Excel Column Header is the exact text of the column heading in the spreadsheet you receive.")
        code_note("Example:\n"
                  "Internal: Contact Details  →  Excel: MOBILE NUMBER\n"
                  "Output uses: {Contact Details}\n"
                  "App reads column: MOBILE NUMBER")
        section("Export and Import")
        help_card("📤 Export Config",
                  "Saves all profiles to a shuttle_config.json file. Share with colleagues.", "#1c2b4a")
        help_card("📥 Import Config",
                  "Load a config file. Choose to merge with existing profiles or replace them all.", "#1c2b4a")
        help_card("↺ Reset Defaults",
                  "Restores original built-in profiles. Use with caution — replaces all custom profiles.",
                  t["EXACT_FG"])

        # Tab buttons
    manifest_btn = tk.Button(
        tab_row,
        text="Shuttle Manifest",
        command=lambda: switch_tab("shuttle_manifest"),
        bg=TAB_ACTIVE,
        fg="#ffffff",
        font=("Segoe UI", 10, "bold"),
        relief="flat",
        padx=16,
        pady=8,
        cursor="hand2",
        activebackground=TAB_ACTIVE,
        activeforeground="#ffffff"
    )
    manifest_btn.pack(side="left", padx=(0, 2))


    dup_btn = tk.Button(
        tab_row,
        text="Duplicate Checker",
        command=lambda: switch_tab("duplicates"),
        bg=TAB_INACTIVE,
        fg="#ffffff",
        font=("Segoe UI", 10, "bold"),
        relief="flat",
        padx=16,
        pady=8,
        cursor="hand2",
        activebackground=TAB_INACTIVE,
        activeforeground="#ffffff"
    )
    dup_btn.pack(side="left", padx=(0, 2))


    clean_btn = tk.Button(
        tab_row,
        text="Clean Manifest",
        command=lambda: switch_tab("clean_manifest"),
        bg=TAB_INACTIVE,
        fg="#ffffff",
        font=("Segoe UI", 10, "bold"),
        relief="flat",
        padx=16,
        pady=8,
        cursor="hand2",
        activebackground=TAB_INACTIVE,
        activeforeground="#ffffff"
    )
    clean_btn.pack(side="left")


    profile_btn = tk.Button(
        tab_row,
        text="Profile Manager",
        command=lambda: switch_tab("profiles"),
        bg=TAB_INACTIVE,
        fg="#ffffff",
        font=("Segoe UI", 10, "bold"),
        relief="flat",
        padx=16,
        pady=8,
        cursor="hand2",
        activebackground=TAB_INACTIVE,
        activeforeground="#ffffff"
    )
    profile_btn.pack(side="left", padx=(0, 2))



    def switch_tab(tab_name):
        active_tab.set(tab_name)

        manifest_btn.config(
            bg=TAB_ACTIVE if tab_name == "shuttle_manifest" else TAB_INACTIVE,
            activebackground=TAB_ACTIVE if tab_name == "shuttle_manifest" else TAB_INACTIVE
        )

        dup_btn.config(
            bg=TAB_ACTIVE if tab_name == "duplicates" else TAB_INACTIVE,
            activebackground=TAB_ACTIVE if tab_name == "duplicates" else TAB_INACTIVE
        )

        profile_btn.config(
            bg=TAB_ACTIVE if tab_name == "profiles" else TAB_INACTIVE,
            activebackground=TAB_ACTIVE if tab_name == "profiles" else TAB_INACTIVE
        )

        clean_btn.config(
            bg=TAB_ACTIVE if tab_name == "clean_manifest" else TAB_INACTIVE,
            activebackground=TAB_ACTIVE if tab_name == "clean_manifest" else TAB_INACTIVE
        )

        if tab_name == "shuttle_manifest":
            render_shuttle_manifest()
        elif tab_name == "duplicates":
            render_duplicates()
        elif tab_name == "profiles":
            render_profiles()
        elif tab_name == "clean_manifest":
            render_clean_manifest()

    # Load default help tab
    switch_tab("shuttle_manifest")

if __name__ == "__main__":
    launch_gui()


