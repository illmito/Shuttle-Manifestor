import os
import json
import tkinter as tk
from tkinter import ttk, messagebox
import string
import re

# Config path
CONFIG_PATH = "config.json"

# Default config
default_config = {
  "__instructions": [
      
    "=== Passenger Formatter Configuration ===",
    "",
    "This file controls how the app reads and formats data pasted from Excel.",
    "If you're not sure what you're doing, please read everything below before editing.",
    "",
    "----------------------------------------",
    "1. STRUCTURE OVERVIEW",
    "----------------------------------------",
    "- 'profiles' contains one or more profile templates.",
    "- Each profile describes how to read the pasted data and how to format the output.",
    "- The app will use the first profile unless updated to support profile selection.",
    "",
    "----------------------------------------",
    "2. WHAT EACH SECTION MEANS",
    "----------------------------------------",
    "'headers':",
    "    - This is a mapping of internal field names to the actual column headers from Excel.",
    "    - The internal field name (on the LEFT) must match any placeholders used in formatting.",
    "    - The value (on the RIGHT) must exactly match how the header appears in Excel.",
    "    Example:",
    "        'First Name': 'First Name' (Excel column header is 'First Name')",
    "",
    "'group_label':",
    "    - This is a heading shown before the formatted seat list.",
    "    - You can use placeholders inside curly braces {like_this} to pull data from the first row.",
    "    - Placeholders MUST match one of the keys in the 'headers' section.",
    "    Example:",
    "        Shuttle: {Transport Pick Up Location} > {Destination Address} At {Transport Pick Up Time}",
    "",
    "'output_format':",
    "    - This controls the format of each individual passenger entry.",
    "    - You can use {seat_number:02} to show the seat number with leading zeros (e.g. 01, 02).",
    "    - Like 'group_label', all placeholders must match something in 'headers'.",
    "    Example:",
    "        Seat Number {seat_number:02}: {First Name} {Last Name} {Contact Details}",
    "",
    "----------------------------------------",
    "3. RULES TO FOLLOW",
    "----------------------------------------",
    "- DO NOT leave empty curly braces like {} — always give them a name: {Flight Time}, {seat_number}",
    "- All placeholder names in 'group_label' and 'output_format' must exist in the 'headers'.",
    "- The Excel data you paste into the app must contain column headers that match the headers defined here.",
    "- If you change any formatting, double-check that all field names are consistent.",
    "- Strings are case-sensitive — 'First Name' is different from 'first name'.",
    "",
    "----------------------------------------",
    "4. TROUBLESHOOTING",
    "----------------------------------------",
    "- If you see a message like '[!] Missing required fields', check for typos in your headers or placeholders.",
    "- If the app crashes or shows '[!] Could not generate group label', make sure the first data row is complete.",
    "- You can always ask ChatGPT to help you fix or extend this file — it understands this format.",
    "",
    "----------------------------------------",
    "5. ADVANCED (Optional)",
    "----------------------------------------",
    "- You can add new profiles (e.g., 'airport_transfer', 'train_manifest') under 'profiles'.",
    "- Make sure each new profile has its own 'headers', 'group_label', and 'output_format'.",
    "- Profile switching is not built into the app UI yet but can be done in code.",
    "---- RECOMMENDED TO NOT CHANGE ANY PROFILE SETTING ---- EDIT BELOW 'shuttle' ---- "
    "",
    "----------------------------------------",


    "End of Instructions"


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
            "group_label": " {Transport Pick Up Location} > {Destination Address} @ {Pickup Info} ",
            "output_format": "Seat Number {seat_number:02} {First Name} {Last Name} {Contact Details} {No of Bags}x Bags"
        }
    }
}

# Load config or create default
if not os.path.exists(CONFIG_PATH):
    with open(CONFIG_PATH, "w") as f:
        json.dump(default_config, f, indent=2)

with open(CONFIG_PATH, "r") as f:
    config = json.load(f)

profile_key = next(iter(config["profiles"]))
profile = config["profiles"][profile_key]





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
                        # Look for date and time anywhere in the string
                        match = re.search(r"\d{2}/\d{2}/\d{3,4}-?\s*\d{1,2}:\d{2}\s*[ap]m", raw_value, re.IGNORECASE)
                        if match:
                            row_data[key] = match.group(0)
                        else:
                            row_data[key] = "[Invalid Pickup Info]"
                    else:
                        row_data[key] = raw_value



            results.append(profile["output_format"].format(**row_data))
            seat_number += 1
        except Exception as e:
            results.append(f"[!] Error processing row: {e}")
    return results


def launch_gui():
    root = tk.Tk()
    root.title("Shuttle Manifester")
    seat_start_var = tk.StringVar(value="1")


    profile_var = tk.StringVar(value=profile_key)


# ADD THIS FUNCTION NEAR THE TOP WITH OTHER UTILITY FUNCTIONS

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
    seat_count_cache = tk.IntVar()


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
                elif line.startswith("[!]"):
                    output_text.insert(tk.END, line + "\n", "error")
                else:
                    start_index = output_text.index("end-1c")
                    output_text.insert(tk.END, line + "\n")

                    match = re.search(r"Seat\s+Number\s+(\d+)", line)
                    if match:
                        seat_num = int(match.group(1))
                        if seat_num >= 22:
                            seat_start = f"{start_index}+{match.start(1)}c"
                            seat_end = f"{start_index}+{match.end(1)}c"
                            output_text.tag_add("seat_warning", seat_start, seat_end)
        group_label_cache.set(formatted[0]) 
                            

        output_text.config(state="normal")  # Make read-only after update



    def copy_output():
        root.clipboard_clear()
        root.clipboard_append(output_text.get("2.0", tk.END))
        root.update()



    def on_clear_all():
        seat_start_var.set("1")
        input_text.delete("1.0", tk.END)
        output_text.delete("1.0", tk.END)


    def create_context_menu(widget):
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(label="Undo", command=lambda: widget.event_generate("<<Undo>>"))
        menu.add_separator()
        menu.add_command(label="Cut", command=lambda: widget.event_generate("<<Cut>>"))
        menu.add_command(label="Copy", command=lambda: widget.event_generate("<<Copy>>"))
        menu.add_command(label="Paste", command=lambda: widget.event_generate("<<Paste>>"))
        menu.add_separator()
        menu.add_command(label="Select All", command=lambda: widget.event_generate("<<SelectAll>>"))

        def show_menu(event):
            widget.focus()
            menu.tk_popup(event.x_root, event.y_root)

        widget.bind("<Button-3>", show_menu)
        widget.bind("<Button-2>", show_menu)
        widget.bind("<Control-z>", lambda e: widget.event_generate("<<Undo>>"))
        widget.bind("<Command-z>", lambda e: widget.event_generate("<<Undo>>"))


    def open_config_editor(parent):
        editor_win = tk.Toplevel(parent)
        editor_win.title("Edit Config.json")
        editor_win.geometry("800x600")

        text_widget = tk.Text(editor_win, wrap="none", font=("Consolas", 10))
        text_widget.pack(fill="both", expand=True)

        try:
            with open(CONFIG_PATH, "r") as f:
                json_data = json.load(f)
            text_widget.insert("1.0", json.dumps(json_data, indent=2))
        except Exception as e:
            messagebox.showerror("Error", f"Could not load config: {e}")
            return

        def save_changes():
            try:
                updated_json = json.loads(text_widget.get("1.0", "end"))
                with open(CONFIG_PATH, "w") as f:
                    json.dump(updated_json, f, indent=2)
                messagebox.showinfo("Saved", "Please Restart Application to apply Changes.")
                editor_win.destroy()
            except json.JSONDecodeError as e:
                messagebox.showerror("JSON Error", f"Invalid JSON: {e}")

        save_button = ttk.Button(editor_win, text="Save", command=save_changes)
        save_button.pack(pady=5)


    # Layout
    root.columnconfigure(0, weight=1)  # Make column stretchable

    # Menu bar
    menubar = tk.Menu(root)
    root.config(menu=menubar)
    config_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Config", menu=config_menu)
    config_menu.add_command(label="Edit JSON Config", command=lambda: open_config_editor(root))


    # Profile selector section
    profile_frame = ttk.Frame(root)
    profile_frame.grid(row=0, column=0, sticky="w", padx=10, pady=(5, 5))

    # Row 0: Profile label + combo box
    ttk.Label(profile_frame, text="Shuttle Manifest Profile:").grid(row=0, column=0, padx=(0, 5), sticky="w")
    profile_combo = ttk.Combobox(profile_frame, textvariable=profile_var, state="readonly", width=30)
    profile_combo["values"] = list(config["profiles"].keys())
    profile_combo.grid(row=0, column=1, sticky="w")
    profile_combo.bind("<<ComboboxSelected>>", on_profile_change)

    # Row 1: Paste instruction label (full-width under both widgets)
    ttk.Label(profile_frame, text="Paste data from Excel (with headers):").grid(
        row=1, column=0, columnspan=2, sticky="w", pady=(15, 0)
)



    # Input Frame
    input_frame = ttk.Frame(root)
    input_frame.grid(row=1, column=0, sticky="nsew", padx=10)

    root.rowconfigure(1, weight=1)  # Allow input frame to expand

    input_scroll = tk.Scrollbar(input_frame)
    input_scroll.pack(side="right", fill="y")

    input_text = tk.Text(input_frame, height=12, width=110, undo=True, wrap="none", yscrollcommand=input_scroll.set)
    input_text.pack(fill="both", expand=True)
    input_scroll.config(command=input_text.yview)
    create_context_menu(input_text)

    # Controls Frame (Buttons and Seat Spinner)

    form = ttk.Frame(root)
    form.grid(row=2, column=0, sticky="ew", padx=10, pady=(5, 5))
    form.columnconfigure(8, weight=1)  # Make column 8 stretchy to push button 9 to the right



    ttk.Label(form, text="Start Seat Number:").grid(row=0, column=0)
    seat_spin = tk.Spinbox(form, from_=1, to=999, textvariable=seat_start_var, width=6)
    seat_spin.grid(row=0, column=1, padx=(5, 20))

    ttk.Button(form, text="Clear All", command=on_clear_all).grid(row=0, column=5, padx=5)
    ttk.Button(form, text="Copy Manifest", command=copy_output).grid(row=0, column=6, padx=5)
    ttk.Button(form, text="Copy Service Seats", command=copy_group_label_with_seats).grid(row=0, column=8, padx=10, sticky="e")


    # Output Label
    # ttk.Label(root, text="Formatted Shuttle Manifest").grid(row=3, column=0, sticky="w", padx=10)

    # Output Frame
    output_frame = ttk.Frame(root)
    output_frame.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 10))
    root.rowconfigure(4, weight=1)  # Allow output frame to expand

    output_scroll = tk.Scrollbar(output_frame)
    output_scroll.pack(side="right", fill="y")

    output_text = tk.Text(output_frame, height=15, width=110, undo=True, wrap="none", yscrollcommand=output_scroll.set)
    output_text.pack(fill="both", expand=True)
    output_scroll.config(command=output_text.yview)
    create_context_menu(output_text)

    # Text Tag Styles
    output_text.config(selectbackground="#ffd966", selectforeground="black")

    output_text.tag_config("group_label", foreground="#fff9ef", background="#4a5b66", selectbackground="#5192bd",
                        font=("Segoe UI", 14, "bold italic"), spacing1=6, spacing3=6,
                        lmargin1=10, lmargin2=10, justify="center")
    
    output_text.tag_config("error", foreground="red", font=("Segoe UI", 10, "italic"))

    output_text.tag_config("seat_warning", background="#ffcccc")  # Light red background



    def trigger_on_paste(event):
        root.after(10, on_format)

    def watch_input_changes(event):
        root.after(10, on_format)

    def watch_seat_change(*args):
        on_format()

    input_text.bind("<Control-v>", trigger_on_paste)
    input_text.bind("<Command-v>", trigger_on_paste)
    input_text.bind("<<Paste>>", trigger_on_paste)
    input_text.bind("<KeyRelease>", watch_input_changes)

    seat_start_var.trace_add("write", watch_seat_change)

    input_text.focus_set()

    root.mainloop()


    root.after(100, on_format)


if __name__ == "__main__":
    launch_gui()


#$# pyinstaller --noconsole --onefile --icon=shuttle.ico main.py