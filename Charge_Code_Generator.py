#!/usr/bin/env python3
"""
Charge Code Generator GUI (CustomTkinter + Accordion) — Version 2.2.0

Changes in v2.2.0:
- New optional input columns supported: Start Date, End Date, OBS, REC/NRE, Cost Type (case-insensitive, flexible matching).
- Level 2 (Contract Type) is numeric-only and determined per row:
  • If Cost Type matches a Contract Types label (e.g., 'FFP' → '1', 'CRNF' → '2'), use the mapped numeric code.
  • If Cost Type is numeric (e.g., '2'), use that number.
  • Otherwise, default to the user-selected Contract Type (numeric code).
- Metadata columns are carried forward on full hierarchy rows, blank on partial parent rows.

Theme fixes:
- Root is always a CustomTkinter window—even with tkinterdnd2 (DnD)—via CTkDnD hybrid class.
- Mapping windows (ttk.Treeview + ttk.Scrollbar) are styled to match CTk dark/light appearance using a TtkStyler.
- Styles update live when Appearance changes (System/Light/Dark).
- White borders eliminated by using flat/zero border styles on ttk widgets.

Features preserved from original:
- Input: Browse file or folder; optional drag & drop (via tkinterdnd2 if installed).
- Level 1: Project Number (text).
- Level 2: Contract Type (dropdown; e.g., "1: FFP", "2: CRNF") + editor dialog.
- Mapping windows (with search, per-row checkboxes, Select All / Clear All, Import/Export):
  • CLIN (with special 'pad to 6' rule)
  • Phase (case-insensitive keys, uppercased)
  • WBS Code (maps numeric 6-char code extracted from WBS text)
  • WBS Description (case-insensitive keys, uppercased)
  (All mapping trees support vertical AND horizontal scrolling.)
- Main window: Vertical scroll via CustomTkinter ScrollableFrame.
- Output: Choose either:
  (A) Output Folder + File Name (default "Charge_Codes") with Drag & Drop, OR
  (B) Single Output File Path (.xlsx) with Drag & Drop
  We force .xlsx as the final extension.
- Threaded run; status log; full error stacktraces.
- Preferences saved to %APPDATA%/ChargeCodeGenerator/config.json

Dependencies:
- Required: pandas, openpyxl, customtkinter
- Optional (for drag & drop): tkinterdnd2

Run:
    pip install pandas openpyxl customtkinter
    pip install tkinterdnd2   # optional, for drag & drop
"""

import re
import pathlib
import threading
import traceback
import json
import csv
import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# --- CustomTkinter ---
import customtkinter as ctk

# Try to enable drag & drop (tkinterdnd2)
DND_AVAILABLE = False
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False

# --------------------------
# Core logic
# --------------------------

def _find_column(df, keywords):
    """Locate a column whose name contains any keyword (case-insensitive)."""
    for c in df.columns:
        cl = str(c).strip().lower()
        for kw in keywords:
            if kw.lower() in cl:
                return c
    return None

def _extract_wbs_code(wbs_text):
    """From WBS text (e.g., '04.03.01 - ...'), find first NN.NN.NN and return NNNNNN."""
    if not isinstance(wbs_text, str):
        return None
    m = re.search(r"\b(\d{2}\.\d{2}\.\d{2})\b", wbs_text)
    if not m:
        return None
    return m.group(1).replace(".", "")

def _extract_wbs_desc(wbs_text):
    """
    From WBS text (e.g., '02.01.02 - System Architecture'), extract the description part:
    • Looks for NN.NN.NN then takes what's after it, trimming leading separators (dash/en-dash/em-dash/colon).
    • Returns None if no description is found.
    """
    if not isinstance(wbs_text, str):
        return None
    s = wbs_text.strip()
    m = re.search(r"\b(\d{2}\.\d{2}\.\d{2})\b", s)
    if not m:
        return None
    rest = s[m.end():].strip()
    rest = re.sub(r"^\s*[-–—:]\s*", "", rest)
    return rest or None

def _map_wbs_code(code, wbs_code_map=None, enable_map=True):
    """Optionally map WBS code via user-defined mapping dict."""
    if code is None:
        return None
    c = str(code).strip()
    if enable_map and wbs_code_map:
        return wbs_code_map.get(c, c)
    return c

def _extract_clin(clin_text, clin_map=None, enable_map=True, pad_to_6=False):
    """
    Extract CLIN from text like 'CLIN 0001' -> '0001'.
    Optional behaviors:
    - pad_to_6: if True, any CLIN shorter than 6 chars is left-padded with zeros to length 6.
    - mapping: applies optional CLIN mapping (e.g., to lot codes).
    """
    if clin_text is None:
        return None
    s = str(clin_text).strip()
    m = re.search(r"\bCLIN\s*(\w+)\b", s, flags=re.IGNORECASE)
    code = m.group(1).strip() if m else (re.search(r"\b(\w+)\b", s).group(1).strip() if re.search(r"\b(\w+)\b", s) else s)
    code = code.strip()
    if pad_to_6 and len(code) < 6:
        code = code.rjust(6, '0')
    if enable_map and clin_map:
        return clin_map.get(code, code)
    return code

def _normalize_phase(phase_text, phase_map=None, enable_map=True):
    """Map phase text via user-defined phase_map; case-insensitive on keys (uppercased)."""
    if phase_text is None:
        return None
    p = str(phase_text).strip()
    if enable_map and phase_map:
        mapped = phase_map.get(p.upper())
        if mapped is not None:
            return mapped
    return p

def _normalize_wbs_desc(desc_text, wbs_desc_map=None, enable_map=True):
    """Map WBS description via user-defined mapping; case-insensitive on keys (uppercased)."""
    if desc_text is None:
        return None
    d = str(desc_text).strip()
    if enable_map and wbs_desc_map:
        mapped = wbs_desc_map.get(d.upper())
        if mapped is not None:
            return mapped
    return d

def read_boe_summary_single(input_path):
    """Read first sheet as strings; tolerate .xlsm."""
    try:
        import openpyxl  # ensure present
    except ImportError:
        raise ImportError("Missing 'openpyxl'. Install with: pip install --upgrade openpyxl")
    df = pd.read_excel(input_path, sheet_name=0, dtype=str, engine="openpyxl")
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    return df

def read_boe_summary_multi(input_path):
    """
    Accept either a file or a folder:
    - If file: read one.
    - If folder: read all .xlsx/.xlsm files and concat.
    """
    p = pathlib.Path(input_path)
    if p.is_file():
        return read_boe_summary_single(p)
    elif p.is_dir():
        files = sorted(list(p.glob("*.xlsx")) + list(p.glob("*.xlsm")))
        if not files:
            raise ValueError(f"No Excel files (*.xlsx, *.xlsm) found in folder: {p}")
        dfs = []
        for f in files:
            df = read_boe_summary_single(f)
            df["__SourceFile__"] = f.name
            dfs.append(df)
        out = pd.concat(dfs, ignore_index=True)
        out = out.astype(str)
        out.columns = [c.strip() if isinstance(c, str) else c for c in out.columns]
        return out
    else:
        raise ValueError(f"Input path is neither a file nor folder: {p}")

def build_paths(
    df, project_number, contract_type,
    clin_map=None, phase_map=None,
    wbs_code_map=None, wbs_desc_map=None,
    enable_clin_map=True, enable_phase_map=True,
    enable_wbs_code_map=True, enable_wbs_desc_map=True,
    pad_clin_to_6=False,
    contract_label_to_code=None,  # NEW: mapping "FFP"->"1", etc.
):
    """
    Build full paths (Level1..Level5) from the BOE Summary; include WBS Description & new metadata; sort & de-duplicate.
    Level 2 is dictated *per row* by 'Cost Type' (if present); otherwise defaults to user-selected contract_type (numeric).
    """
    # Required base columns:
    clin_col = _find_column(df, ["clin"])
    wbs_col = _find_column(df, ["wbs"])
    phase_col = _find_column(df, ["phase"])
    missing = [name for name, col in [("CLIN", clin_col), ("WBS", wbs_col), ("Phase", phase_col)] if col is None]
    if missing:
        raise ValueError(f"Missing columns: {', '.join(missing)}. Found: {list(df.columns)}")

    # Optional new columns (flexible matching, case-insensitive)
    start_col = _find_column(df, [
        "start date", "start", "start_dt", "startdate", "start-date",
        "period start", "start of period", "begin date", "begin"
    ])
    end_col   = _find_column(df, [
        "end date", "end", "end_dt", "enddate", "end-date",
        "finish date", "finish", "period end", "end of period"
    ])
    obs_col   = _find_column(df, [
        "obs", "organizational breakdown structure", "org breakdown", "org", "org code", "obs code", "obs id"
    ])
    recnre_col = _find_column(df, [
        "rec/nre", "rec nre", "recnre", "rec", "nre", "recurring", "non-recurring", "recurring/nonrecurring"
    ])
    costtype_col = _find_column(df, [
        "cost type", "costtype", "cost-type", "contract type", "ctype", "contract type per row"
    ])

    rows = []
    for _, r in df.iterrows():
        # Existing level extraction / mapping:
        level3 = _extract_clin(
            r.get(clin_col),
            clin_map=clin_map,
            enable_map=enable_clin_map,
            pad_to_6=pad_clin_to_6
        )
        raw_wbs_code = _extract_wbs_code(r.get(wbs_col))
        level4 = _map_wbs_code(raw_wbs_code, wbs_code_map=wbs_code_map, enable_map=enable_wbs_code_map)
        level5 = _normalize_phase(r.get(phase_col), phase_map=phase_map, enable_map=enable_phase_map)

        raw_desc = _extract_wbs_desc(r.get(wbs_col))
        wbs_desc_final = _normalize_wbs_desc(raw_desc, wbs_desc_map=wbs_desc_map, enable_map=enable_wbs_desc_map)

        if not (level3 and level4 and level5):
            continue

        level1 = str(project_number).strip()

        # --- Level 2 (numeric-only): choose via Cost Type when possible; otherwise fallback to user-selected numeric code ---
        level2_default = str(contract_type).strip() if contract_type else "1"
        level2 = level2_default
        if costtype_col is not None:
            ct_raw = r.get(costtype_col)
            if ct_raw is not None:
                ct = str(ct_raw).strip()
                if ct:
                    # Try mapping label -> code (e.g., "FFP" -> "1", "CRNF" -> "2")
                    key = ct.upper()
                    if contract_label_to_code and key in contract_label_to_code:
                        level2 = contract_label_to_code[key]
                    elif re.fullmatch(r"\d+", ct):
                        # If it's a number string (e.g., "2"), use it directly
                        level2 = ct
                    else:
                        # Unrecognized label -> keep default numeric code
                        level2 = level2_default

        # --- gather optional metadata ---
        start_date = str(r.get(start_col)).strip() if start_col is not None and r.get(start_col) is not None else ""
        end_date   = str(r.get(end_col)).strip()   if end_col   is not None and r.get(end_col)   is not None else ""
        obs_val    = str(r.get(obs_col)).strip()   if obs_col   is not None and r.get(obs_col)   is not None else ""
        recnre_val = str(r.get(recnre_col)).strip()if recnre_col is not None and r.get(recnre_col) is not None else ""
        cost_type  = str(r.get(costtype_col)).strip() if costtype_col is not None and r.get(costtype_col) is not None else ""

        rows.append({
            "Level 1": level1,
            "Level 2": level2,
            "Level 3": level3,
            "Level 4": level4,
            "Level 5": level5,
            "WBS Description": wbs_desc_final or "",
            # NEW metadata columns in the base paths frame:
            "Start Date": start_date,
            "End Date": end_date,
            "OBS": obs_val,
            "REC/NRE": recnre_val,
            "Cost Type": cost_type,
        })

    if not rows:
        raise ValueError("No valid rows found (CLIN/WBS/Phase missing).")

    paths = pd.DataFrame(rows).drop_duplicates()
    paths = paths.sort_values(by=["Level 1", "Level 2", "Level 3", "Level 4", "Level 5"]).reset_index(drop=True)
    return paths

def expand_hierarchy(paths):
    """
    Expand sorted full paths into hierarchical rows with partial parent lines and full lines.
    Adds WBS Description and new metadata columns; metadata is blank on partial rows, populated on full rows.
    """
    out = []
    last = {"Level 1": None, "Level 2": None, "Level 3": None, "Level 4": None, "Level 5": None}

    metadata_cols = ["WBS Description", "Start Date", "End Date", "OBS", "REC/NRE", "Cost Type"]

    def _row(levels_present, paths_row):
        vals = {lvl: paths_row[lvl] if lvl in levels_present else "" for lvl in ["Level 1", "Level 2", "Level 3", "Level 4", "Level 5"]}
        parts = [paths_row[lvl] for lvl in levels_present]
        vals["Charge String"] = ".".join(parts) + "."
        # metadata blank for partial rows:
        for mc in metadata_cols:
            vals[mc] = ""
        return vals

    for _, paths_row in paths.iterrows():
        if paths_row["Level 1"] != last["Level 1"]:
            out.append(_row(["Level 1"], paths_row))
            last["Level 1"] = paths_row["Level 1"]
            last["Level 2"] = last["Level 3"] = last["Level 4"] = last["Level 5"] = None

        if paths_row["Level 2"] != last["Level 2"]:
            out.append(_row(["Level 1", "Level 2"], paths_row))
            last["Level 2"] = paths_row["Level 2"]
            last["Level 3"] = last["Level 4"] = last["Level 5"] = None

        if paths_row["Level 3"] != last["Level 3"]:
            out.append(_row(["Level 1", "Level 2", "Level 3"], paths_row))
            last["Level 3"] = paths_row["Level 3"]
            last["Level 4"] = last["Level 5"] = None

        if paths_row["Level 4"] != last["Level 4"]:
            out.append(_row(["Level 1", "Level 2", "Level 3", "Level 4"], paths_row))
            last["Level 4"] = paths_row["Level 4"]
            last["Level 5"] = None

        full_vals = {lvl: paths_row[lvl] for lvl in ["Level 1", "Level 2", "Level 3", "Level 4", "Level 5"]}
        full_vals["Charge String"] = ".".join([paths_row[lvl] for lvl in ["Level 1", "Level 2", "Level 3", "Level 4", "Level 5"]])
        # metadata populated for full row:
        for mc in metadata_cols:
            full_vals[mc] = paths_row.get(mc, "")
        out.append(full_vals)
        last["Level 5"] = paths_row["Level 5"]

    df_out = pd.DataFrame(out).drop_duplicates().reset_index(drop=True)
    return df_out

def write_output_excel(df_out, output_path):
    output_path = pathlib.Path(output_path)
    if output_path.suffix.lower() == ".xlsm":
        output_path = output_path.with_suffix(".xlsx")
    with pd.ExcelWriter(output_path, engine="openpyxl") as xw:
        df_out.to_excel(xw, index=False, sheet_name="Charge Codes")
    return output_path

# --------------------------
# Theme-aware ttk styler
# --------------------------

class TtkStyler:
    """
    A tiny style adapter to make ttk widgets (Treeview/Scrollbars) match CustomTkinter appearance.
    Call .apply() whenever appearance changes. Assign "Custom.Treeview" to Treeviews.
    """
    def __init__(self):
        self.style = ttk.Style()

    def apply(self):
        # Force 'clam' so bg/fg configs are respected on most platforms.
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        mode = ctk.get_appearance_mode()  # "Dark" or "Light"
        if mode == "Dark":
            bg = "#2b2b2b"         # main bg
            fg = "#eeeeee"         # text
            tr_bg = "#3a3a3a"      # tree bg
            hd_bg = "#444444"      # header bg
            sel_bg = "#1f6aa5"     # CTk blue
            sel_fg = "#ffffff"
            sb_bg = "#404040"      # scrollbar bg
            trough = "#2b2b2b"
        else:
            bg = "#ebebeb"
            fg = "#000000"
            tr_bg = "#f5f5f5"
            hd_bg = "#d9d9d9"
            sel_bg = "#1f6aa5"
            sel_fg = "#ffffff"
            sb_bg = "#d0d0d0"
            trough = "#ebebeb"

        # Treeview rows & body
        self.style.configure(
            "Custom.Treeview",
            background=tr_bg,
            foreground=fg,
            fieldbackground=tr_bg,
            borderwidth=0,
            relief="flat",
            rowheight=24,
        )
        self.style.map(
            "Custom.Treeview",
            background=[("selected", sel_bg)],
            foreground=[("selected", sel_fg)],
        )

        # Header styling
        self.style.configure(
            "Custom.Treeview.Heading",
            background=hd_bg,
            foreground=fg,
            borderwidth=0,
            relief="flat",
        )
        self.style.map(
            "Custom.Treeview.Heading",
            background=[("active", hd_bg)],
            foreground=[("active", fg)]
        )

        # Scrollbars
        for orient in ("Vertical", "Horizontal"):
            name = f"{orient}.TScrollbar"
            self.style.configure(
                name,
                background=sb_bg,
                troughcolor=trough,
                bordercolor=trough,
                lightcolor=trough,
                darkcolor=trough,
            )

    def style_tree(self, tree: ttk.Treeview):
        """Apply our style and clean up borders."""
        try:
            tree.configure(style="Custom.Treeview")
            tree.configure(selectmode="extended")
        except Exception:
            pass


# --------------------------
# CustomTkinter GUI
# --------------------------

class MappingDialog(ctk.CTkToplevel):
    """Simple modal dialog to add/edit a mapping pair."""
    def __init__(self, parent, title, from_label="From", to_label="To", initial_from="", initial_to=""):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.result = None

        pad_x, pad_y = 8, 6
        frm = ctk.CTkFrame(self)
        frm.pack(fill="both", expand=True, padx=pad_x, pady=pad_y)

        ctk.CTkLabel(frm, text=from_label + ":").grid(row=0, column=0, sticky="w", padx=pad_x, pady=pad_y)
        self.var_from = tk.StringVar(value=initial_from)
        ctk.CTkEntry(frm, textvariable=self.var_from, width=280).grid(row=0, column=1, sticky="w", padx=pad_x, pady=pad_y)

        ctk.CTkLabel(frm, text=to_label + ":").grid(row=1, column=0, sticky="w", padx=pad_x, pady=pad_y)
        self.var_to = tk.StringVar(value=initial_to)
        ctk.CTkEntry(frm, textvariable=self.var_to, width=280).grid(row=1, column=1, sticky="w", padx=pad_x, pady=pad_y)

        btns = ctk.CTkFrame(frm)
        btns.grid(row=2, column=0, columnspan=2, sticky="e", padx=pad_x, pady=pad_y)
        ctk.CTkButton(btns, text="Cancel", command=self._cancel).pack(side="right", padx=5)
        ctk.CTkButton(btns, text="OK", command=self._ok).pack(side="right", padx=5)

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_visibility()
        self.focus()
        self.wait_window(self)

    def _ok(self):
        from_val = self.var_from.get().strip()
        to_val = self.var_to.get().strip()
        if not from_val:
            messagebox.showerror("Validation", "The 'From' value cannot be empty.")
            return
        if not to_val:
            messagebox.showerror("Validation", "The 'To' value cannot be empty.")
            return
        self.result = (from_val, to_val)
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()

class ContractTypesEditor(ctk.CTkToplevel):
    """Modal editor for Contract Type list: items like '1: FFP', '2: FPI', '3: CPFF', '4: CPIF' """
    def __init__(self, parent, contract_items):
        super().__init__(parent)
        self.title("Edit Contract Types (Level 2)")
        self.resizable(True, True)
        self.result = None

        pad_x, pad_y = 8, 6
        frm = ctk.CTkFrame(self)
        frm.pack(fill="both", expand=True, padx=pad_x, pady=pad_y)

        ctk.CTkLabel(frm, text="Contract Types (Code / Label):").grid(row=0, column=0, sticky="w", padx=pad_x, pady=pad_y)

        # Tree + Scrollbars
        self.tree = ttk.Treeview(frm, columns=("Code", "Label"), show="headings", height=10)
        self.tree.heading("Code", text="Code")
        self.tree.heading("Label", text="Label")
        self.tree.column("Code", width=120, anchor="center")
        self.tree.column("Label", width=260, anchor="center")
        self.tree.grid(row=1, column=0, sticky="nsew", padx=pad_x, pady=pad_y)

        yscroll = ttk.Scrollbar(frm, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=1, column=1, sticky="ns")
        xscroll = ttk.Scrollbar(frm, orient="horizontal", command=self.tree.xview)
        xscroll.grid(row=2, column=0, sticky="ew", padx=pad_x, pady=pad_y)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        for item in contract_items:
            code, label = self._split_item(item)
            self.tree.insert("", tk.END, values=(code, label))

        btns = ctk.CTkFrame(frm)
        btns.grid(row=3, column=0, sticky="w", padx=pad_x, pady=pad_y)
        ctk.CTkButton(btns, text="Add", command=self.add_item).pack(side="left", padx=4)
        ctk.CTkButton(btns, text="Edit", command=self.edit_item).pack(side="left", padx=4)
        ctk.CTkButton(btns, text="Delete", command=self.delete_item).pack(side="left", padx=4)
        ctk.CTkButton(btns, text="Import…", command=self.import_items).pack(side="left", padx=4)
        ctk.CTkButton(btns, text="Export…", command=self.export_items).pack(side="left", padx=4)

        action = ctk.CTkFrame(frm)
        action.grid(row=4, column=0, sticky="e", padx=pad_x, pady=pad_y)
        ctk.CTkButton(action, text="Cancel", command=self._cancel).pack(side="right", padx=5)
        ctk.CTkButton(action, text="OK", command=self._ok).pack(side="right", padx=5)

        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(1, weight=1)

        # Make ttk tree adopt theme
        self._styler = TtkStyler()
        self._styler.apply()
        self._styler.style_tree(self.tree)

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_visibility()
        self.focus()
        self.wait_window(self)

    def _split_item(self, s):
        s = str(s).strip()
        if ":" in s:
            code, label = s.split(":", 1)
            return code.strip(), label.strip()
        return s, ""

    def add_item(self):
        dlg = MappingDialog(self, "Add Contract Type", from_label="Code", to_label="Label")
        if dlg.result:
            code, label = dlg.result
            code = code.strip()
            label = label.strip()
            if not code or not label:
                messagebox.showerror("Validation", "Both Code and Label are required.")
                return
            for iid in self.tree.get_children():
                c, _ = self.tree.item(iid, "values")
                if c == code:
                    self.tree.item(iid, values=(code, label))
                    return
            self.tree.insert("", tk.END, values=(code, label))

    def edit_item(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Edit", "Select one row to edit.")
            return
        iid = sel[0]
        code, label = self.tree.item(iid, "values")
        dlg = MappingDialog(self, "Edit Contract Type", from_label="Code", to_label="Label",
                            initial_from=code, initial_to=label)
        if dlg.result:
            code2, label2 = dlg.result
            code2 = code2.strip()
            label2 = label2.strip()
            if not code2 or not label2:
                messagebox.showerror("Validation", "Both Code and Label are required.")
                return
            self.tree.item(iid, values=(code2, label2))

    def delete_item(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Delete", "Select one or more rows to delete.")
            return
        for iid in sel:
            self.tree.delete(iid)

    def import_items(self):
        path = filedialog.askopenfilename(
            title="Import Contract Types (CSV or JSON)",
            filetypes=[("CSV Files", "*.csv"), ("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        if not path:
            return
        p = pathlib.Path(path)
        try:
            if p.suffix.lower() == ".json":
                with open(p, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    items = [(str(k), str(v)) for k, v in data.items()]
                elif isinstance(data, list):
                    items = []
                    for row in data:
                        if isinstance(row, dict):
                            items.append((str(row.get("Code", "")).strip(), str(row.get("Label", "")).strip()))
                        else:
                            c, l = self._split_item(row)
                            items.append((c, l))
                else:
                    raise ValueError("JSON must be an object {code: label} or a list.")
            else:
                with open(p, "r", encoding="utf-8-sig", newline="") as f:
                    reader = csv.reader(f)
                    rows = list(reader)
                start_idx = 1 if rows and rows[0] and rows[0][0].lower().startswith("code") else 0
                items = []
                for row in rows[start_idx:]:
                    if len(row) < 2:
                        continue
                    items.append((row[0].strip(), row[1].strip()))
            for iid in self.tree.get_children():
                self.tree.delete(iid)
            for code, label in items:
                if code and label:
                    self.tree.insert("", tk.END, values=(code, label))
            messagebox.showinfo("Import", f"Imported contract types from:\n{path}")
        except Exception as ex:
            messagebox.showerror("Import Error", f"Failed to import:\n{ex}")

    def export_items(self):
        path = filedialog.asksaveasfilename(
            title="Export Contract Types",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("JSON Files", "*.json")]
        )
        if not path:
            return
        p = pathlib.Path(path)
        items = []
        for iid in self.tree.get_children():
            code, label = self.tree.item(iid, "values")
            items.append((code, label))
        try:
            if p.suffix.lower() == ".json":
                obj = {code: label for code, label in items}
                with open(p, "w", encoding="utf-8") as f:
                    json.dump(obj, f, indent=2)
            else:
                with open(p, "w", encoding="utf-8", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(["Code", "Label"])
                    for code, label in items:
                        writer.writerow([code, label])
            messagebox.showinfo("Export", f"Exported contract types to:\n{path}")
        except Exception as ex:
            messagebox.showerror("Export Error", f"Failed to export:\n{ex}")

    def _ok(self):
        items = []
        for iid in self.tree.get_children():
            code, label = self.tree.item(iid, "values")
            code = str(code).strip()
            label = str(label).strip()
            if code and label:
                items.append(f"{code}: {label}")
        if not items:
            if not messagebox.askyesno("Empty List", "No contract types defined. Continue with empty list?"):
                return
        self.result = items
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()

# ---------- Accordion ----------

class AccordionSection(ctk.CTkFrame):
    """A simple expandable/collapsible section."""
    def __init__(self, parent, title: str, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self._expanded = True
        self._title = title

        self.header = ctk.CTkButton(
            self, text=f"▼ {title}", anchor="w",
            fg_color="transparent", hover=False, command=self.toggle
        )
        self.header.grid(row=0, column=0, sticky="ew", padx=4, pady=(2, 0))

        self.content = ctk.CTkFrame(self)
        self.content.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 4))

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

    def toggle(self):
        self._expanded = not self._expanded
        self.header.configure(text=("▼ " if self._expanded else "▶ ") + self._title)
        if self._expanded:
            self.content.grid()
        else:
            self.content.grid_remove()

class ChargeCodesGUI:
    SPECIAL_CLIN_PAD_KEY = "__PAD_TO_6__"  # internal marker for the rule

    def __init__(self, root):
        self.root = root
        self.root.title("Charge Code Generator (BOE Summary → Hierarchy) — v2.2.0")

        # Theme & scaling
        ctk.set_appearance_mode("System")  # "Light", "Dark", or "System"
        ctk.set_default_color_theme("blue")
        ctk.set_widget_scaling(1.0)
        ctk.set_window_scaling(1.0)

        # Ttk styler for Treeviews/Scrollbars
        self._styler = TtkStyler()
        self._styler.apply()

        # Config paths
        appdata = os.environ.get('APPDATA', str(pathlib.Path.home()))
        self.config_dir = pathlib.Path(appdata) / 'ChargeCodeGenerator'
        self.config_path = self.config_dir / 'config.json'

        # Vars
        self.input_path = tk.StringVar()
        self.project_number = tk.StringVar()
        self.contract_type = tk.StringVar()

        self.output_mode = tk.StringVar(value="folder")  # 'folder' or 'file'
        self.output_dir = tk.StringVar()
        self.output_filename = tk.StringVar(value="Charge_Codes")
        self.output_file_path = tk.StringVar()

        # Level 2 defaults
        self.contract_options = ["1: FFP", "2: FPI", "3: CPFF", "4: CPIF", "5: FPAF", "7: T&M", "8: CPAF"]

        # Mapping toggles
        self.use_clin_map = tk.BooleanVar(value=True)
        self.use_phase_map = tk.BooleanVar(value=True)
        self.use_wbs_code_map = tk.BooleanVar(value=True)
        self.use_wbs_desc_map = tk.BooleanVar(value=True)

        # Mapping datasets
        self.clin_data = []       # {'selected':bool,'from':str,'to':str}
        self.phase_data = []      # keys uppercased on display
        self.wbs_code_data = []
        self.wbs_desc_data = []

        # Search vars
        self.clin_search = tk.StringVar()
        self.phase_search = tk.StringVar()
        self.wbs_code_search = tk.StringVar()
        self.wbs_desc_search = tk.StringVar()

        # Top bar (theme switch)
        topbar = ctk.CTkFrame(self.root)
        topbar.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 0))
        ctk.CTkLabel(topbar, text="Appearance:").pack(side="left", padx=(4, 4))
        self.theme_opt = ctk.CTkOptionMenu(topbar, values=["System", "Light", "Dark"], command=self._on_theme_change)
        self.theme_opt.set("System")
        self.theme_opt.pack(side="left")

        # Main scrollable content
        scrollable = ctk.CTkScrollableFrame(self.root)
        scrollable.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        frm = scrollable  # place all widgets inside this scrollable frame

        # --- Input row ---
        ctk.CTkLabel(frm, text="Input (Excel file or folder):").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        self.input_entry = ctk.CTkEntry(frm, textvariable=self.input_path, width=500)
        self.input_entry.grid(row=0, column=1, sticky="ew", padx=8, pady=6)
        ctk.CTkButton(frm, text="Browse File…", command=self.browse_input_file).grid(row=0, column=2, padx=4, pady=6)
        ctk.CTkButton(frm, text="Browse Folder…", command=self.browse_input_folder).grid(row=0, column=3, padx=4, pady=6)

        if DND_AVAILABLE:
            self._safe_register_dnd(self.input_entry, self._on_drop_input)
        else:
            ctk.CTkLabel(frm, text="(Drag & drop available if 'tkinterdnd2' is installed)").grid(row=1, column=1, sticky="w", padx=8, pady=4)

        # --- Project Number ---
        ctk.CTkLabel(frm, text="Project Number (Level 1):").grid(row=2, column=0, sticky="w", padx=8, pady=6)
        ctk.CTkEntry(frm, textvariable=self.project_number, width=240).grid(row=2, column=1, sticky="w", padx=8, pady=6)

        # --- Contract Type ---
        ctk.CTkLabel(frm, text="Contract Type (Level 2 default):").grid(row=3, column=0, sticky="w", padx=8, pady=6)
        self.contract_combo = ctk.CTkComboBox(frm, values=self.contract_options, variable=self.contract_type, width=240)
        self.contract_combo.grid(row=3, column=1, sticky="w", padx=8, pady=6)
        if self.contract_options:
            self.contract_combo.set(self.contract_options[0])
        ctk.CTkButton(frm, text="Edit…", command=self.edit_contract_types).grid(row=3, column=2, padx=4, pady=6)

        # --- Output Mode ---
        ctk.CTkLabel(frm, text="Output Mode:").grid(row=4, column=0, sticky="w", padx=8, pady=6)
        output_mode_frame = ctk.CTkFrame(frm)
        output_mode_frame.grid(row=4, column=1, sticky="w", padx=8, pady=6)
        rb_folder = ctk.CTkRadioButton(output_mode_frame, text="Folder + File Name", variable=self.output_mode, value="folder", command=self._on_output_mode_change)
        rb_file = ctk.CTkRadioButton(output_mode_frame, text="Single File Path", variable=self.output_mode, value="file", command=self._on_output_mode_change)
        rb_folder.pack(side="left", padx=(0, 12))
        rb_file.pack(side="left")

        # --- Output rows (folder mode default) ---
        self.row_output_folder_label = ctk.CTkLabel(frm, text="Output Folder:")
        self.row_output_folder_label.grid(row=5, column=0, sticky="w", padx=8, pady=6)
        self.row_output_folder_entry = ctk.CTkEntry(frm, textvariable=self.output_dir, width=500)
        self.row_output_folder_entry.grid(row=5, column=1, sticky="ew", padx=8, pady=6)
        ctk.CTkButton(frm, text="Browse…", command=self.browse_output_dir).grid(row=5, column=2, padx=4, pady=6)
        if DND_AVAILABLE:
            self._safe_register_dnd(self.row_output_folder_entry, self._on_drop_output_folder)

        self.row_output_name_label = ctk.CTkLabel(frm, text="Output File Name:")
        self.row_output_name_label.grid(row=6, column=0, sticky="w", padx=8, pady=6)
        self.row_output_name_entry = ctk.CTkEntry(frm, textvariable=self.output_filename, width=240)
        self.row_output_name_entry.grid(row=6, column=1, sticky="w", padx=8, pady=6)
        ctk.CTkLabel(frm, text="(default: Charge_Codes; .xlsx will be added)").grid(row=6, column=2, sticky="w", padx=4, pady=6)

        # -- Single output path (hidden initially) --
        self.row_output_file_label = ctk.CTkLabel(frm, text="Output File (.xlsx):")
        self.row_output_file_entry = ctk.CTkEntry(frm, textvariable=self.output_file_path, width=500)
        self.row_output_file_browse = ctk.CTkButton(frm, text="Browse…", command=self.browse_output_file)
        if DND_AVAILABLE:
            self._safe_register_dnd(self.row_output_file_entry, self._on_drop_output_file)

        # ---------------- Mapping Editors (Accordion) ----------------
        editors = ctk.CTkFrame(frm)
        editors.grid(row=8, column=0, columnspan=4, sticky="nsew", padx=8, pady=8)
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(10, weight=1)
        editors.columnconfigure(0, weight=1)
        editors.rowconfigure(0, weight=1)

        # ---- CLIN Mapping ----
        self.acc_clin = AccordionSection(editors, "CLIN (Contract) Mapping")
        self.acc_clin.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        top_clin = ctk.CTkFrame(self.acc_clin.content)
        top_clin.grid(row=0, column=0, sticky="ew", padx=4, pady=4)
        ctk.CTkCheckBox(top_clin, text="Apply CLIN mapping", variable=self.use_clin_map).pack(side="left")
        ctk.CTkLabel(top_clin, text="Search:").pack(side="left", padx=(12, 4))
        ctk.CTkEntry(top_clin, textvariable=self.clin_search, width=180).pack(side="left")

        self.clin_tree = self._create_selectable_mapping_tree(self.acc_clin.content, start_row=1, columns=("✓", "From", "To"))
        self._styler.style_tree(self.clin_tree)
        self._seed_clin_defaults()
        self._refresh_mapping_tree("clin")

        clin_btns = ctk.CTkFrame(self.acc_clin.content)
        clin_btns.grid(row=3, column=0, sticky="w", padx=4, pady=4)
        ctk.CTkButton(clin_btns, text="Add", command=self.add_clin).pack(side="left", padx=4)
        ctk.CTkButton(clin_btns, text="Delete", command=self.delete_clin).pack(side="left", padx=4)
        ctk.CTkButton(clin_btns, text="Select All", command=lambda: self.select_all("clin")).pack(side="left", padx=4)
        ctk.CTkButton(clin_btns, text="Clear All", command=lambda: self.clear_all("clin")).pack(side="left", padx=4)
        ctk.CTkButton(clin_btns, text="Import…", command=lambda: self.import_mapping("clin")).pack(side="left", padx=4)
        ctk.CTkButton(clin_btns, text="Export…", command=lambda: self.export_mapping("clin")).pack(side="left", padx=4)

        # ---- Phase Mapping ----
        self.acc_phase = AccordionSection(editors, "Phase Mapping")
        self.acc_phase.grid(row=1, column=0, sticky="nsew", padx=4, pady=4)
        top_phase = ctk.CTkFrame(self.acc_phase.content)
        top_phase.grid(row=0, column=0, sticky="ew", padx=4, pady=4)
        ctk.CTkCheckBox(top_phase, text="Apply Phase mapping", variable=self.use_phase_map).pack(side="left")
        ctk.CTkLabel(top_phase, text="Search:").pack(side="left", padx=(12, 4))
        ctk.CTkEntry(top_phase, textvariable=self.phase_search, width=180).pack(side="left")

        self.phase_tree = self._create_selectable_mapping_tree(self.acc_phase.content, start_row=1, columns=("✓", "From (Upper)", "To"))
        self._styler.style_tree(self.phase_tree)
        self._seed_phase_defaults()
        self._refresh_mapping_tree("phase")

        phase_btns = ctk.CTkFrame(self.acc_phase.content)
        phase_btns.grid(row=3, column=0, sticky="w", padx=4, pady=4)
        ctk.CTkButton(phase_btns, text="Add", command=self.add_phase).pack(side="left", padx=4)
        ctk.CTkButton(phase_btns, text="Delete", command=self.delete_phase).pack(side="left", padx=4)
        ctk.CTkButton(phase_btns, text="Select All", command=lambda: self.select_all("phase")).pack(side="left", padx=4)
        ctk.CTkButton(phase_btns, text="Clear All", command=lambda: self.clear_all("phase")).pack(side="left", padx=4)
        ctk.CTkButton(phase_btns, text="Import…", command=lambda: self.import_mapping("phase")).pack(side="left", padx=4)
        ctk.CTkButton(phase_btns, text="Export…", command=lambda: self.export_mapping("phase")).pack(side="left", padx=4)

        # ---- WBS Code Mapping ----
        self.acc_wbs_code = AccordionSection(editors, "WBS Code Mapping")
        self.acc_wbs_code.grid(row=2, column=0, sticky="nsew", padx=4, pady=4)
        top_wbs_code = ctk.CTkFrame(self.acc_wbs_code.content)
        top_wbs_code.grid(row=0, column=0, sticky="ew", padx=4, pady=4)
        ctk.CTkCheckBox(top_wbs_code, text="Apply WBS code mapping", variable=self.use_wbs_code_map).pack(side="left")
        ctk.CTkLabel(top_wbs_code, text="Search:").pack(side="left", padx=(12, 4))
        ctk.CTkEntry(top_wbs_code, textvariable=self.wbs_code_search, width=180).pack(side="left")

        self.wbs_code_tree = self._create_selectable_mapping_tree(self.acc_wbs_code.content, start_row=1, columns=("✓", "From (WBS Code)", "To (Mapped Code)"))
        self._styler.style_tree(self.wbs_code_tree)
        self._seed_wbs_code_defaults()
        self._refresh_mapping_tree("wbs_code")

        wbs_code_btns = ctk.CTkFrame(self.acc_wbs_code.content)
        wbs_code_btns.grid(row=3, column=0, sticky="w", padx=4, pady=4)
        ctk.CTkButton(wbs_code_btns, text="Add", command=self.add_wbs_code).pack(side="left", padx=4)
        ctk.CTkButton(wbs_code_btns, text="Delete", command=self.delete_wbs_code).pack(side="left", padx=4)
        ctk.CTkButton(wbs_code_btns, text="Select All", command=lambda: self.select_all("wbs_code")).pack(side="left", padx=4)
        ctk.CTkButton(wbs_code_btns, text="Clear All", command=lambda: self.clear_all("wbs_code")).pack(side="left", padx=4)
        ctk.CTkButton(wbs_code_btns, text="Import…", command=lambda: self.import_mapping("wbs_code")).pack(side="left", padx=4)
        ctk.CTkButton(wbs_code_btns, text="Export…", command=lambda: self.export_mapping("wbs_code")).pack(side="left", padx=4)

        # ---- WBS Description Mapping ----
        self.acc_wbs_desc = AccordionSection(editors, "WBS Description Mapping")
        self.acc_wbs_desc.grid(row=3, column=0, sticky="nsew", padx=4, pady=4)
        top_wbs_desc = ctk.CTkFrame(self.acc_wbs_desc.content)
        top_wbs_desc.grid(row=0, column=0, sticky="ew", padx=4, pady=4)
        ctk.CTkCheckBox(top_wbs_desc, text="Apply WBS description mapping", variable=self.use_wbs_desc_map).pack(side="left")
        ctk.CTkLabel(top_wbs_desc, text="Search:").pack(side="left", padx=(12, 4))
        ctk.CTkEntry(top_wbs_desc, textvariable=self.wbs_desc_search, width=180).pack(side="left")

        self.wbs_desc_tree = self._create_selectable_mapping_tree(self.acc_wbs_desc.content, start_row=1, columns=("✓", "From (Upper)", "To (Mapped Description)"))
        self._styler.style_tree(self.wbs_desc_tree)
        self._seed_wbs_desc_defaults()
        self._refresh_mapping_tree("wbs_desc")

        wbs_desc_btns = ctk.CTkFrame(self.acc_wbs_desc.content)
        wbs_desc_btns.grid(row=3, column=0, sticky="w", padx=4, pady=4)
        ctk.CTkButton(wbs_desc_btns, text="Add", command=self.add_wbs_desc).pack(side="left", padx=4)
        ctk.CTkButton(wbs_desc_btns, text="Delete", command=self.delete_wbs_desc).pack(side="left", padx=4)
        ctk.CTkButton(wbs_desc_btns, text="Select All", command=lambda: self.select_all("wbs_desc")).pack(side="left", padx=4)
        ctk.CTkButton(wbs_desc_btns, text="Clear All", command=lambda: self.clear_all("wbs_desc")).pack(side="left", padx=4)
        ctk.CTkButton(wbs_desc_btns, text="Import…", command=lambda: self.import_mapping("wbs_desc")).pack(side="left", padx=4)
        ctk.CTkButton(wbs_desc_btns, text="Export…", command=lambda: self.export_mapping("wbs_desc")).pack(side="left", padx=4)

        # --- Load user preferences (after trees created) ---
        self.load_config()

        # --- Actions ---
        actions = ctk.CTkFrame(frm)
        actions.grid(row=9, column=0, columnspan=4, sticky="ew", padx=8, pady=8)
        self.btn_run = ctk.CTkButton(actions, text="Run", command=self.run_clicked)
        self.btn_run.pack(side="left", padx=6)
        self.btn_save = ctk.CTkButton(actions, text="Save Preferences", command=self.save_preferences)
        self.btn_save.pack(side="left", padx=6)
        self.btn_reset = ctk.CTkButton(actions, text="Reset to Defaults", command=self.reset_to_defaults)
        self.btn_reset.pack(side="left", padx=6)

        # --- Status ---
        ctk.CTkLabel(frm, text="Status:").grid(row=10, column=0, sticky="nw", padx=8, pady=6)
        self.status_box = ctk.CTkTextbox(frm, height=200, width=700)
        self.status_box.grid(row=10, column=1, columnspan=3, sticky="nsew", padx=8, pady=6)

        # Expandable layout
        for col in (0, 1, 2, 3):
            frm.columnconfigure(col, weight=(1 if col == 1 else 0))
        frm.rowconfigure(10, weight=1)

        # Search live filters
        self.clin_search.trace_add("write", lambda *_: self._refresh_mapping_tree("clin"))
        self.phase_search.trace_add("write", lambda *_: self._refresh_mapping_tree("phase"))
        self.wbs_code_search.trace_add("write", lambda *_: self._refresh_mapping_tree("wbs_code"))
        self.wbs_desc_search.trace_add("write", lambda *_: self._refresh_mapping_tree("wbs_desc"))

        # Toggle checkbox by click
        self.clin_tree.bind("<Button-1>", lambda e: self._on_tree_click(e, "clin"))
        self.phase_tree.bind("<Button-1>", lambda e: self._on_tree_click(e, "phase"))
        self.wbs_code_tree.bind("<Button-1>", lambda e: self._on_tree_click(e, "wbs_code"))
        self.wbs_desc_tree.bind("<Button-1>", lambda e: self._on_tree_click(e, "wbs_desc"))

        # Init output mode UI
        self._on_output_mode_change()

    # ---------- Appearance ----------
    def _on_theme_change(self, mode: str):
        try:
            ctk.set_appearance_mode(mode)
        except Exception:
            pass
        # Reapply ttk styles so Treeviews & Scrollbars match new mode
        self._styler.apply()
        # Explicitly restyle all trees (useful if new Style objects)
        for tree in (self.clin_tree, self.phase_tree, self.wbs_code_tree, self.wbs_desc_tree):
            self._styler.style_tree(tree)

    # ---------- Drag & Drop handlers ----------
    def _safe_register_dnd(self, widget, drop_callback):
        """Attempt to register DnD on widget; fallback silently if not possible."""
        if not DND_AVAILABLE:
            return
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", drop_callback)
        except Exception:
            # Some CustomTkinter widgets may not expose dnd methods; ignore gracefully.
            pass

    def _normalize_dropped_paths(self, data: str):
        """Convert tkdnd drop data to a list of paths (handles braces & spaces)."""
        data = data.strip()
        paths, buf, brace = [], "", False
        for ch in data:
            if ch == "{":
                brace = True
                buf += ch
            elif ch == "}":
                brace = False
                buf += ch
                paths.append(buf.strip("{}"))
                buf = ""
            elif ch == " " and not brace:
                if buf:
                    paths.append(buf)
                    buf = ""
            else:
                buf += ch
        if buf:
            paths.append(buf)
        return paths

    def _on_drop_input(self, event):
        paths = self._normalize_dropped_paths(event.data)
        if paths:
            self.input_path.set(paths[0])

    def _on_drop_output_folder(self, event):
        paths = self._normalize_dropped_paths(event.data)
        if not paths:
            return
        p = pathlib.Path(paths[0])
        if p.is_file():
            p = p.parent
        if p.exists():
            self.output_dir.set(str(p))

    def _on_drop_output_file(self, event):
        paths = self._normalize_dropped_paths(event.data)
        if not paths:
            return
        p = pathlib.Path(paths[0])
        if p.is_dir():
            fname = self._sanitize_filename(self.output_filename.get().strip() or "Charge_Codes")
            fp = (p / fname).with_suffix(".xlsx")
            self.output_file_path.set(str(fp))
            return
        if p.suffix.lower() != ".xlsx":
            p = p.with_suffix(".xlsx")
        if p.parent.exists():
            self.output_file_path.set(str(p))
        else:
            messagebox.showerror("Invalid Path", f"Folder does not exist:\n{p.parent}")

    # ---------- Contract types ----------
    def edit_contract_types(self):
        editor = ContractTypesEditor(self.root, self.contract_options)
        if editor.result is not None:
            self.contract_options = editor.result
            self.contract_combo.configure(values=self.contract_options)
            current = self.contract_type.get()
            if current in self.contract_options:
                self.contract_combo.set(current)
            elif self.contract_options:
                self.contract_combo.set(self.contract_options[0])

    def _parse_contract_code(self, selected_text):
        """Parse 'code: label' -> 'code'. If no colon, return as-is."""
        s = str(selected_text).strip()
        if ":" in s:
            code, _ = s.split(":", 1)
            return code.strip()
        return s

    def _contract_label_to_code_map(self):
        """
        Build a {LABEL_UPPER: CODE} map from self.contract_options that look like 'CODE: LABEL'.
        Example: ['1: FFP', '2: CRNF'] -> {'FFP': '1', 'CRNF': '2'}
        """
        out = {}
        for item in self.contract_options:
            s = str(item).strip()
            if ":" in s:
                code, label = s.split(":", 1)
                code = code.strip()
                label = label.strip()
                if label:
                    out[label.upper()] = code
        return out

    # ---------- Output mode UI ----------
    def _on_output_mode_change(self):
        mode = self.output_mode.get()
        if mode == "folder":
            self.row_output_folder_label.grid()
            self.row_output_folder_entry.grid()
            self.row_output_name_label.grid()
            self.row_output_name_entry.grid()
            self.row_output_file_label.grid_remove()
            self.row_output_file_entry.grid_remove()
            self.row_output_file_browse.grid_remove()
        else:
            self.row_output_folder_label.grid_remove()
            self.row_output_folder_entry.grid_remove()
            self.row_output_name_label.grid_remove()
            self.row_output_name_entry.grid_remove()
            self.row_output_file_label.grid(row=7, column=0, sticky="w", padx=8, pady=6)
            self.row_output_file_entry.grid(row=7, column=1, sticky="ew", padx=8, pady=6)
            self.row_output_file_browse.grid(row=7, column=2, padx=8, pady=6)
            self._safe_register_dnd(self.row_output_file_entry, self._on_drop_output_file)

    # ---------- Tree creation with scrollbars ----------
    def _create_selectable_mapping_tree(self, parent, start_row=0, columns=("✓", "From", "To")):
        tree_container = ctk.CTkFrame(parent)
        tree_container.grid(row=start_row, column=0, sticky="nsew", padx=4, pady=4)

        tree = ttk.Treeview(tree_container, columns=columns, show="headings", height=10)
        for idx, col in enumerate(columns):
            tree.heading(col, text=col)
            width = 120 if idx == 0 else 260
            tree.column(col, width=width, anchor="center")
        tree.grid(row=0, column=0, sticky="nsew")

        yscroll = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        xscroll.grid(row=start_row + 1, column=0, sticky="ew", padx=4, pady=(0, 4))

        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)
        parent.rowconfigure(start_row, weight=1)
        parent.columnconfigure(0, weight=1)
        return tree

    # ---------- Mapping data management ----------
    def _seed_clin_defaults(self):
        self.clin_data = [
            {"selected": False, "from": self.SPECIAL_CLIN_PAD_KEY, "to": "Add chars until 6 chars"},
        ]

    def _seed_phase_defaults(self):
        self.phase_data = [
            {"selected": True, "from": "BID", "to": "1BID"},
            {"selected": True, "from": "LOE", "to": "ZLOE"},
            {"selected": True, "from": "RAV", "to": "ZRAV"},
            {"selected": True, "from": "PDV", "to": "ZPDV"},
            {"selected": True, "from": "DDV", "to": "ZDDV"},
            {"selected": True, "from": "COD", "to": "ZCOD"},
            {"selected": True, "from": "HSI", "to": "ZHSI"},
            {"selected": True, "from": "CVT", "to": "ZCVT"},
            {"selected": True, "from": "GAT", "to": "ZGAT"},
            {"selected": True, "from": "PROD", "to": "PROD"},
        ]

    def _seed_wbs_code_defaults(self):
        self.wbs_code_data = []

    def _seed_wbs_desc_defaults(self):
        self.wbs_desc_data = []

    def _filter_rows(self, dataset, query, upper_keys=False):
        q = query.strip().lower()
        out = []
        for row in dataset:
            f = row["from"]
            t = row["to"]
            f_cmp = (f.upper() if upper_keys else f).lower()
            t_cmp = t.lower()
            if not q or (q in f_cmp or q in t_cmp):
                out.append(row)
        return out

    def _refresh_mapping_tree(self, which):
        if which == "clin":
            tree = self.clin_tree
            data = self._filter_rows(self.clin_data, self.clin_search.get(), upper_keys=False)
        elif which == "phase":
            tree = self.phase_tree
            data = self._filter_rows(self.phase_data, self.phase_search.get(), upper_keys=True)
        elif which == "wbs_code":
            tree = self.wbs_code_tree
            data = self._filter_rows(self.wbs_code_data, self.wbs_code_search.get(), upper_keys=False)
        else:  # wbs_desc
            tree = self.wbs_desc_tree
            data = self._filter_rows(self.wbs_desc_data, self.wbs_desc_search.get(), upper_keys=True)

        for iid in tree.get_children():
            tree.delete(iid)
        tree._row_refs = {}
        for row in data:
            mark = "☑" if row["selected"] else "☐"
            if which in ("phase", "wbs_desc"):
                from_display = row["from"].upper()
                to_display = row["to"]
            else:
                if which == "clin" and row["from"] == self.SPECIAL_CLIN_PAD_KEY:
                    from_display = "CLIN w/less than 6 characters"
                    to_display = row["to"]
                else:
                    from_display = row["from"]
                    to_display = row["to"]
            iid = tree.insert("", tk.END, values=(mark, from_display, to_display))
            tree._row_refs[iid] = row

    def _on_tree_click(self, event, which):
        tree = {
            "clin": self.clin_tree,
            "phase": self.phase_tree,
            "wbs_code": self.wbs_code_tree,
            "wbs_desc": self.wbs_desc_tree,
        }[which]
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = tree.identify_column(event.x)
        if col != "#1":  # only checkbox column
            return
        iid = tree.identify_row(event.y)
        if not iid:
            return
        row = getattr(tree, "_row_refs", {}).get(iid)
        if not row:
            return
        row["selected"] = not row["selected"]
        mark = "☑" if row["selected"] else "☐"
        vals = list(tree.item(iid, "values"))
        vals[0] = mark
        tree.item(iid, values=tuple(vals))

    # ---- CLIN mapping CRUD ----
    def add_clin(self):
        dlg = MappingDialog(self.root, "Add CLIN Mapping", from_label="From (CLIN code)", to_label="To (Lot code)")
        if dlg.result:
            k, v = dlg.result
            k = k.strip()
            v = v.strip()
            if not k or not v:
                messagebox.showerror("Validation", "Both 'From' and 'To' are required.")
                return
            if k == self.SPECIAL_CLIN_PAD_KEY:
                messagebox.showerror("Validation", "The special rule key is reserved by the system.")
                return
            for row in self.clin_data:
                if row["from"] == k:
                    row["to"] = v
                    row["selected"] = True
                    self._refresh_mapping_tree("clin")
                    return
            self.clin_data.append({"selected": True, "from": k, "to": v})
            self._refresh_mapping_tree("clin")

    def delete_clin(self):
        sel = self.clin_tree.selection()
        if not sel:
            messagebox.showinfo("Delete", "Select one or more rows in the CLIN mapping to delete.")
            return
        refs = self.clin_tree._row_refs
        to_remove = [refs[iid] for iid in sel if refs.get(iid) and refs[iid]["from"] != self.SPECIAL_CLIN_PAD_KEY]
        self.clin_data = [r for r in self.clin_data if r not in to_remove]
        self._refresh_mapping_tree("clin")

    # ---- Phase mapping CRUD ----
    def add_phase(self):
        dlg = MappingDialog(self.root, "Add Phase Mapping", from_label="From (Phase)", to_label="To (Mapped Phase)")
        if dlg.result:
            k, v = dlg.result
            k = k.strip().upper()
            v = v.strip()
            if not k or not v:
                messagebox.showerror("Validation", "Both 'From' and 'To' are required.")
                return
            for row in self.phase_data:
                if row["from"].upper() == k:
                    row["to"] = v
                    row["selected"] = True
                    self._refresh_mapping_tree("phase")
                    return
            self.phase_data.append({"selected": True, "from": k, "to": v})
            self._refresh_mapping_tree("phase")

    def delete_phase(self):
        sel = self.phase_tree.selection()
        if not sel:
            messagebox.showinfo("Delete", "Select one or more rows in the Phase mapping to delete.")
            return
        refs = self.phase_tree._row_refs
        self.phase_data = [r for r in self.phase_data if r not in [refs[iid] for iid in sel]]
        self._refresh_mapping_tree("phase")

    # ---- WBS Code mapping CRUD ----
    def add_wbs_code(self):
        dlg = MappingDialog(self.root, "Add WBS Code Mapping", from_label="From (WBS Code)", to_label="To (Mapped Code)")
        if dlg.result:
            k, v = dlg.result
            k = k.strip()
            v = v.strip()
            if not k or not v:
                messagebox.showerror("Validation", "Both 'From' and 'To' are required.")
                return
            for row in self.wbs_code_data:
                if row["from"] == k:
                    row["to"] = v
                    row["selected"] = True
                    self._refresh_mapping_tree("wbs_code")
                    return
            self.wbs_code_data.append({"selected": True, "from": k, "to": v})
            self._refresh_mapping_tree("wbs_code")

    def delete_wbs_code(self):
        sel = self.wbs_code_tree.selection()
        if not sel:
            messagebox.showinfo("Delete", "Select one or more rows in the WBS code mapping to delete.")
            return
        refs = self.wbs_code_tree._row_refs
        self.wbs_code_data = [r for r in self.wbs_code_data if r not in [refs[iid] for iid in sel]]
        self._refresh_mapping_tree("wbs_code")

    # ---- WBS Description mapping CRUD ----
    def add_wbs_desc(self):
        dlg = MappingDialog(self.root, "Add WBS Description Mapping", from_label="From (Description)", to_label="To (Mapped Description)")
        if dlg.result:
            k, v = dlg.result
            k = k.strip().upper()
            v = v.strip()
            if not k or not v:
                messagebox.showerror("Validation", "Both 'From' and 'To' are required.")
                return
            for row in self.wbs_desc_data:
                if row["from"].upper() == k:
                    row["to"] = v
                    row["selected"] = True
                    self._refresh_mapping_tree("wbs_desc")
                    return
            self.wbs_desc_data.append({"selected": True, "from": k, "to": v})
            self._refresh_mapping_tree("wbs_desc")

    def delete_wbs_desc(self):
        sel = self.wbs_desc_tree.selection()
        if not sel:
            messagebox.showinfo("Delete", "Select one or more rows in the WBS description mapping to delete.")
            return
        refs = self.wbs_desc_tree._row_refs
        self.wbs_desc_data = [r for r in self.wbs_desc_data if r not in [refs[iid] for iid in sel]]
        self._refresh_mapping_tree("wbs_desc")

    # ---- Selection helpers ----
    def select_all(self, which):
        tree = {
            "clin": self.clin_tree,
            "phase": self.phase_tree,
            "wbs_code": self.wbs_code_tree,
            "wbs_desc": self.wbs_desc_tree,
        }[which]
        for iid in tree.get_children():
            row = tree._row_refs.get(iid)
            if row:
                if which == "clin" and row["from"] == self.SPECIAL_CLIN_PAD_KEY:
                    continue
                row["selected"] = True
                vals = list(tree.item(iid, "values"))
                vals[0] = "☑"
                tree.item(iid, values=tuple(vals))

    def clear_all(self, which):
        tree = {
            "clin": self.clin_tree,
            "phase": self.phase_tree,
            "wbs_code": self.wbs_code_tree,
            "wbs_desc": self.wbs_desc_tree,
        }[which]
        for iid in tree.get_children():
            row = tree._row_refs.get(iid)
            if row:
                row["selected"] = False
                vals = list(tree.item(iid, "values"))
                vals[0] = "☐"
                tree.item(iid, values=tuple(vals))

    # ---- Import/Export for mappings (supports all 4 types) ----
    def import_mapping(self, which):
        path = filedialog.askopenfilename(
            title=f"Import {which.replace('_', ' ').title()} Mapping (CSV or JSON)",
            filetypes=[("CSV Files", "*.csv"), ("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        if not path:
            return
        p = pathlib.Path(path)
        try:
            items = []
            if p.suffix.lower() == ".json":
                with open(p, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if not isinstance(data, dict):
                    raise ValueError("JSON must be an object mapping 'from' -> 'to'.")
                for k, v in data.items():
                    items.append((str(k), str(v)))
            else:
                with open(p, "r", encoding="utf-8-sig", newline="") as f:
                    reader = csv.reader(f)
                    rows = list(reader)
                start_idx = 1 if rows and rows[0] and rows[0][0].lower().startswith("from") else 0
                for row in rows[start_idx:]:
                    if len(row) < 2:
                        continue
                    items.append((row[0].strip(), row[1].strip()))

            if which == "clin":
                target = self.clin_data
                special_row = None
                for r in target:
                    if r["from"] == self.SPECIAL_CLIN_PAD_KEY:
                        special_row = r
                        break
                new_target = [special_row] if special_row else []
                for k, v in items:
                    if k == self.SPECIAL_CLIN_PAD_KEY:
                        if isinstance(v, str) and v.strip().lower() in ("on", "true", "yes", "1"):
                            if special_row:
                                special_row["selected"] = True
                        elif isinstance(v, str) and v.strip().lower() in ("off", "false", "no", "0"):
                            if special_row:
                                special_row["selected"] = False
                        continue
                    found = False
                    for r in new_target:
                        if r and r["from"] == k:
                            r["to"] = v
                            r["selected"] = True
                            found = True
                            break
                    if not found:
                        new_target.append({"selected": True, "from": k, "to": v})
                self.clin_data = [r for r in new_target if r]

            elif which in ("phase", "wbs_desc"):
                target = self.phase_data if which == "phase" else self.wbs_desc_data
                for k, v in items:
                    k_use = k.upper()
                    found = False
                    for r in target:
                        if r["from"].upper() == k_use:
                            r["to"] = v
                            r["selected"] = True
                            found = True
                            break
                    if not found:
                        target.append({"selected": True, "from": k_use, "to": v})

            elif which == "wbs_code":
                target = self.wbs_code_data
                for k, v in items:
                    found = False
                    for r in target:
                        if r["from"] == k:
                            r["to"] = v
                            r["selected"] = True
                            found = True
                            break
                    if not found:
                        target.append({"selected": True, "from": k, "to": v})

            self._refresh_mapping_tree(which)
            messagebox.showinfo("Import", f"Imported {which.replace('_', ' ').title()} mappings from:\n{path}")

        except Exception as ex:
            messagebox.showerror("Import Error", f"Failed to import:\n{ex}")

    def export_mapping(self, which):
        path = filedialog.asksaveasfilename(
            title=f"Export {which.replace('_', ' ').title()} Mapping",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("JSON Files", "*.json")]
        )
        if not path:
            return
        p = pathlib.Path(path)
        target = {
            "clin": self.clin_data,
            "phase": self.phase_data,
            "wbs_code": self.wbs_code_data,
            "wbs_desc": self.wbs_desc_data
        }[which]
        try:
            if p.suffix.lower() == ".json":
                if which == "clin":
                    obj = {}
                    for r in target:
                        key = r["from"]
                        val = ("on" if (r["from"] == self.SPECIAL_CLIN_PAD_KEY and r["selected"]) else r["to"])
                        obj[key] = val
                    with open(p, "w", encoding="utf-8") as f:
                        json.dump(obj, f, indent=2)
                elif which in ("phase", "wbs_desc"):
                    obj = {r["from"].upper(): r["to"] for r in target}
                    with open(p, "w", encoding="utf-8") as f:
                        json.dump(obj, f, indent=2)
                else:  # wbs_code
                    obj = {r["from"]: r["to"] for r in target}
                    with open(p, "w", encoding="utf-8") as f:
                        json.dump(obj, f, indent=2)
            else:
                with open(p, "w", encoding="utf-8", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(["From", "To"])
                    for r in target:
                        if which in ("phase", "wbs_desc"):
                            key = r["from"].upper()
                        else:
                            key = r["from"]
                        val = ("on" if (which == "clin" and r["from"] == self.SPECIAL_CLIN_PAD_KEY and r["selected"]) else r["to"])
                        writer.writerow([key, val])
            messagebox.showinfo("Export", f"Exported {which.replace('_', ' ').title()} mappings to:\n{path}")
        except Exception as ex:
            messagebox.showerror("Export Error", f"Failed to export:\n{ex}")

    # ---------- Helpers to read flags & dicts ----------
    def _clin_pad_enabled(self) -> bool:
        for r in self.clin_data:
            if r["from"] == self.SPECIAL_CLIN_PAD_KEY and r["selected"]:
                return True
        return False

    def _build_mapping_dict(self, which):
        src = {
            "clin": self.clin_data,
            "phase": self.phase_data,
            "wbs_code": self.wbs_code_data,
            "wbs_desc": self.wbs_desc_data
        }[which]
        d = {}
        for r in src:
            if r["selected"]:
                if which == "clin" and r["from"] == self.SPECIAL_CLIN_PAD_KEY:
                    continue
                key = r["from"].upper() if which in ("phase", "wbs_desc") else r["from"]
                d[key] = r["to"]
        return d

    # ---------- IO helpers ----------
    def log(self, msg):
        try:
            self.status_box.insert("end", msg + "\n")
            self.status_box.see("end")
        except Exception:
            # Fallback to stdout if textbox isn't ready
            print(msg)

    def browse_input_file(self):
        path = filedialog.askopenfilename(
            title="Select BOE Summary Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if path:
            self.input_path.set(path)

    def browse_input_folder(self):
        path = filedialog.askdirectory(
            title="Select BOE Summary folder",
            mustexist=True
        )
        if path:
            self.input_path.set(path)

    def browse_output_dir(self):
        path = filedialog.askdirectory(
            title="Select Output Folder",
            mustexist=True
        )
        if path:
            self.output_dir.set(path)

    def browse_output_file(self):
        path = filedialog.asksaveasfilename(
            title="Select Output File (.xlsx)",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            p = pathlib.Path(path)
            if p.suffix.lower() != ".xlsx":
                p = p.with_suffix(".xlsx")
            self.output_file_path.set(str(p))

    def _sanitize_filename(self, name: str) -> str:
        sanitized = re.sub(r'[<>:"/\\|?*]+', "_", name).strip()
        return sanitized or "Charge_Codes"

    def _final_output_path(self) -> pathlib.Path:
        mode = self.output_mode.get()
        if mode == "folder":
            folder = pathlib.Path(self.output_dir.get().strip())
            fname = self._sanitize_filename(self.output_filename.get().strip() or "Charge_Codes")
            p = pathlib.Path(fname)
            if p.suffix:
                if p.suffix.lower() != ".xlsx":
                    p = p.with_suffix(".xlsx")
            else:
                p = p.with_suffix(".xlsx")
            return folder / p.name
        else:
            fp = pathlib.Path(self.output_file_path.get().strip())
            if fp.suffix.lower() != ".xlsx":
                fp = fp.with_suffix(".xlsx")
            return fp

    def validate(self):
        errs = []
        in_path = self.input_path.get().strip()
        proj = self.project_number.get().strip()
        sel = self.contract_type.get().strip()
        mode = self.output_mode.get()

        if not in_path:
            errs.append("Input path (file or folder) is required.")
        else:
            p = pathlib.Path(in_path)
            if not p.exists():
                errs.append(f"Input path not found: {in_path}")
            elif p.is_file() and p.suffix.lower() not in [".xlsx", ".xlsm"]:
                errs.append("Input file must be .xlsx or .xlsm.")
            elif p.is_dir():
                candidates = list(p.glob("*.xlsx")) + list(p.glob("*.xlsm"))
                if not candidates:
                    errs.append("No Excel files (*.xlsx, *.xlsm) found in the selected folder.")

        if not proj:
            errs.append("Project Number (Level 1) is required.")
        if not sel:
            errs.append("Contract Type (Level 2 default) is required.")

        if mode == "folder":
            out_dir = self.output_dir.get().strip()
            out_name = self.output_filename.get().strip() or "Charge_Codes"
            if not out_dir:
                errs.append("Output folder is required (or switch to 'Single File Path' mode).")
            elif not pathlib.Path(out_dir).exists():
                errs.append(f"Output folder not found: {out_dir}")
            sanitized = self._sanitize_filename(out_name)
            if sanitized != out_name:
                errs.append(f"Output file name contained invalid characters; suggested: '{sanitized}'")
        else:
            out_file = self.output_file_path.get().strip()
            if not out_file:
                errs.append("Output file path is required in 'Single File Path' mode.")
            else:
                fp = pathlib.Path(out_file)
                if not fp.parent.exists():
                    errs.append(f"Output file folder does not exist: {fp.parent}")
                if fp.suffix and fp.suffix.lower() not in [".xlsx", ".xlsm"]:
                    errs.append("Output file must end with .xlsx (we will enforce .xlsx).")

        try:
            _ = self._final_output_path()
        except Exception as ex:
            errs.append(f"Could not construct output path: {ex}")

        return errs

    def run_clicked(self):
        errs = self.validate()
        try:
            self.status_box.delete("1.0", "end")
        except Exception:
            pass
        if errs:
            for e in errs:
                self.log(f"❌ {e}")
            messagebox.showerror("Validation Error", "\n".join(errs))
            return

        self.btn_run.configure(state="disabled")
        self.log("▶ Starting…")
        t = threading.Thread(target=self._do_work, daemon=True)
        t.start()

    def _do_work(self):
        try:
            in_path = self.input_path.get().strip()
            proj = self.project_number.get().strip()
            selected = self.contract_type.get().strip()
            ctype = self._parse_contract_code(selected) or "1"
            final_out_path = self._final_output_path()

            # Build mapping dicts from selected rows only
            clin_map = self._build_mapping_dict("clin")
            phase_map = self._build_mapping_dict("phase")
            wbs_code_map = self._build_mapping_dict("wbs_code")
            wbs_desc_map = self._build_mapping_dict("wbs_desc")

            enable_clin = bool(self.use_clin_map.get())
            enable_phase = bool(self.use_phase_map.get())
            enable_wbs_code = bool(self.use_wbs_code_map.get())
            enable_wbs_desc = bool(self.use_wbs_desc_map.get())

            # Rule flag from CLIN mapping UI
            pad_clin_to_6 = self._clin_pad_enabled()

            # Build {LABEL_UPPER: CODE} mapping for Cost Type -> Level 2
            contract_label_to_code = self._contract_label_to_code_map()

            self.log(f"Loading input from: {in_path}")
            df_boe = read_boe_summary_multi(in_path)

            self.log("Building paths…")
            paths = build_paths(
                df_boe,
                project_number=proj,
                contract_type=ctype,  # default numeric code when Cost Type missing/unrecognized
                clin_map=clin_map,
                phase_map=phase_map,
                wbs_code_map=wbs_code_map,
                wbs_desc_map=wbs_desc_map,
                enable_clin_map=enable_clin,
                enable_phase_map=enable_phase,
                enable_wbs_code_map=enable_wbs_code,
                enable_wbs_desc_map=enable_wbs_desc,
                pad_clin_to_6=pad_clin_to_6,
                contract_label_to_code=contract_label_to_code,  # NEW
            )

            self.log("Expanding hierarchy…")
            df_out = expand_hierarchy(paths)

            final_out_path.parent.mkdir(parents=True, exist_ok=True)
            self.log(f"Writing Excel to: {final_out_path}")
            final_out = write_output_excel(df_out, final_out_path)

            self.log(f"✅ Done. Rows written (hierarchical): {len(df_out)}")
            self.log(f"📄 Output: {final_out}")
            messagebox.showinfo("Success", f"Wrote Excel:\n{final_out}\n\nRows: {len(df_out)}")

        except Exception as ex:
            err_text = "".join(traceback.format_exception(type(ex), ex, ex.__traceback__))
            self.log("❌ Error occurred:\n" + err_text)
            messagebox.showerror("Error", f"An error occurred:\n\n{ex}\n\nSee status for details.")
        finally:
            try:
                self.btn_run.configure(state="normal")
            except Exception:
                pass

    def load_config(self):
        if not self.config_path.exists():
            return
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            # Load mappings
            if 'clin_data' in config:
                self.clin_data = config['clin_data']
            if 'phase_data' in config:
                self.phase_data = config['phase_data']
            if 'wbs_code_data' in config:
                self.wbs_code_data = config['wbs_code_data']
            if 'wbs_desc_data' in config:
                self.wbs_desc_data = config['wbs_desc_data']
            # Load toggles
            if 'use_clin_map' in config:
                self.use_clin_map.set(config['use_clin_map'])
            if 'use_phase_map' in config:
                self.use_phase_map.set(config['use_phase_map'])
            if 'use_wbs_code_map' in config:
                self.use_wbs_code_map.set(config['use_wbs_code_map'])
            if 'use_wbs_desc_map' in config:
                self.use_wbs_desc_map.set(config['use_wbs_desc_map'])
            # Load contract options
            if 'contract_options' in config:
                self.contract_options = config['contract_options']
                self.contract_combo.configure(values=self.contract_options)
                if self.contract_options:
                    self.contract_combo.set(self.contract_options[0])
            # Refresh trees
            self._refresh_mapping_tree("clin")
            self._refresh_mapping_tree("phase")
            self._refresh_mapping_tree("wbs_code")
            self._refresh_mapping_tree("wbs_desc")
        except Exception as ex:
            messagebox.showerror("Config Load Error", f"Failed to load preferences:\n{ex}")

    def save_preferences(self):
        try:
            self.config_dir.mkdir(parents=True, exist_ok=True)
            config = {
                'clin_data': self.clin_data,
                'phase_data': self.phase_data,
                'wbs_code_data': self.wbs_code_data,
                'wbs_desc_data': self.wbs_desc_data,
                'use_clin_map': self.use_clin_map.get(),
                'use_phase_map': self.use_phase_map.get(),
                'use_wbs_code_map': self.use_wbs_code_map.get(),
                'use_wbs_desc_map': self.use_wbs_desc_map.get(),
                'contract_options': self.contract_options,
            }
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2)
            messagebox.showinfo("Saved", "Preferences saved successfully.")
        except Exception as ex:
            messagebox.showerror("Save Error", f"Failed to save preferences:\n{ex}")

    def reset_to_defaults(self):
        if not messagebox.askyesno("Reset", "Reset all mappings and settings to defaults? This cannot be undone."):
            return
        self._seed_clin_defaults()
        self._seed_phase_defaults()
        self._seed_wbs_code_defaults()
        self._seed_wbs_desc_defaults()
        self.use_clin_map.set(True)
        self.use_phase_map.set(True)
        self.use_wbs_code_map.set(True)
        self.use_wbs_desc_map.set(True)
        self.contract_options = ["1: FFP", "2: FPI", "3: CPFF", "4: CPIF", "5: FPAF", "7: T&M", "8: CPAF"]
        self.contract_combo.configure(values=self.contract_options)
        if self.contract_options:
            self.contract_combo.set(self.contract_options[0])
        self._refresh_mapping_tree("clin")
        self._refresh_mapping_tree("phase")
        self._refresh_mapping_tree("wbs_code")
        self._refresh_mapping_tree("wbs_desc")
        messagebox.showinfo("Reset", "Reset to defaults completed.")

# --- Root creation with DnD + CTk theme ---

def main_gui():
    # --- Create root (CTk + optional DnD) ---
    if DND_AVAILABLE:
        class CTkDnD(ctk.CTk, TkinterDnD.Tk):
            def __init__(self, *args, **kwargs):
                ctk.CTk.__init__(self, *args, **kwargs)
                TkinterDnD.Tk.__init__(self, *args, **kwargs)
        root = CTkDnD()
    else:
        root = ctk.CTk()

    root.geometry("1180x900")
    root.minsize(900, 650)

    # --- Close handler: only run when user clicks X ---
    def on_close():
        try:
            root.quit()     # exit mainloop
        finally:
            root.destroy()  # destroy windows

    root.protocol("WM_DELETE_WINDOW", on_close)

    app = ChargeCodesGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main_gui()
