#!/usr/bin/env python3
"""
Charge Code Generator GUI (Tkinter)

Features:
- Input: Browse file or folder; optional drag & drop (via tkinterdnd2 if installed).
- Level 1: Project Number (text).
- Level 2: Contract Type (dropdown; e.g., "1: FFP", "2: CRNF") + editor dialog.
- CLIN & Phase mappings: tables with search bars, per-row checkboxes, Select All / Clear All, Import/Export.
  * Columns centered for readability.
- Output: Choose either:
    (A) Output Folder + File Name (default "Charge_Codes") with Drag & Drop, OR
    (B) Single Output File Path (.xlsx) with Drag & Drop
  We force .xlsx as the final extension.
- Threaded run; status log; full error stacktraces.

Dependencies:
- Required: pandas, openpyxl
- Optional (for drag & drop): tkinterdnd2
"""

import re
import pathlib
import threading
import traceback
import json
import csv
import pandas as pd
from tkinter import Tk, StringVar, BooleanVar, END, Toplevel
from tkinter import filedialog, messagebox
from tkinter import ttk

# Try to enable drag & drop (tkinterdnd2)
DND_AVAILABLE = False
TkRootClass = Tk
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
    TkRootClass = TkinterDnD.Tk
except Exception:
    DND_AVAILABLE = False
    TkRootClass = Tk

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

def _extract_clin(clin_text, clin_map=None, enable_map=True):
    """From text like 'CLIN 0001' -> '0001'; applies optional CLIN mapping (e.g., to lot codes)."""
    if clin_text is None:
        return None
    s = str(clin_text).strip()
    m = re.search(r"\bCLIN\s*(\w+)\b", s, flags=re.IGNORECASE)
    code = m.group(1).strip() if m else (
        re.search(r"\b(\w+)\b", s).group(1).strip() if re.search(r"\b(\w+)\b", s) else s
    )
    code = code.strip()
    if enable_map and clin_map:
        return clin_map.get(code, code)
    return code

def _normalize_phase(phase_text, phase_map=None, enable_map=True):
    """Map phase text via user-defined phase_map; case-insensitive on keys."""
    if phase_text is None:
        return None
    p = str(phase_text).strip()
    if enable_map and phase_map:
        mapped = phase_map.get(p.upper())
        if mapped is not None:
            return mapped
    return p

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
            df["__SourceFile__"] = f.name  # optional provenance
            dfs.append(df)
        out = pd.concat(dfs, ignore_index=True)
        out = out.astype(str)
        out.columns = [c.strip() if isinstance(c, str) else c for c in out.columns]
        return out
    else:
        raise ValueError(f"Input path is neither a file nor folder: {p}")

def build_paths(df, project_number, contract_type, clin_map=None, phase_map=None,
                enable_clin_map=True, enable_phase_map=True):
    """Build full paths (Level1..Level5) from the BOE Summary; sort & de-duplicate."""
    clin_col  = _find_column(df, ["clin"])
    wbs_col   = _find_column(df, ["wbs"])
    phase_col = _find_column(df, ["phase"])
    missing = [name for name, col in [("CLIN", clin_col), ("WBS", wbs_col), ("Phase", phase_col)] if col is None]
    if missing:
        raise ValueError(f"Missing columns: {', '.join(missing)}. Found: {list(df.columns)}")

    rows = []
    for _, r in df.iterrows():
        level3 = _extract_clin(r.get(clin_col), clin_map=clin_map, enable_map=enable_clin_map)
        level4 = _extract_wbs_code(r.get(wbs_col))
        level5 = _normalize_phase(r.get(phase_col), phase_map=phase_map, enable_map=enable_phase_map)
        if not (level3 and level4 and level5):
            continue
        level1 = str(project_number).strip()
        level2 = str(contract_type).strip() if contract_type else "1"
        rows.append({"Level 1": level1, "Level 2": level2, "Level 3": level3, "Level 4": level4, "Level 5": level5})

    if not rows:
        raise ValueError("No valid rows found (CLIN/WBS/Phase missing).")

    paths = pd.DataFrame(rows).drop_duplicates()
    paths = paths.sort_values(by=["Level 1", "Level 2", "Level 3", "Level 4", "Level 5"]).reset_index(drop=True)
    return paths

def expand_hierarchy(paths):
    """Expand sorted full paths into hierarchical rows with partial parent lines and full lines."""
    out = []
    last = {"Level 1": None, "Level 2": None, "Level 3": None, "Level 4": None, "Level 5": None}

    def _row(levels_present, paths_row):
        vals = {lvl: paths_row[lvl] if lvl in levels_present else "" for lvl in ["Level 1", "Level 2", "Level 3", "Level 4", "Level 5"]}
        parts = [paths_row[lvl] for lvl in levels_present]
        vals["Charge String"] = ".".join(parts) + "."
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
# Tkinter GUI Implementation
# --------------------------

class MappingDialog(Toplevel):
    """Simple modal dialog to add/edit a mapping pair."""
    def __init__(self, parent, title, from_label="From", to_label="To", initial_from="", initial_to=""):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.result = None

        pad = {"padx": 8, "pady": 6}
        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, **pad)

        ttk.Label(frm, text=from_label + ":").grid(row=0, column=0, sticky="w", **pad)
        self.var_from = StringVar(value=initial_from)
        ttk.Entry(frm, textvariable=self.var_from, width=30).grid(row=0, column=1, sticky="w", **pad)

        ttk.Label(frm, text=to_label + ":").grid(row=1, column=0, sticky="w", **pad)
        self.var_to = StringVar(value=initial_to)
        ttk.Entry(frm, textvariable=self.var_to, width=30).grid(row=1, column=1, sticky="w", **pad)

        btns = ttk.Frame(frm)
        btns.grid(row=2, column=0, columnspan=2, sticky="e", **pad)
        ttk.Button(btns, text="Cancel", command=self._cancel).pack(side="right", padx=5)
        ttk.Button(btns, text="OK", command=self._ok).pack(side="right", padx=5)

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
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

class ContractTypesEditor(Toplevel):
    """Modal editor for Contract Type list: items like '1: FFP', '2: CRNF'."""
    def __init__(self, parent, contract_items):
        super().__init__(parent)
        self.title("Edit Contract Types (Level 2)")
        self.resizable(True, True)
        self.result = None

        pad = {"padx": 8, "pady": 6}
        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, **pad)

        ttk.Label(frm, text="Contract Types (Code / Label):").grid(row=0, column=0, sticky="w", **pad)
        self.tree = ttk.Treeview(frm, columns=("Code", "Label"), show="headings", height=10)
        self.tree.heading("Code", text="Code")
        self.tree.heading("Label", text="Label")
        self.tree.column("Code", width=100, anchor="center")
        self.tree.column("Label", width=220, anchor="center")
        self.tree.grid(row=1, column=0, sticky="nsew", **pad)

        yscroll = ttk.Scrollbar(frm, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=yscroll.set)

        for item in contract_items:
            code, label = self._split_item(item)
            self.tree.insert("", END, values=(code, label))

        btns = ttk.Frame(frm)
        btns.grid(row=2, column=0, sticky="w", **pad)
        ttk.Button(btns, text="Add", command=self.add_item).pack(side="left", padx=4)
        ttk.Button(btns, text="Edit", command=self.edit_item).pack(side="left", padx=4)
        ttk.Button(btns, text="Delete", command=self.delete_item).pack(side="left", padx=4)
        ttk.Button(btns, text="Import…", command=self.import_items).pack(side="left", padx=4)
        ttk.Button(btns, text="Export…", command=self.export_items).pack(side="left", padx=4)

        action = ttk.Frame(frm)
        action.grid(row=3, column=0, sticky="e", **pad)
        ttk.Button(action, text="Cancel", command=self._cancel).pack(side="right", padx=5)
        ttk.Button(action, text="OK", command=self._ok).pack(side="right", padx=5)

        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(1, weight=1)

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
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
            self.tree.insert("", END, values=(code, label))

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
                    self.tree.insert("", END, values=(code, label))
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

class ChargeCodesGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Charge Code Generator (BOE Summary → Hierarchy)")

        # Vars
        self.input_path = StringVar()
        self.project_number = StringVar()
        self.contract_type = StringVar()  # holds selected string, e.g., "1: FFP"

        # Output mode: 'folder' or 'file'
        self.output_mode = StringVar(value="folder")
        self.output_dir = StringVar()
        self.output_filename = StringVar(value="Charge_Codes")
        self.output_file_path = StringVar()

        # Level 2 defaults (edit here)
        self.contract_options = ["1: FFP", "2: CRNF"]

        # Mapping toggles
        self.use_clin_map = BooleanVar(value=True)
        self.use_phase_map = BooleanVar(value=True)

        # Mapping datasets with selection flags
        self.clin_data = []   # list of dicts: {'selected':bool,'from':str,'to':str}
        self.phase_data = []  # same structure

        # Search vars
        self.clin_search = StringVar()
        self.phase_search = StringVar()

        pad = {"padx": 8, "pady": 6}

        frm = ttk.Frame(root)
        frm.pack(fill="both", expand=True, **pad)

        # Row 0: Input (file or folder) with browse + optional drag & drop
        ttk.Label(frm, text="Input (Excel file or folder):").grid(row=0, column=0, sticky="w", **pad)
        self.input_entry = ttk.Entry(frm, textvariable=self.input_path, width=60)
        self.input_entry.grid(row=0, column=1, sticky="ew", **pad)
        if DND_AVAILABLE:
            try:
                self.input_entry.drop_target_register(DND_FILES)
                self.input_entry.dnd_bind("<<Drop>>", self._on_drop_input)
            except Exception:
                pass
        ttk.Button(frm, text="Browse File…", command=self.browse_input_file).grid(row=0, column=2, **pad)
        ttk.Button(frm, text="Browse Folder…", command=self.browse_input_folder).grid(row=0, column=3, **pad)
        if not DND_AVAILABLE:
            ttk.Label(frm, text="(Drag & drop available if 'tkinterdnd2' is installed)").grid(row=1, column=1, sticky="w", **pad)

        # Row 2: Project Number
        ttk.Label(frm, text="Project Number (Level 1):").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.project_number, width=30).grid(row=2, column=1, sticky="w", **pad)

        # Row 3: Contract Type (Level 2) dropdown + editor
        ttk.Label(frm, text="Contract Type (Level 2):").grid(row=3, column=0, sticky="w", **pad)
        self.contract_combo = ttk.Combobox(
            frm, textvariable=self.contract_type, values=self.contract_options, state="readonly", width=30
        )
        self.contract_combo.grid(row=3, column=1, sticky="w", **pad)
        if self.contract_options:
            self.contract_combo.current(0)
        ttk.Button(frm, text="Edit…", command=self.edit_contract_types).grid(row=3, column=2, **pad)

        # Row 4: Output mode controls
        ttk.Label(frm, text="Output Mode:").grid(row=4, column=0, sticky="w", **pad)
        rb_folder = ttk.Radiobutton(frm, text="Folder + File Name", variable=self.output_mode, value="folder", command=self._on_output_mode_change)
        rb_file   = ttk.Radiobutton(frm, text="Single File Path",    variable=self.output_mode, value="file",   command=self._on_output_mode_change)
        rb_folder.grid(row=4, column=1, sticky="w", **pad)
        rb_file.grid(row=4, column=2, sticky="w", **pad)

        # Row 5-6: Folder + File Name (default mode)
        self.row_output_folder_label = ttk.Label(frm, text="Output Folder:")
        self.row_output_folder_label.grid(row=5, column=0, sticky="w", **pad)
        self.row_output_folder_entry = ttk.Entry(frm, textvariable=self.output_dir, width=60)
        self.row_output_folder_entry.grid(row=5, column=1, sticky="ew", **pad)
        self.row_output_folder_browse = ttk.Button(frm, text="Browse…", command=self.browse_output_dir)
        self.row_output_folder_browse.grid(row=5, column=2, **pad)
        # DnD for Output Folder
        if DND_AVAILABLE:
            try:
                self.row_output_folder_entry.drop_target_register(DND_FILES)
                self.row_output_folder_entry.dnd_bind("<<Drop>>", self._on_drop_output_folder)
            except Exception:
                pass

        self.row_output_name_label = ttk.Label(frm, text="Output File Name:")
        self.row_output_name_label.grid(row=6, column=0, sticky="w", **pad)
        self.row_output_name_entry = ttk.Entry(frm, textvariable=self.output_filename, width=30)
        self.row_output_name_entry.grid(row=6, column=1, sticky="w", **pad)
        ttk.Label(frm, text="(default: Charge_Codes; .xlsx will be added)").grid(row=6, column=2, sticky="w", **pad)

        # Row 7: Output File Path (hidden initially)
        self.row_output_file_label = ttk.Label(frm, text="Output File (.xlsx):")
        self.row_output_file_entry = ttk.Entry(frm, textvariable=self.output_file_path, width=60)
        self.row_output_file_browse = ttk.Button(frm, text="Browse…", command=self.browse_output_file)
        # DnD for Output File Path
        if DND_AVAILABLE:
            try:
                self.row_output_file_entry.drop_target_register(DND_FILES)
                self.row_output_file_entry.dnd_bind("<<Drop>>", self._on_drop_output_file)
            except Exception:
                pass

        # Row 8: Mapping editors container
        editors = ttk.Frame(frm)
        editors.grid(row=8, column=0, columnspan=4, sticky="nsew", **pad)

        # ---- CLIN Mapping Editor ----
        clin_frame = ttk.LabelFrame(editors, text="CLIN (Contract) Mapping")
        clin_frame.grid(row=0, column=0, sticky="nsew", **pad)

        top_clin = ttk.Frame(clin_frame)
        top_clin.grid(row=0, column=0, sticky="ew", **pad)
        ttk.Checkbutton(top_clin, text="Apply CLIN mapping", variable=self.use_clin_map).pack(side="left")
        ttk.Label(top_clin, text="Search:").pack(side="left", padx=(12, 4))
        ent_clin_search = ttk.Entry(top_clin, textvariable=self.clin_search, width=20)
        ent_clin_search.pack(side="left")

        self.clin_tree = self._create_selectable_mapping_tree(clin_frame, start_row=1, columns=("✓", "From", "To"))
        self._seed_clin_defaults()
        self._refresh_mapping_tree("clin")

        clin_btns = ttk.Frame(clin_frame)
        clin_btns.grid(row=2, column=0, sticky="w", **pad)
        ttk.Button(clin_btns, text="Add", command=self.add_clin).pack(side="left", padx=4)
        ttk.Button(clin_btns, text="Delete", command=self.delete_clin).pack(side="left", padx=4)
        ttk.Button(clin_btns, text="Select All", command=lambda: self.select_all("clin")).pack(side="left", padx=4)
        ttk.Button(clin_btns, text="Clear All", command=lambda: self.clear_all("clin")).pack(side="left", padx=4)
        ttk.Button(clin_btns, text="Import…", command=lambda: self.import_mapping("clin")).pack(side="left", padx=4)
        ttk.Button(clin_btns, text="Export…", command=lambda: self.export_mapping("clin")).pack(side="left", padx=4)

        # ---- Phase Mapping Editor ----
        phase_frame = ttk.LabelFrame(editors, text="Phase Mapping")
        phase_frame.grid(row=0, column=1, sticky="nsew", **pad)

        top_phase = ttk.Frame(phase_frame)
        top_phase.grid(row=0, column=0, sticky="ew", **pad)
        ttk.Checkbutton(top_phase, text="Apply Phase mapping", variable=self.use_phase_map).pack(side="left")
        ttk.Label(top_phase, text="Search:").pack(side="left", padx=(12, 4))
        ent_phase_search = ttk.Entry(top_phase, textvariable=self.phase_search, width=20)
        ent_phase_search.pack(side="left")

        self.phase_tree = self._create_selectable_mapping_tree(phase_frame, start_row=1, columns=("✓", "From (Upper)", "To"))
        self._seed_phase_defaults()
        self._refresh_mapping_tree("phase")

        phase_btns = ttk.Frame(phase_frame)
        phase_btns.grid(row=2, column=0, sticky="w", **pad)
        ttk.Button(phase_btns, text="Add", command=self.add_phase).pack(side="left", padx=4)
        ttk.Button(phase_btns, text="Delete", command=self.delete_phase).pack(side="left", padx=4)
        ttk.Button(phase_btns, text="Select All", command=lambda: self.select_all("phase")).pack(side="left", padx=4)
        ttk.Button(phase_btns, text="Clear All", command=lambda: self.clear_all("phase")).pack(side="left", padx=4)
        ttk.Button(phase_btns, text="Import…", command=lambda: self.import_mapping("phase")).pack(side="left", padx=4)
        ttk.Button(phase_btns, text="Export…", command=lambda: self.export_mapping("phase")).pack(side="left", padx=4)

        # Row 9: Run button
        self.btn_run = ttk.Button(frm, text="Run", command=self.run_clicked)
        self.btn_run.grid(row=9, column=1, sticky="w", **pad)

        # Row 10: Status
        ttk.Label(frm, text="Status:").grid(row=10, column=0, sticky="nw", **pad)
        from tkinter import Text
        self.status_box = Text(frm, height=10, width=90, wrap="word")
        self.status_box.grid(row=10, column=1, columnspan=3, sticky="nsew", **pad)

        # Expandable layout
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(10, weight=1)
        editors.columnconfigure(0, weight=1)
        editors.columnconfigure(1, weight=1)

        # Search live filters
        self.clin_search.trace_add("write", lambda *_: self._refresh_mapping_tree("clin"))
        self.phase_search.trace_add("write", lambda *_: self._refresh_mapping_tree("phase"))

        # Toggle checkbox by click
        self.clin_tree.bind("<Button-1>", lambda e: self._on_tree_click(e, "clin"))
        self.phase_tree.bind("<Button-1>", lambda e: self._on_tree_click(e, "phase"))

        # Init output mode UI
        self._on_output_mode_change()

    # ---------- Drag & Drop handlers ----------
    def _normalize_dropped_paths(self, data: str):
        """
        Convert tkdnd drop data to a list of paths (handles braces & spaces).
        """
        data = data.strip()
        paths = []
        buf = ""
        brace = False
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
        """Handle drop data; take the first path."""
        paths = self._normalize_dropped_paths(event.data)
        if paths:
            self.input_path.set(paths[0])

    def _on_drop_output_folder(self, event):
        """
        Handle DnD onto Output Folder entry.
        Accepts a folder; if a file is dropped, uses its parent folder.
        """
        paths = self._normalize_dropped_paths(event.data)
        if not paths:
            return
        p = pathlib.Path(paths[0])
        if p.is_file():
            p = p.parent
        if p.exists():
            self.output_dir.set(str(p))

    def _on_drop_output_file(self, event):
        """
        Handle DnD onto Output File entry (single file mode).
        Accepts a file or a folder; enforces .xlsx extension.
        """
        paths = self._normalize_dropped_paths(event.data)
        if not paths:
            return
        p = pathlib.Path(paths[0])

        if p.is_dir():
            # If a folder is dropped, construct a file path using current filename (or default)
            fname = self._sanitize_filename(self.output_filename.get().strip() or "Charge_Codes")
            fp = (p / fname).with_suffix(".xlsx")
            self.output_file_path.set(str(fp))
            return

        # If a file is dropped, enforce .xlsx
        if p.suffix.lower() != ".xlsx":
            p = p.with_suffix(".xlsx")
        # Ensure parent exists
        if p.parent.exists():
            self.output_file_path.set(str(p))
        else:
            messagebox.showerror("Invalid Path", f"Folder does not exist:\n{p.parent}")

    # ---------- Contract types ----------
    def edit_contract_types(self):
        editor = ContractTypesEditor(self.root, self.contract_options)
        if editor.result is not None:
            self.contract_options = editor.result
            self.contract_combo["values"] = self.contract_options
            current = self.contract_type.get()
            if current in self.contract_options:
                self.contract_combo.set(current)
            elif self.contract_options:
                self.contract_combo.current(0)

    def _parse_contract_code(self, selected_text):
        """Parse 'code: label' -> 'code'. If no colon, return as-is."""
        s = str(selected_text).strip()
        if ":" in s:
            code, _ = s.split(":", 1)
            return code.strip()
        return s

    # ---------- Output mode UI ----------
    def _on_output_mode_change(self):
        mode = self.output_mode.get()
        if mode == "folder":
            # Show folder+filename
            self.row_output_folder_label.grid()
            self.row_output_folder_entry.grid()
            self.row_output_folder_browse.grid()
            self.row_output_name_label.grid()
            self.row_output_name_entry.grid()
            # Hide file path controls
            self.row_output_file_label.grid_remove()
            self.row_output_file_entry.grid_remove()
            self.row_output_file_browse.grid_remove()
        else:
            # Hide folder+filename
            self.row_output_folder_label.grid_remove()
            self.row_output_folder_entry.grid_remove()
            self.row_output_folder_browse.grid_remove()
            self.row_output_name_label.grid_remove()
            self.row_output_name_entry.grid_remove()
            # Show file path controls
            self.row_output_file_label.grid(row=7, column=0, sticky="w", padx=8, pady=6)
            self.row_output_file_entry.grid(row=7, column=1, sticky="ew", padx=8, pady=6)
            self.row_output_file_browse.grid(row=7, column=2, padx=8, pady=6)
            # (Re)register DnD if needed (already registered in __init__, safe to repeat)
            if DND_AVAILABLE:
                try:
                    self.row_output_file_entry.drop_target_register(DND_FILES)
                    self.row_output_file_entry.dnd_bind("<<Drop>>", self._on_drop_output_file)
                except Exception:
                    pass

    # ---------- Tree creation with centered columns ----------
    def _create_selectable_mapping_tree(self, parent, start_row=0, columns=("✓", "From", "To")):
        tree = ttk.Treeview(parent, columns=columns, show="headings", height=10)
        for idx, col in enumerate(columns):
            tree.heading(col, text=col)
            width = 80 if idx == 0 else 200
            tree.column(col, width=width, anchor="center")  # center content
        tree.grid(row=start_row, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        yscroll.grid(row=start_row, column=1, sticky="ns")
        tree.configure(yscrollcommand=yscroll.set)
        parent.rowconfigure(start_row, weight=1)
        return tree

    # ---------- Mapping data management ----------
    def _seed_clin_defaults(self):
        # Edit CLIN defaults here (or leave empty to let users import/edit)
        self.clin_data = [
            # {"selected": True,  "from": "0001", "to": "5031AA"},
            # {"selected": True,  "from": "0002", "to": "5031AB"},
        ]

    def _seed_phase_defaults(self):
        # Edit Phase defaults here
        self.phase_data = [
            {"selected": True, "from": "BID",  "to": "1BID"},
            {"selected": True, "from": "LOE",  "to": "ZLOE"},
            {"selected": True, "from": "RAV",  "to": "ZRAV"},
            {"selected": True, "from": "PDV",  "to": "ZPDR"},
            {"selected": True, "from": "DDV",  "to": "ZCDR"},
            {"selected": True, "from": "COD",  "to": "ZCOD"},
            {"selected": True, "from": "HSI",  "to": "ZHSI"},
            {"selected": True, "from": "CVT",  "to": "ZCVT"},
            {"selected": True, "from": "GAT",  "to": "ZGAT"},
            {"selected": True, "from": "PROD", "to": "ZPROD"},
        ]

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
        else:
            tree = self.phase_tree
            data = self._filter_rows(self.phase_data, self.phase_search.get(), upper_keys=True)

        for iid in tree.get_children():
            tree.delete(iid)

        tree._row_refs = {}
        for row in data:
            mark = "☑" if row["selected"] else "☐"
            from_display = row["from"].upper() if which == "phase" else row["from"]
            iid = tree.insert("", END, values=(mark, from_display, row["to"]))
            tree._row_refs[iid] = row

    def _on_tree_click(self, event, which):
        tree = self.clin_tree if which == "clin" else self.phase_tree
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = tree.identify_column(event.x)
        if col != "#1":  # only checkbox column
            return
        iid = tree.identify_row(event.y)
        if not iid:
            return
        row = tree._row_refs.get(iid)
        if not row:
            return
        row["selected"] = not row["selected"]
        mark = "☑" if row["selected"] else "☐"
        vals = list(tree.item(iid, "values"))
        vals[0] = mark
        tree.item(iid, values=tuple(vals))

    def add_clin(self):
        dlg = MappingDialog(self.root, "Add CLIN Mapping", from_label="From (CLIN code)", to_label="To (Lot code)")
        if dlg.result:
            k, v = dlg.result
            k = k.strip()
            v = v.strip()
            if not k or not v:
                messagebox.showerror("Validation", "Both 'From' and 'To' are required.")
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
        self.clin_data = [r for r in self.clin_data if r not in [refs[iid] for iid in sel]]
        self._refresh_mapping_tree("clin")

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

    def select_all(self, which):
        tree = self.clin_tree if which == "clin" else self.phase_tree
        for iid in tree.get_children():
            row = tree._row_refs.get(iid)
            if row:
                row["selected"] = True
                vals = list(tree.item(iid, "values"))
                vals[0] = "☑"
                tree.item(iid, values=tuple(vals))

    def clear_all(self, which):
        tree = self.clin_tree if which == "clin" else self.phase_tree
        for iid in tree.get_children():
            row = tree._row_refs.get(iid)
            if row:
                row["selected"] = False
                vals = list(tree.item(iid, "values"))
                vals[0] = "☐"
                tree.item(iid, values=tuple(vals))

    def import_mapping(self, which):
        path = filedialog.askopenfilename(
            title=f"Import {which.capitalize()} Mapping (CSV or JSON)",
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
            target = self.clin_data if which == "clin" else self.phase_data
            for k, v in items:
                k_use = k if which == "clin" else k.upper()
                found = False
                for r in target:
                    if (r["from"] == k_use) or (which == "phase" and r["from"].upper() == k_use):
                        r["to"] = v
                        r["selected"] = True
                        found = True
                        break
                if not found:
                    target.append({"selected": True, "from": k_use, "to": v})
            self._refresh_mapping_tree(which)
            messagebox.showinfo("Import", f"Imported {which} mappings from:\n{path}")
        except Exception as ex:
            messagebox.showerror("Import Error", f"Failed to import:\n{ex}")

    def export_mapping(self, which):
        path = filedialog.asksaveasfilename(
            title=f"Export {which.capitalize()} Mapping",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("JSON Files", "*.json")]
        )
        if not path:
            return
        p = pathlib.Path(path)
        target = self.clin_data if which == "clin" else self.phase_data
        try:
            if p.suffix.lower() == ".json":
                obj = { (r["from"].upper() if which == "phase" else r["from"]) : r["to"] for r in target }
                with open(p, "w", encoding="utf-8") as f:
                    json.dump(obj, f, indent=2)
            else:
                with open(p, "w", encoding="utf-8", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow(["From", "To"])
                    for r in target:
                        writer.writerow([(r["from"].upper() if which == "phase" else r["from"]), r["to"]])
            messagebox.showinfo("Export", f"Exported {which} mappings to:\n{path}")
        except Exception as ex:
            messagebox.showerror("Export Error", f"Failed to export:\n{ex}")

    # ---------- IO helpers ----------
    def log(self, msg):
        self.status_box.insert(END, msg + "\n")
        self.status_box.see(END)

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
            # ensure .xlsx
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
            # Enforce .xlsx
            if fp.suffix.lower() != ".xlsx":
                fp = fp.with_suffix(".xlsx")
            return fp

    def _build_mapping_dict(self, which):
        """Build mapping dict using ONLY rows with selected=True (phase keys uppercased)."""
        src = self.clin_data if which == "clin" else self.phase_data
        d = {}
        for r in src:
            if r["selected"]:
                key = r["from"].upper() if which == "phase" else r["from"]
                d[key] = r["to"]
        return d

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
            errs.append("Contract Type (Level 2) is required.")

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
                # The parent directory must exist
                if not fp.parent.exists():
                    errs.append(f"Output file folder does not exist: {fp.parent}")
                # will enforce .xlsx later; warn if wrong
                if fp.suffix and fp.suffix.lower() not in [".xlsx", ".xlsm"]:
                    errs.append("Output file must end with .xlsx (we will enforce .xlsx).")

        try:
            _ = self._final_output_path()
        except Exception as ex:
            errs.append(f"Could not construct output path: {ex}")

        return errs

    def run_clicked(self):
        errs = self.validate()
        self.status_box.delete(1.0, END)
        if errs:
            for e in errs:
                self.log(f"❌ {e}")
            messagebox.showerror("Validation Error", "\n".join(errs))
            return

        self.btn_run.config(state="disabled")
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
            enable_clin = bool(self.use_clin_map.get())
            enable_phase = bool(self.use_phase_map.get())

            self.log(f"Loading input from: {in_path}")
            df_boe = read_boe_summary_multi(in_path)

            self.log("Building paths…")
            paths = build_paths(
                df_boe,
                project_number=proj,
                contract_type=ctype,
                clin_map=clin_map,
                phase_map=phase_map,
                enable_clin_map=enable_clin,
                enable_phase_map=enable_phase,
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
            self.btn_run.config(state="normal")

def main_gui():
    root = TkRootClass()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass

    app = ChargeCodesGUI(root)
    root.geometry("1180x800")
    root.mainloop()

if __name__ == "__main__":
    main_gui()
