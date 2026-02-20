#!/usr/bin/env python3
"""
RetreatPlacer Desktop UI — CustomTkinter GUI for retreat room placement.

Usage:
    pip install customtkinter ortools openpyxl pandas
    python RetreatPlacerUI.py

Place this file alongside RetreatPlacer.py. It imports the solver directly
(no subprocess needed) and runs it in a background thread with live output.
"""

import os
import sys
import io
import threading
import traceback
import platform
import subprocess
import copy
from pathlib import Path
from datetime import datetime
from collections import defaultdict

import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk

import pandas as pd

# ---------------------------------------------------------------------------
# Attempt to import RetreatPlacer — locate it relative to this script
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
SOLVER_PATH = SCRIPT_DIR / "RetreatPlacer.py"

_solver_available = False
try:
    sys.path.insert(0, str(SCRIPT_DIR))
    import RetreatPlacer
    _solver_available = True
except ImportError:
    pass

# ---------------------------------------------------------------------------
# Theme & Constants
# ---------------------------------------------------------------------------
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

ACCENT       = "#e94560"
ACCENT_HOVER = "#c23152"
SUCCESS      = "#52c41a"
WARNING      = "#f6aa1c"
MUTED        = "#888888"
BG_CARD      = "#1a1a2e"
BG_DARK      = "#0f0f1a"
EMPTY_COLOR  = "#333344"
SELECTED_BG  = "#2a2a5e"

BUILDING_COLORS = {
    "Black River":  "#e94560",
    "Cottonwood":   "#52c41a",
    "Cottonwood ":  "#52c41a",
    "Diamond":      "#f6aa1c",
    "Echo Canyon":  "#8da9c4",
    "Four Peaks":   "#f4a261",
}

DEFAULT_BLDG_COLOR = "#74b9ff"


def get_bldg_color(name):
    return BUILDING_COLORS.get(name, BUILDING_COLORS.get(name.strip(), DEFAULT_BLDG_COLOR))


# ---------------------------------------------------------------------------
# Redirect stdout/stderr to a callback
# ---------------------------------------------------------------------------
class OutputCapture(io.StringIO):
    """Captures writes to stdout/stderr and forwards them to a callback."""
    def __init__(self, callback, original):
        super().__init__()
        self.callback = callback
        self.original = original

    def write(self, text):
        if text:
            self.callback(text)
            if self.original:
                self.original.write(text)
        return len(text) if text else 0

    def flush(self):
        if self.original:
            self.original.flush()


# ═══════════════════════════════════════════════════════════════════
# Main Application
# ═══════════════════════════════════════════════════════════════════
class RetreatPlacerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("RetreatPlacer")
        self.geometry("1060x750")
        self.minsize(860, 620)

        # State
        self.room_file = ctk.StringVar(value="")
        self.people_file = ctk.StringVar(value="")
        self.output_file = ctk.StringVar(value="")
        self.is_running = False

        # Solver results (original, from solver)
        self.results = None
        self.unplaced = None
        self.rooms_df = None  # room capacities for empty-slot display

        # Editable placement state — these are mutated by drag-drop
        self.edit_results = None      # list of dicts (same schema as self.results)
        self.edit_unplaced = None     # list of (person_dict, reasons, attach_res)
        self.has_manual_edits = False

        # Click-to-move state
        self._drag_source = None  # ("placed", bldg, room, person) or ("unplaced", idx, person)
        self._last_status = ""

        # Spinner animation state
        self._spinner_job = None
        self._spinner_idx = 0
        self._solver_start_time = None

        # Layout: sidebar + main content
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main()

        # Show the files tab by default
        self._show_tab("files")

    # ═══════════════════════════════════════════════════════════════
    # SIDEBAR
    # ═══════════════════════════════════════════════════════════════
    def _build_sidebar(self):
        sidebar = ctk.CTkFrame(self, width=200, corner_radius=0, fg_color=BG_DARK)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)

        logo_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        logo_frame.pack(fill="x", padx=16, pady=(20, 8))
        ctk.CTkLabel(logo_frame, text="🏕  RetreatPlacer",
                     font=ctk.CTkFont(size=18, weight="bold"),
                     text_color="#ffffff").pack(anchor="w")
        ctk.CTkLabel(logo_frame, text="Room Assignment Solver",
                     font=ctk.CTkFont(size=11), text_color=MUTED).pack(anchor="w", pady=(2, 0))

        ctk.CTkFrame(sidebar, height=1, fg_color="#333333").pack(fill="x", padx=16, pady=(16, 12))

        self.nav_buttons = {}
        for tab_id, label in [("files", "📁  Files"), ("run", "⚡  Run Solver"),
                               ("results", "📊  Results"), ("log", "📝  Console Log")]:
            btn = ctk.CTkButton(
                sidebar, text=label, anchor="w", font=ctk.CTkFont(size=13),
                fg_color="transparent", text_color="#cccccc", hover_color="#222244",
                height=38, corner_radius=8, command=lambda t=tab_id: self._show_tab(t))
            btn.pack(fill="x", padx=10, pady=2)
            self.nav_buttons[tab_id] = btn

        ctk.CTkFrame(sidebar, fg_color="transparent").pack(fill="both", expand=True)

        ctk.CTkLabel(sidebar, text="OR-Tools CP-SAT Solver",
                     font=ctk.CTkFont(size=10), text_color="#444444").pack(padx=16, pady=(0, 4))
        status = "✓ Solver loaded" if _solver_available else "✗ Solver not found"
        ctk.CTkLabel(sidebar, text=status, font=ctk.CTkFont(size=10),
                     text_color=SUCCESS if _solver_available else ACCENT).pack(padx=16, pady=(0, 16))

    # ═══════════════════════════════════════════════════════════════
    # MAIN CONTENT
    # ═══════════════════════════════════════════════════════════════
    def _build_main(self):
        self.main_frame = ctk.CTkFrame(self, fg_color="#111122", corner_radius=0)
        self.main_frame.grid(row=0, column=1, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)

        self.tab_frames = {}
        self.tab_frames["files"] = self._build_files_tab()
        self.tab_frames["run"] = self._build_run_tab()
        self.tab_frames["results"] = self._build_results_tab()
        self.tab_frames["log"] = self._build_log_tab()
        for frame in self.tab_frames.values():
            frame.grid(row=0, column=0, sticky="nsew")

    def _show_tab(self, tab_id):
        for tid, frame in self.tab_frames.items():
            frame.grid_remove()
        self.tab_frames[tab_id].grid()
        for tid, btn in self.nav_buttons.items():
            if tid == tab_id:
                btn.configure(fg_color="#222244", text_color="#ffffff")
            else:
                btn.configure(fg_color="transparent", text_color="#cccccc")

    # ═══════════════════════════════════════════════════════════════
    # FILES TAB
    # ═══════════════════════════════════════════════════════════════
    def _build_files_tab(self):
        frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        frame.grid_columnconfigure(0, weight=1)
        scroll = ctk.CTkScrollableFrame(frame, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=24, pady=16)
        scroll.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(scroll, text="Input & Output Files",
                     font=ctk.CTkFont(size=20, weight="bold"), text_color="#ffffff",
                     anchor="w").grid(row=0, column=0, sticky="w", pady=(0, 4))
        ctk.CTkLabel(scroll, text="Select your room map and people list. The solver will generate the output file.",
                     font=ctk.CTkFont(size=12), text_color=MUTED, anchor="w",
                     wraplength=600).grid(row=1, column=0, sticky="w", pady=(0, 20))

        self._file_picker(scroll, 2, "Room Map", "RoomMap.xlsx",
                          self.room_file, [("Excel", "*.xlsx *.csv")], on_change=self._on_room_file_changed)
        self._file_picker(scroll, 3, "People to Place", "PeopleToPlace.xlsx",
                          self.people_file, [("Excel", "*.xlsx *.csv")], on_change=self._on_people_file_changed)
        self._file_picker(scroll, 4, "Output File (optional)",
                          "FilledRoomMap.xlsx — defaults to same folder as Room Map",
                          self.output_file, [("Excel", "*.xlsx")], save=True)

        self.stats_frame = ctk.CTkFrame(scroll, fg_color=BG_CARD, corner_radius=12)
        self.stats_frame.grid(row=5, column=0, sticky="ew", pady=(20, 0))
        self.stats_label = ctk.CTkLabel(self.stats_frame, text="Upload files to see data summary",
                                        font=ctk.CTkFont(size=12), text_color=MUTED,
                                        wraplength=500, justify="left")
        self.stats_label.pack(padx=20, pady=16, anchor="w")
        return frame

    def _file_picker(self, parent, row, label, sublabel, var, filetypes, save=False, on_change=None):
        card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=12)
        card.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        card.grid_columnconfigure(1, weight=1)

        lbl_frame = ctk.CTkFrame(card, fg_color="transparent")
        lbl_frame.grid(row=0, column=0, padx=(16, 8), pady=14, sticky="w")
        ctk.CTkLabel(lbl_frame, text=label, font=ctk.CTkFont(size=13, weight="bold"),
                     text_color="#dddddd").pack(anchor="w")
        ctk.CTkLabel(lbl_frame, text=sublabel, font=ctk.CTkFont(size=10),
                     text_color="#666666").pack(anchor="w")

        ctk.CTkEntry(card, textvariable=var, state="readonly",
                     font=ctk.CTkFont(family="Courier", size=11),
                     fg_color="#0d0d1a", border_color="#333333", text_color="#aaaaaa",
                     height=32).grid(row=0, column=1, padx=4, pady=14, sticky="ew")

        def browse():
            if save:
                init_dir = str(Path(self.room_file.get()).parent) if self.room_file.get() else ""
                path = filedialog.asksaveasfilename(title=f"Save {label}", filetypes=filetypes,
                                                    defaultextension=".xlsx",
                                                    initialdir=init_dir or None,
                                                    initialfile="FilledRoomMap.xlsx")
            else:
                path = filedialog.askopenfilename(title=f"Select {label}",
                                                  filetypes=filetypes + [("All files", "*.*")])
            if path:
                var.set(path)
                if on_change:
                    on_change(path)

        ctk.CTkButton(card, text="Browse…", width=90, height=32, font=ctk.CTkFont(size=12),
                      fg_color=ACCENT, hover_color=ACCENT_HOVER, corner_radius=8,
                      command=browse).grid(row=0, column=2, padx=(4, 16), pady=14)

    def _on_room_file_changed(self, path):
        try:
            pd.read_excel(path)
            self._update_stats()
        except Exception as e:
            self.stats_label.configure(text=f"Error reading room file: {e}", text_color=ACCENT)

    def _on_people_file_changed(self, path):
        try:
            pd.read_excel(path)
            self._update_stats()
        except Exception as e:
            self.stats_label.configure(text=f"Error reading people file: {e}", text_color=ACCENT)

    def _update_stats(self):
        lines = []
        try:
            if self.room_file.get():
                rdf = pd.read_excel(self.room_file.get())
                rdf.columns = rdf.columns.str.strip()
                bottom = int(rdf['#BottomBunk'].sum())
                top = int(rdf['#TopBunk'].sum())
                bldgs = sorted(rdf['BuildingName'].str.strip().unique())
                lines.append(f"🏠  {len(rdf)} rooms across {len(bldgs)} buildings  •  "
                             f"{bottom + top} total beds ({bottom} bottom, {top} top)")
                lines.append(f"     Buildings: {', '.join(bldgs)}")
        except Exception:
            pass
        try:
            if self.people_file.get():
                pdf = pd.read_excel(self.people_file.get())
                pdf.columns = pdf.columns.str.strip()
                for col in ['OrgName', 'GroupName', 'AttachName', 'BunkPref', 'RoomLocationPref']:
                    if col in pdf.columns:
                        pdf[col] = pdf[col].fillna('').astype(str).str.strip()
                orgs = sorted(set(pdf['OrgName'].unique()) - {'', 'nan', 'NaN'})
                groups = sorted(set(pdf['GroupName'].unique()) - {'', 'nan', 'NaN'})
                floor1 = sum((pdf['RoomLocationPref'] == '1') | (pdf['RoomLocationPref'] == 1))
                bottom = sum(pdf['BunkPref'].str.lower() == 'bottom')
                attached = sum((pdf['AttachName'] != '') & (pdf['AttachName'] != 'nan'))
                lines.append(f"\n👥  {len(pdf)} people  •  {len(orgs)} orgs  •  {len(groups)} groups")
                lines.append(f"     Orgs: {', '.join(orgs)}")
                if groups:
                    lines.append(f"     Groups: {', '.join(groups)}")
                lines.append(f"     Constraints: {floor1} floor 1  •  {bottom} bottom  •  {attached} attached")
        except Exception:
            pass
        if lines:
            self.stats_label.configure(text="\n".join(lines), text_color="#cccccc")
        else:
            self.stats_label.configure(text="Upload files to see data summary", text_color=MUTED)

    # ═══════════════════════════════════════════════════════════════
    # RUN TAB
    # ═══════════════════════════════════════════════════════════════
    def _build_run_tab(self):
        frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        inner = ctk.CTkFrame(frame, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=24, pady=16)

        ctk.CTkLabel(inner, text="Run Solver", font=ctk.CTkFont(size=20, weight="bold"),
                     text_color="#ffffff", anchor="w").pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(inner,
                     text="Click the button below to run the OR-Tools CP-SAT solver. "
                          "You'll be taken to the Console Log to watch progress.",
                     font=ctk.CTkFont(size=12), text_color=MUTED, anchor="w",
                     wraplength=600, justify="left").pack(fill="x", pady=(0, 20))

        config_card = ctk.CTkFrame(inner, fg_color=BG_CARD, corner_radius=12)
        config_card.pack(fill="x", pady=(0, 16))
        ctk.CTkLabel(config_card, text="⚙  Solver Configuration",
                     font=ctk.CTkFont(size=13, weight="bold"), text_color=WARNING,
                     anchor="w").pack(padx=20, pady=(16, 8), anchor="w")
        for key, val in [("Engine", "Google OR-Tools CP-SAT"), ("Time Limit", "600 seconds"),
                         ("Workers", "8 threads"), ("Placement Weight", "10,000 per person"),
                         ("Group Cohesion", "1,000 per pair"),
                         ("Attach Weight", "800 per pair (one-directional)"),
                         ("Building Affinity", "200 per person"), ("Org Cohesion", "100 per pair")]:
            rf = ctk.CTkFrame(config_card, fg_color="transparent")
            rf.pack(fill="x", padx=20, pady=1)
            ctk.CTkLabel(rf, text=key, width=160, font=ctk.CTkFont(size=11),
                         text_color="#666666", anchor="w").pack(side="left")
            ctk.CTkLabel(rf, text=val, font=ctk.CTkFont(family="Courier", size=11),
                         text_color="#cccccc", anchor="w").pack(side="left")
        ctk.CTkFrame(config_card, height=12, fg_color="transparent").pack()

        self.run_button = ctk.CTkButton(
            inner, text="▶  Run Solver", height=48,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color=ACCENT, hover_color=ACCENT_HOVER, corner_radius=10,
            command=self._run_solver)
        self.run_button.pack(fill="x", pady=(8, 0))

        self.run_status_label = ctk.CTkLabel(
            inner, text="", font=ctk.CTkFont(size=12), text_color=MUTED,
            anchor="w", wraplength=600, justify="left")
        self.run_status_label.pack(fill="x", pady=(12, 0))

        self.progress_bar = ctk.CTkProgressBar(inner, mode="indeterminate",
                                                progress_color=ACCENT, height=4)
        self.progress_bar.pack(fill="x", pady=(8, 0))
        self.progress_bar.set(0)
        return frame

    def _run_solver(self):
        if self.is_running:
            return
        room_path = self.room_file.get()
        people_path = self.people_file.get()
        if not room_path or not os.path.isfile(room_path):
            messagebox.showerror("Missing File", "Please select a valid Room Map file.")
            self._show_tab("files")
            return
        if not people_path or not os.path.isfile(people_path):
            messagebox.showerror("Missing File", "Please select a valid People to Place file.")
            self._show_tab("files")
            return
        if not _solver_available:
            messagebox.showerror("Solver Not Found",
                                 f"Could not import RetreatPlacer.py.\n\n"
                                 f"Expected: {SOLVER_PATH}\n\n"
                                 f"Place RetreatPlacerUI.py alongside RetreatPlacer.py.")
            return

        output_path = self.output_file.get()
        if not output_path:
            output_path = str(Path(room_path).parent / "FilledRoomMap.xlsx")
            self.output_file.set(output_path)

        # [CHANGE 1] Clear log and auto-switch to Console Log tab
        self._clear_log()
        self._show_tab("log")

        self.is_running = True
        self.run_button.configure(state="disabled", text="⏳  Solver Running…")
        self.progress_bar.start()
        self.run_status_label.configure(text="Solver is running… switched to Console Log.",
                                        text_color=WARNING)

        # [CHANGE 2] Start running indicator in log tab
        self._start_log_spinner()

        threading.Thread(target=self._solver_thread,
                         args=(room_path, people_path, output_path), daemon=True).start()

    def _solver_thread(self, room_path, people_path, output_path):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        def on_output(text):
            self.after(0, self._append_log, text)
        sys.stdout = OutputCapture(on_output, old_stdout)
        sys.stderr = OutputCapture(on_output, old_stderr)
        try:
            on_output(f"{'='*60}\n")
            on_output(f"RetreatPlacer — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            on_output(f"{'='*60}\n\n")
            on_output(f"Room map  : {room_path}\nPeople    : {people_path}\nOutput    : {output_path}\n\n")

            rooms_df, people_df = RetreatPlacer.load_data(room_path, people_path)
            slots, _ = RetreatPlacer.build_slots(rooms_df)
            on_output(f"Rooms: {len(rooms_df)}  |  People: {len(people_df)}  |  "
                      f"Slots: {len(slots)} (Bottom: {sum(1 for s in slots if s[3]=='Bottom')}, "
                      f"Top: {sum(1 for s in slots if s[3]=='Top')})\n\nSolving with OR-Tools CP-SAT ...\n")

            results, unplaced, slots, attach_warnings, resolved_attach = \
                RetreatPlacer.solve_placement(rooms_df, people_df)
            RetreatPlacer.print_debug(results, unplaced, slots)

            if results is not None:
                RetreatPlacer.write_output(results, unplaced, output_path, attach_warnings, resolved_attach)
                on_output(f"\n✓ Output saved to: {output_path}\n")
                self.results = results
                self.unplaced = unplaced
                self.rooms_df = rooms_df
                self.after(0, self._on_solver_done, True, output_path, len(results), len(unplaced))
            else:
                on_output("\n✗ No placements — no output file written.\n")
                self.after(0, self._on_solver_done, False, "", 0, 0)
        except Exception as e:
            on_output(f"\n✗ ERROR: {e}\n{traceback.format_exc()}\n")
            self.after(0, self._on_solver_done, False, "", 0, 0)
        finally:
            sys.stdout, sys.stderr = old_stdout, old_stderr

    def _on_solver_done(self, success, output_path, placed, unplaced_count):
        self.is_running = False
        self.run_button.configure(state="normal", text="▶  Run Solver")
        self.progress_bar.stop()
        self.progress_bar.set(0)
        self._stop_log_spinner()

        if success:
            self.run_status_label.configure(
                text=f"✓ Complete — {placed} placed, {unplaced_count} unplaced. Output: {Path(output_path).name}",
                text_color=SUCCESS)
            self.edit_results = copy.deepcopy(self.results)
            self.edit_unplaced = copy.deepcopy(self.unplaced)
            self.has_manual_edits = False
            self._drag_source = None
            self._populate_results_tab()
        else:
            self.run_status_label.configure(text="✗ Solver failed. Check Console Log.", text_color=ACCENT)

    # ═══════════════════════════════════════════════════════════════
    # RESULTS TAB  [CHANGES 3, 4, 5]
    # ═══════════════════════════════════════════════════════════════
    def _build_results_tab(self):
        frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.results_scroll = ctk.CTkScrollableFrame(frame, fg_color="transparent")
        self.results_scroll.pack(fill="both", expand=True, padx=24, pady=16)
        self.results_scroll.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.results_scroll,
                     text="No results yet.\n\nRun the solver to see results here.",
                     font=ctk.CTkFont(size=13), text_color=MUTED, justify="center").pack(pady=80)
        return frame

    def _get_room_capacity(self, building_name, room_name):
        if self.rooms_df is None:
            return (0, 0)
        match = self.rooms_df[
            (self.rooms_df['BuildingName'].str.strip() == building_name.strip()) &
            (self.rooms_df['RoomName'].str.strip() == room_name.strip())]
        if match.empty:
            return (0, 0)
        row = match.iloc[0]
        return (int(row['#BottomBunk']), int(row['#TopBunk']))

    def _populate_results_tab(self):
        for widget in self.results_scroll.winfo_children():
            widget.destroy()

        results = self.edit_results
        unplaced = self.edit_unplaced
        if not results:
            ctk.CTkLabel(self.results_scroll, text="No results available.",
                         font=ctk.CTkFont(size=13), text_color=MUTED).pack(pady=80)
            return

        # ── Header with unsaved indicator
        hdr = ctk.CTkFrame(self.results_scroll, fg_color="transparent")
        hdr.pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(hdr, text="Placement Results",
                     font=ctk.CTkFont(size=20, weight="bold"),
                     text_color="#ffffff", anchor="w").pack(side="left")
        if self.has_manual_edits:
            ctk.CTkLabel(hdr, text="  ●  Unsaved changes",
                         font=ctk.CTkFont(size=12, weight="bold"),
                         text_color=WARNING, anchor="w").pack(side="left", padx=(8, 0))

        # ── Instructions + selection status
        status_text = self._last_status or \
            "Click a person to select them, then click an empty slot to move them."
        self._results_status_lbl = ctk.CTkLabel(
            self.results_scroll, text=status_text,
            font=ctk.CTkFont(size=11), text_color="#888866" if self._drag_source else "#666666",
            anchor="w", wraplength=700, justify="left")
        self._results_status_lbl.pack(fill="x", pady=(0, 12))

        # ── Deselect button (shown when something is selected)
        if self._drag_source:
            desel_frame = ctk.CTkFrame(self.results_scroll, fg_color="transparent")
            desel_frame.pack(fill="x", pady=(0, 8))
            ctk.CTkButton(desel_frame, text="✕  Deselect", height=26, width=100,
                          font=ctk.CTkFont(size=11), fg_color="#444444",
                          hover_color="#555555", corner_radius=6,
                          command=self._deselect).pack(side="left")

        # ── Summary stats
        unplaced_count = len(unplaced) if unplaced else 0
        stats_row = ctk.CTkFrame(self.results_scroll, fg_color="transparent")
        stats_row.pack(fill="x", pady=(0, 16))
        for label, value, color in [
            ("Placed", str(len(results)), SUCCESS),
            ("Unplaced", str(unplaced_count), ACCENT if unplaced_count > 0 else SUCCESS),
            ("Buildings", str(len(set(r['BuildingName'] for r in results))), "#8da9c4"),
            ("Rooms Used", str(len(set(f"{r['BuildingName']}|{r['RoomName']}" for r in results))), WARNING),
        ]:
            card = ctk.CTkFrame(stats_row, fg_color=BG_CARD, corner_radius=10)
            card.pack(side="left", fill="x", expand=True, padx=(0, 8))
            ctk.CTkLabel(card, text=value, font=ctk.CTkFont(size=28, weight="bold"),
                         text_color=color).pack(padx=16, pady=(12, 2))
            ctk.CTkLabel(card, text=label, font=ctk.CTkFont(size=11),
                         text_color=MUTED).pack(padx=16, pady=(0, 12))

        # ── Building breakdown with empty slots [CHANGE 3]
        by_building = defaultdict(lambda: defaultdict(list))
        for r in results:
            by_building[r['BuildingName']][r['RoomName']].append(r)

        # Include all rooms from rooms_df (even completely empty ones)
        if self.rooms_df is not None:
            for _, row in self.rooms_df.iterrows():
                bldg, rm = row['BuildingName'], row['RoomName']
                if rm not in by_building[bldg]:
                    by_building[bldg][rm] = []

        for bldg_name in sorted(by_building.keys()):
            rooms = by_building[bldg_name]
            bldg_color = get_bldg_color(bldg_name)
            total_in_bldg = sum(len(ppl) for ppl in rooms.values())
            total_cap_bldg = sum(sum(self._get_room_capacity(bldg_name, rn))
                                 for rn in rooms.keys())

            bldg_card = ctk.CTkFrame(self.results_scroll, fg_color=BG_CARD, corner_radius=12,
                                     border_width=1, border_color=bldg_color)
            bldg_card.pack(fill="x", pady=(0, 12))

            hdr = ctk.CTkFrame(bldg_card, fg_color="transparent")
            hdr.pack(fill="x", padx=16, pady=(12, 8))
            ctk.CTkLabel(hdr, text=f"  {bldg_name.strip()}",
                         font=ctk.CTkFont(size=14, weight="bold"),
                         text_color=bldg_color, anchor="w").pack(side="left")
            ctk.CTkLabel(hdr, text=f"{total_in_bldg}/{total_cap_bldg} beds · {len(rooms)} rooms",
                         font=ctk.CTkFont(size=11), text_color=MUTED, anchor="e").pack(side="right")

            # Org breakdown
            org_counts = defaultdict(int)
            for rp in rooms.values():
                for p in rp:
                    org_counts[p.get('OrgName', '') or '(none)'] += 1
            if org_counts:
                org_text = "  •  ".join(f"{o}: {c}" for o, c in sorted(org_counts.items()))
                ctk.CTkLabel(bldg_card, text=f"     {org_text}", font=ctk.CTkFont(size=10),
                             text_color="#666666", anchor="w", wraplength=700,
                             justify="left").pack(fill="x", padx=16, pady=(0, 4))

            # Rooms grid
            rooms_frame = ctk.CTkFrame(bldg_card, fg_color="transparent")
            rooms_frame.pack(fill="x", padx=12, pady=(4, 12))
            col = 0
            for room_name in sorted(rooms.keys()):
                people_in_room = rooms[room_name]
                bottom_cap, top_cap = self._get_room_capacity(bldg_name, room_name)
                total_cap = bottom_cap + top_cap

                room_frame = ctk.CTkFrame(rooms_frame, fg_color="#0d0d1a", corner_radius=8,
                                          border_width=1, border_color="#222222")
                room_frame.grid(row=col // 4, column=col % 4, padx=4, pady=4, sticky="nsew")
                rooms_frame.grid_columnconfigure(col % 4, weight=1)

                # Room header with occupancy
                room_hdr = ctk.CTkFrame(room_frame, fg_color="transparent")
                room_hdr.pack(fill="x", padx=8, pady=(6, 2))
                ctk.CTkLabel(room_hdr, text=room_name, font=ctk.CTkFont(size=11, weight="bold"),
                             text_color="#cccccc").pack(side="left")
                occ_color = SUCCESS if len(people_in_room) == total_cap else (
                    "#666666" if len(people_in_room) < total_cap else ACCENT)
                ctk.CTkLabel(room_hdr, text=f"{len(people_in_room)}/{total_cap}",
                             font=ctk.CTkFont(family="Courier", size=9),
                             text_color=occ_color).pack(side="right")

                # Separate into bottom and top bunk occupants
                bottom_people = [p for p in people_in_room if p.get('Bunk') == 'Bottom']
                top_people = [p for p in people_in_room if p.get('Bunk') != 'Bottom']

                # Bottom bunks
                for person in bottom_people:
                    self._render_person_slot(room_frame, person, bldg_name, room_name,
                                             "Bottom", bldg_color)
                for _ in range(bottom_cap - len(bottom_people)):
                    self._render_empty_slot(room_frame, bldg_name, room_name, "Bottom")

                # Top bunks
                for person in top_people:
                    self._render_person_slot(room_frame, person, bldg_name, room_name,
                                             "Top", bldg_color)
                for _ in range(top_cap - len(top_people)):
                    self._render_empty_slot(room_frame, bldg_name, room_name, "Top")

                ctk.CTkFrame(room_frame, height=2, fg_color="transparent").pack()
                col += 1

        # ── Unplaced section
        if unplaced and len(unplaced) > 0:
            ctk.CTkFrame(self.results_scroll, height=2, fg_color=ACCENT).pack(fill="x", pady=(16, 8))
            ctk.CTkLabel(self.results_scroll, text=f"⚠  Unplaced People ({len(unplaced)})",
                         font=ctk.CTkFont(size=16, weight="bold"), text_color=ACCENT,
                         anchor="w").pack(fill="x", pady=(8, 4))
            ctk.CTkLabel(self.results_scroll,
                         text="Click a person to select, then click an empty slot above to place them.",
                         font=ctk.CTkFont(size=11), text_color="#666666",
                         anchor="w").pack(fill="x", pady=(0, 8))

            for up_idx, (person, reasons, attach_res) in enumerate(unplaced):
                name = f"{person.get('FirstName', '')} {person.get('LastName', '')}"
                org = person.get('OrgName', '')
                grp = person.get('GroupName', '')
                meta = [x for x in [org, grp] if x]
                suffix = f"  ({', '.join(meta)})" if meta else ""

                # Highlight if this unplaced person is currently selected
                is_selected = (self._drag_source and self._drag_source[0] == "unplaced"
                               and self._drag_source[1] == up_idx)
                lbl = ctk.CTkLabel(
                    self.results_scroll, text=f"  {'● ' if is_selected else '○  '}{name}{suffix}",
                    font=ctk.CTkFont(size=11, weight="bold" if is_selected else "normal"),
                    text_color=WARNING if is_selected else "#dd8888",
                    anchor="w", cursor="hand2")
                lbl.pack(fill="x", padx=8, pady=1)
                lbl.bind("<Button-1>", lambda e, i=up_idx: self._on_click_unplaced(i))

                # Show reasons on hover / compact
                if reasons:
                    reason_text = "; ".join(reasons) if isinstance(reasons, list) else str(reasons)
                    ctk.CTkLabel(self.results_scroll, text=f"       {reason_text}",
                                 font=ctk.CTkFont(size=9), text_color="#555555",
                                 anchor="w", wraplength=650, justify="left").pack(fill="x", padx=8)

        # ── Bottom bar [CHANGE 5]
        ctk.CTkFrame(self.results_scroll, height=1, fg_color="#333333").pack(fill="x", pady=(20, 12))
        bottom_bar = ctk.CTkFrame(self.results_scroll, fg_color="transparent")
        bottom_bar.pack(fill="x", pady=(0, 8))

        save_active = self.has_manual_edits
        ctk.CTkButton(
            bottom_bar, text="💾  Save Changes to Output File", height=42,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=SUCCESS if save_active else "#333344",
            hover_color="#3d9918" if save_active else "#444466",
            text_color="#ffffff" if save_active else "#888888",
            corner_radius=10, command=self._save_manual_changes
        ).pack(side="left", padx=(0, 8))

        if self.output_file.get() and os.path.isfile(self.output_file.get()):
            ctk.CTkButton(bottom_bar, text="📂  Open Output File", height=42,
                          font=ctk.CTkFont(size=13), fg_color=ACCENT, hover_color=ACCENT_HOVER,
                          corner_radius=10, command=self._open_output_file).pack(side="left", padx=(0, 8))
            ctk.CTkButton(bottom_bar, text="📁  Open Folder", height=42,
                          font=ctk.CTkFont(size=13), fg_color="#333344", hover_color="#444466",
                          corner_radius=10, command=self._open_output_folder).pack(side="left")

        self._show_tab("results")

    # ─── Slot rendering ──────────────────────────────────────────

    def _render_person_slot(self, parent, person, bldg, room, bunk_type, bldg_color):
        indicator = "▾" if bunk_type == 'Bottom' else "▴"
        color = bldg_color if bunk_type == 'Bottom' else MUTED
        name = f"{person['FirstName']} {person['LastName']}"
        org = person.get('OrgName', '')
        suffix = f"  ({org})" if org else ""

        # Highlight if selected
        is_selected = (self._drag_source and self._drag_source[0] == "placed"
                       and self._drag_source[1] == bldg and self._drag_source[2] == room
                       and self._drag_source[3]['FirstName'] == person['FirstName']
                       and self._drag_source[3]['LastName'] == person['LastName'])

        lbl = ctk.CTkLabel(
            parent, text=f" {indicator} {name}{suffix}",
            font=ctk.CTkFont(size=10, weight="bold" if is_selected else "normal"),
            text_color=WARNING if is_selected else color,
            anchor="w", cursor="hand2")
        lbl.pack(padx=8, pady=0, anchor="w")
        lbl.bind("<Button-1>", lambda e, b=bldg, r=room, p=person: self._on_click_person(b, r, p))

    def _render_empty_slot(self, parent, bldg, room, bunk_type):
        indicator = "▾" if bunk_type == 'Bottom' else "▴"

        # Highlight empty slots when a person is selected (drop targets)
        is_target = self._drag_source is not None
        lbl = ctk.CTkLabel(
            parent, text=f" {indicator} {'┄┄ drop here ┄┄' if is_target else '── empty ──'}",
            font=ctk.CTkFont(size=10),
            text_color=WARNING if is_target else EMPTY_COLOR,
            anchor="w", cursor="hand2" if is_target else "arrow")
        lbl.pack(padx=8, pady=0, anchor="w")
        lbl.bind("<Button-1>", lambda e, b=bldg, r=room, bt=bunk_type: self._on_click_empty(b, r, bt))

    # ─── Click-to-move logic [CHANGE 4] ─────────────────────────

    def _on_click_person(self, bldg, room, person):
        name = f"{person['FirstName']} {person['LastName']}"
        self._drag_source = ("placed", bldg, room, person)
        self._last_status = f"Selected: {name}  —  click an empty slot to move, or click another person."
        self._populate_results_tab()

    def _on_click_unplaced(self, unplaced_idx):
        if self.edit_unplaced and 0 <= unplaced_idx < len(self.edit_unplaced):
            person = self.edit_unplaced[unplaced_idx][0]
            name = f"{person.get('FirstName', '')} {person.get('LastName', '')}"
            self._drag_source = ("unplaced", unplaced_idx, person)
            self._last_status = f"Selected unplaced: {name}  —  click an empty slot to place them."
            self._populate_results_tab()

    def _on_click_empty(self, bldg, room, bunk_type):
        if self._drag_source is None:
            self._last_status = "Click a person first, then click an empty slot."
            self._populate_results_tab()
            return

        src = self._drag_source[0]

        if src == "placed":
            _, src_bldg, src_room, person = self._drag_source
            # Remove from source
            self.edit_results = [
                r for r in self.edit_results
                if not (r['BuildingName'] == src_bldg and r['RoomName'] == src_room
                        and r['FirstName'] == person['FirstName']
                        and r['LastName'] == person['LastName'])]
            # Add to target
            new_entry = dict(person)
            new_entry['BuildingName'] = bldg
            new_entry['RoomName'] = room
            new_entry['Bunk'] = bunk_type
            if self.rooms_df is not None:
                match = self.rooms_df[
                    (self.rooms_df['BuildingName'].str.strip() == bldg.strip()) &
                    (self.rooms_df['RoomName'].str.strip() == room.strip())]
                if not match.empty:
                    new_entry['RoomFloor'] = int(match.iloc[0]['RoomFloor'])
            self.edit_results.append(new_entry)
            name = f"{person['FirstName']} {person['LastName']}"
            self._last_status = f"✓ Moved {name} → {bldg.strip()} / {room} ({bunk_type})"

        elif src == "unplaced":
            _, up_idx, person = self._drag_source
            if 0 <= up_idx < len(self.edit_unplaced):
                self.edit_unplaced.pop(up_idx)
            new_entry = {
                'BuildingName': bldg, 'RoomName': room,
                'FirstName': person.get('FirstName', ''), 'LastName': person.get('LastName', ''),
                'OrgName': person.get('OrgName', ''), 'GroupName': person.get('GroupName', ''),
                'RoomFloor': 1, 'Bunk': bunk_type,
                'AttachName': person.get('AttachName', ''), 'AttachResolved': '',
            }
            if self.rooms_df is not None:
                match = self.rooms_df[
                    (self.rooms_df['BuildingName'].str.strip() == bldg.strip()) &
                    (self.rooms_df['RoomName'].str.strip() == room.strip())]
                if not match.empty:
                    new_entry['RoomFloor'] = int(match.iloc[0]['RoomFloor'])
            self.edit_results.append(new_entry)
            name = f"{person.get('FirstName', '')} {person.get('LastName', '')}"
            self._last_status = f"✓ Placed {name} → {bldg.strip()} / {room} ({bunk_type})"

        self._drag_source = None
        self.has_manual_edits = True
        self._populate_results_tab()

    def _deselect(self):
        self._drag_source = None
        self._last_status = ""
        self._populate_results_tab()

    # ─── Save manual changes [CHANGE 5] ─────────────────────────

    def _save_manual_changes(self):
        if not self.has_manual_edits:
            messagebox.showinfo("No Changes", "No manual changes to save.")
            return
        output_path = self.output_file.get()
        if not output_path:
            output_path = filedialog.asksaveasfilename(
                title="Save Output File", filetypes=[("Excel", "*.xlsx")],
                defaultextension=".xlsx", initialfile="FilledRoomMap.xlsx")
            if not output_path:
                return
            self.output_file.set(output_path)
        try:
            unplaced_for_write = []
            if self.edit_unplaced:
                for item in self.edit_unplaced:
                    if isinstance(item, tuple) and len(item) == 3:
                        unplaced_for_write.append(item)
                    else:
                        unplaced_for_write.append((item, ["Manually unplaced"], ""))
            RetreatPlacer.write_output(self.edit_results, unplaced_for_write, output_path, [], None)
            self.has_manual_edits = False
            self._last_status = f"✓ Saved to {Path(output_path).name}"
            messagebox.showinfo("Saved", f"Output saved to:\n{output_path}")
            self._populate_results_tab()
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save:\n{e}")

    def _open_output_file(self):
        path = self.output_file.get()
        if path and os.path.isfile(path):
            self._open_file_cross_platform(path)

    def _open_output_folder(self):
        path = self.output_file.get()
        if path:
            self._open_file_cross_platform(str(Path(path).parent))

    @staticmethod
    def _open_file_cross_platform(path):
        system = platform.system()
        try:
            if system == "Darwin":
                subprocess.Popen(["open", path])
            elif system == "Windows":
                os.startfile(path)
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open: {e}")

    # ═══════════════════════════════════════════════════════════════
    # LOG TAB  [CHANGE 2: running indicator with spinner + timer]
    # ═══════════════════════════════════════════════════════════════
    def _build_log_tab(self):
        frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(2, weight=1)

        hdr = ctk.CTkFrame(frame, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", padx=24, pady=(16, 0))
        ctk.CTkLabel(hdr, text="Console Log", font=ctk.CTkFont(size=20, weight="bold"),
                     text_color="#ffffff", anchor="w").pack(side="left")
        ctk.CTkButton(hdr, text="Clear", width=60, height=28, font=ctk.CTkFont(size=11),
                      fg_color="#333344", hover_color="#444466", corner_radius=6,
                      command=self._clear_log).pack(side="right")

        # Running indicator bar [CHANGE 2]
        self.log_indicator_frame = ctk.CTkFrame(frame, fg_color=BG_CARD, corner_radius=8, height=36)
        self.log_indicator_frame.grid(row=1, column=0, sticky="ew", padx=24, pady=(8, 4))

        self.log_spinner_label = ctk.CTkLabel(
            self.log_indicator_frame, text="",
            font=ctk.CTkFont(family="Courier", size=12, weight="bold"),
            text_color=ACCENT, anchor="w")
        self.log_spinner_label.pack(side="left", padx=(12, 0))

        self.log_elapsed_label = ctk.CTkLabel(
            self.log_indicator_frame, text="",
            font=ctk.CTkFont(family="Courier", size=11), text_color=MUTED, anchor="w")
        self.log_elapsed_label.pack(side="left", padx=(8, 0))

        self.log_progress_bar = ctk.CTkProgressBar(
            self.log_indicator_frame, mode="indeterminate",
            progress_color=ACCENT, height=3, width=180)
        self.log_progress_bar.pack(side="right", padx=(0, 12))
        self.log_progress_bar.set(0)

        self.log_indicator_frame.grid_remove()  # hidden until solver runs

        self.log_text = ctk.CTkTextbox(
            frame, font=ctk.CTkFont(family="Courier", size=11),
            fg_color="#0d0d1a", text_color="#cccccc",
            border_width=1, border_color="#222222",
            corner_radius=10, wrap="word", state="disabled")
        self.log_text.grid(row=2, column=0, sticky="nsew", padx=24, pady=(0, 16))
        return frame

    def _start_log_spinner(self):
        self._spinner_idx = 0
        self._solver_start_time = datetime.now()
        self.log_indicator_frame.grid()
        self.log_progress_bar.start()
        self._animate_spinner()

    def _animate_spinner(self):
        if not self.is_running:
            return
        frames = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        self._spinner_idx = (self._spinner_idx + 1) % len(frames)
        self.log_spinner_label.configure(text=f"{frames[self._spinner_idx]}  Solver running…")
        elapsed = datetime.now() - self._solver_start_time
        mins, secs = divmod(int(elapsed.total_seconds()), 60)
        self.log_elapsed_label.configure(text=f"{mins:02d}:{secs:02d}")
        self._spinner_job = self.after(100, self._animate_spinner)

    def _stop_log_spinner(self):
        if self._spinner_job:
            self.after_cancel(self._spinner_job)
            self._spinner_job = None
        self.log_progress_bar.stop()
        self.log_progress_bar.set(0)
        if self._solver_start_time:
            elapsed = datetime.now() - self._solver_start_time
            mins, secs = divmod(int(elapsed.total_seconds()), 60)
            self.log_spinner_label.configure(text="✓  Solver finished", text_color=SUCCESS)
            self.log_elapsed_label.configure(text=f"({mins:02d}:{secs:02d})")
        self.after(2000, self._hide_spinner_and_go_results)

    def _hide_spinner_and_go_results(self):
        self.log_indicator_frame.grid_remove()
        if self.edit_results:
            self._show_tab("results")

    def _append_log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")


# ═══════════════════════════════════════════════════════════════════
# Entry Point
# ═══════════════════════════════════════════════════════════════════
def main():
    app = RetreatPlacerApp()
    app.mainloop()


if __name__ == "__main__":
    main()