import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv
from datetime import datetime
import sys
import os


# Main application class for the GUI
class KXTaggerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Kayak Cross Tagger")
        self.bg_color = "#212121"
        self.root.resizable(True, True)
        self.root.configure(bg=self.bg_color)

        # --- APP ICON CONFIGURATION ---
        try:
            if getattr(sys, 'frozen', False):
                application_path = os.path.dirname(sys.executable)
            else:
                application_path = os.path.dirname(os.path.abspath(__file__))

            icon_path = os.path.join(application_path, "app_icon.icns")

            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
            else:
                print(f"Warning: Icon file not found at {icon_path}")
        except Exception as e:
            print(f"Error loading app icon: {e}")

        # --- WIDGET STYLING ---
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("TFrame", background=self.bg_color)
        self.style.configure(
            "TLabel",
            background=self.bg_color,
            foreground="black",
            font=("Space Mono", 10, "bold"),
        )

        # Standard Button Styles
        self.style.configure("TButton", font=("Space Mono", 14, "bold"), padding=10)
        self.style.map(
            "TButton",
            background=[("!active", "#CCCCCC"), ("active", "#AAAAAA")],
            foreground=[("!active", "black")],
        )
        self.style.configure(
            "Small.TButton", font=("Space Mono", 12, "bold"), padding=5
        )
        self.style.map(
            "Small.TButton",
            background=[("!active", "#CCCCCC"), ("active", "#AAAAAA")],
            foreground=[("!active", "black")],
        )

        # Style for the Instructions Button
        self.style.configure(
            "Purple.TButton",
            font=("Space Mono", 16, "bold"),
            padding=5,
            background="#9B40CD",
            foreground="black",
        )
        self.style.map(
            "Purple.TButton",
            background=[("!active", "#9B40CD"), ("active", "#C9A9C9")],
        )

        # Action Button Styles
        self.style.configure(
            "Small.Red.TButton", font=("Space Mono", 12, "bold"), padding=5
        )
        self.style.map(
            "Small.Red.TButton",
            background=[("!active", "#B22222"), ("active", "#8B0000")],
            foreground=[("!active", "white")],
        )
        self.style.configure(
            "Small.Green.TButton", font=("Space Mono", 12, "bold"), padding=5
        )
        self.style.map(
            "Small.Green.TButton",
            background=[("!active", "#228B22"), ("active", "#006400")],
            foreground=[("!active", "white")],
        )

        # Style for selected/active buttons
        self.style.configure(
            "Active.TButton",
            font=("Space Mono", 14, "bold"),
            padding=10,
            background="#FF8C00",
        )
        self.style.configure(
            "Small.Active.TButton",
            font=("Space Mono", 12, "bold"),
            padding=5,
            background="#FF8C00",
        )

        # New smaller style for action buttons
        self.style.configure(
            "ExtraSmall.TButton", font=("Space Mono", 10, "bold"), padding=4
        )
        self.style.map(
            "ExtraSmall.TButton",
            background=[("!active", "#CCCCCC"), ("active", "#AAAAAA")],
            foreground=[("!active", "black")],
        )

        # Up/Down buttons always red/green
        self.style.configure(
            "ExtraSmall.Red.TButton",
            font=("Space Mono", 10, "bold"),
            padding=4,
            background="#B22222",
            foreground="white",
        )
        self.style.map(
            "ExtraSmall.Red.TButton",
            background=[("!active", "#B22222")],
            foreground=[("!active", "white")],
        )

        self.style.configure(
            "ExtraSmall.Green.TButton",
            font=("Space Mono", 10, "bold"),
            padding=4,
            background="#228B22",
            foreground="white",
        )
        self.style.map(
            "ExtraSmall.Green.TButton",
            background=[("!active", "#228B22")],
            foreground=[("!active", "white")],
        )

        # Active styles
        self.style.configure(
            "ExtraSmall.Active.TButton",
            font=("Space Mono", 10, "bold"),
            padding=4,
            background="#FF8C00",
        )
        self.style.configure(
            "ExtraSmall.Red.Active.TButton",
            font=("Space Mono", 10, "bold"),
            padding=4,
            background="#DC143C",
            foreground="white",
        )
        self.style.configure(
            "ExtraSmall.Green.Active.TButton",
            font=("Space Mono", 10, "bold"),
            padding=4,
            background="#32CD32",
            foreground="white",
        )

        # Combobox Style
        self.style.configure("TCombobox", font=("Space Mono", 9), padding=3)
        self.root.option_add("*TCombobox*Listbox.background", "white")
        self.root.option_add("*TCombobox*Listbox.foreground", "black")
        self.root.option_add("*TCombobox*Listbox.selectBackground", "#ddd")
        self.root.option_add("*TCombobox*Listbox.selectForeground", "black")
        self.style.map(
            "TCombobox",
            fieldbackground=[("readonly", "white")],
            selectbackground=[("readonly", "white")],
            selectforeground=[("readonly", "black")],
            foreground=[("readonly", "black")],
        )

        # --- DATA & STATE MANAGEMENT ---
        self.tagged_data = []
        self.extra_headers = [] # To store custom columns
        self.standard_headers = [
            "Year", "Competition", "Gender", "Phase", "Gate", "BIB", "Ramp Position",
            "Action", "Order", "Final Position", "Upstream Tactic", "Athlete Name", "Faults"
        ]
        self.num_paddlers_var = tk.IntVar(value=4)
        self.paddler_ramp_positions = {}
        self.disabled_positions = set()
        self.faulted_bibs = set()
        self.dns_bibs = set() # Track DNS bibs to exclude from finish count
        self.selected_paddler_setup = None
        self.selected_gate = None
        self.selected_actions = set()
        self.paddler_order_sequence = []
        self.finish_line_sequence = []
        self.phase_final_positions = {}
        self.upstream_tactic_actions = []
        self.paddler_buttons = {}
        self.ramp_position_buttons = {}
        self.gate_buttons = {}
        self.gate_order = [] 
        self.action_buttons = {}
        
        # State for athlete name handling
        self.athlete_name_vars = {p_key: tk.StringVar(self.root) for p_key in ["P1", "P2", "P3", "P4"]}
        self.athlete_name_comboboxes = {}
        self.male_athlete_names = set()
        self.female_athlete_names = set()

        # Bib data mapping internal P-number to UI name, color, and CSV character
        self.bib_data = {
            "P1": {"name": "RED", "color": "#E5296B", "csv_char": "R"},
            "P2": {"name": "GREEN", "color": "#0BC2A3", "csv_char": "G"},
            "P3": {"name": "BLUE", "color": "#095cd9", "csv_char": "B"},
            "P4": {"name": "YELLOW", "color": "#FEEA63", "csv_char": "Y"},
        }

        # Autosave path setup
        try:
            home_dir = os.path.expanduser("~")
            app_data_dir = os.path.join(home_dir, "Desktop", "data")
            os.makedirs(app_data_dir, exist_ok=True)
            self.autosave_path = os.path.join(app_data_dir, "kx_race_analysis_git.csv")
            print(f"Saving data to: {self.autosave_path}")
        except Exception as e:
            print(f"Warning: Could not access directory path: {e}")
            self.autosave_path = "kx_race_analysis_git.csv"

        self.setup_ui()
        self._load_existing_data()

    def _load_existing_data(self):
        """Loads data from the autosave CSV, preserving any extra columns and populating athlete lists."""
        if not os.path.isfile(self.autosave_path):
            self.log_to_display(f"No existing data file found. A new one will be created at:\n{self.autosave_path}")
            return
        try:
            with open(self.autosave_path, "r", newline="", encoding='utf-8') as f:
                if os.path.getsize(self.autosave_path) == 0:
                    self.log_to_display("Autosave file is empty. Starting fresh.")
                    return

                reader = csv.DictReader(f)
                self.tagged_data = list(reader)

                # Find extra headers
                if self.tagged_data:
                    all_headers = self.tagged_data[0].keys()
                    self.extra_headers = [h for h in all_headers if h not in self.standard_headers]
                    if self.extra_headers:
                         self.log_to_display(f"Found extra columns, will preserve: {', '.join(self.extra_headers)}")
                
                # Populate athlete name lists
                for row in self.tagged_data:
                    name = row.get("Athlete Name", "").strip()
                    gender = row.get("Gender")
                    if name:
                        if gender == "M":
                            self.male_athlete_names.add(name)
                        elif gender == "W":
                            self.female_athlete_names.add(name)

                self.log_to_display(f"Loaded {len(self.tagged_data)} existing tags from {self.autosave_path}")
                self._update_athlete_name_dropdowns()

        except Exception as e:
            messagebox.showerror("Load Error", f"Could not read autosave file: {self.autosave_path}\nError: {e}")
            self.tagged_data = []
            self.extra_headers = []

    def _write_csv(self, filepath):
        """Writes the current data to a CSV file, including extra headers."""
        all_keys = set(self.standard_headers)
        for row in self.tagged_data:
            all_keys.update(row.keys())

        final_headers = self.standard_headers + sorted([h for h in all_keys if h not in self.standard_headers])

        try:
            with open(filepath, "w", newline="", encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=final_headers, restval='')
                writer.writeheader()
                writer.writerows(self.tagged_data)
            return True, None
        except IOError as e:
            return False, e

    def autosave_csv(self):
        """Rewrites the entire CSV file with the current in-memory data."""
        success, error = self._write_csv(self.autosave_path)
        if not success:
            messagebox.showerror("Autosave Error", f"Could not write to file: {self.autosave_path}\nError: {error}")

    def cleanup_csv_data(self):
        """One-off utility to clean up Final Position column in existing data."""
        if not self.tagged_data:
            messagebox.showinfo("Cleanup", "No data loaded to clean up.")
            return

        confirm = messagebox.askyesno(
            "Confirm Cleanup", 
            "This will remove 'Final Position' values from all gate rows except 'Finish', 'DNS', and 'Fault' rows to match the new format. Proceed?"
        )
        if not confirm:
            return

        count = 0
        for row in self.tagged_data:
            gate = row.get("Gate", "")
            action = row.get("Action", "")
            faults = row.get("Faults", "")
            
            # Keep rank only for terminal/result rows
            is_terminal = (gate == "Finish" or action == "DNS" or action == "FLT" or (faults and "FLT" in faults))
            
            if not is_terminal and row.get("Final Position"):
                row["Final Position"] = ""
                count += 1
        
        if count > 0:
            self.autosave_csv()
            self.log_to_display(f"--- CLEANUP COMPLETE: Updated {count} rows. ---")
            messagebox.showinfo("Cleanup Complete", f"Successfully cleaned up {count} rows. Changes saved to CSV.")
        else:
            messagebox.showinfo("Cleanup", "No rows needed cleaning (dataset already matches current format).")

    # --- UI BUILDING ---
    def setup_ui(self):
        # Header
        header_label = tk.Label(
            self.root,
            text="KAYAK CROSS TAGGER",
            font=("Space Mono", 28, "bold"),
            bg=self.bg_color,
            fg="white",
            pady=10,
        )
        header_label.pack(pady=(5, 10))

        # Race Details Frame
        details_frame = ttk.Frame(self.root, style="TFrame")
        details_frame.pack(pady=5, fill="x", padx=50)

        tk.Label(
            details_frame,
            text="YEAR",
            bg=self.bg_color,
            fg="white",
            font=("Space Mono", 14, "bold"),
        ).grid(row=0, column=0, padx=5, pady=5)
        self.year_var = tk.StringVar(self.root)
        years = [
            str(y) for y in range(datetime.now().year, datetime.now().year - 21, -1)
        ]
        self.year_var.set(years[0])
        year_menu = ttk.Combobox(
            details_frame,
            textvariable=self.year_var,
            values=years,
            state="readonly",
            width=8,
        )
        year_menu.grid(row=1, column=0, padx=5, pady=5)

        tk.Label(
            details_frame,
            text="COMPETITION",
            bg=self.bg_color,
            fg="white",
            font=("Space Mono", 14, "bold"),
        ).grid(row=0, column=1, padx=5, pady=5)
        self.comp_var = tk.StringVar(self.root)
        competitions = [
            "WC1", "WC2", "WC3", "WC4", "WC5", "WCh", "OLY", "CCh", "NCh", "WRR", "ICF",
        ]
        self.comp_var.set(competitions[0])
        comp_menu = ttk.Combobox(
            details_frame,
            textvariable=self.comp_var,
            values=competitions,
            state="readonly",
            width=10,
        )
        comp_menu.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(
            details_frame,
            text="PHASE",
            bg=self.bg_color,
            fg="white",
            font=("Space Mono", 14, "bold"),
        ).grid(row=0, column=2, padx=5, pady=5)
        self.phase_var = tk.StringVar(self.root)

        phases = (
            [f"R{i}" for i in range(1, 12)]
            + [f"Rep{i}" for i in range(1, 6)]
            + [f"H{i}" for i in range(1, 11)]
            + [f"QF{i}" for i in range(1, 5)]
            + ["SF1", "SF2", "Final"]
        )
        self.phase_var.set(phases[0])
        phase_menu = ttk.Combobox(
            details_frame,
            textvariable=self.phase_var,
            values=phases,
            state="readonly",
            width=10,
        )
        phase_menu.grid(row=1, column=2, padx=5, pady=5)
        phase_menu.bind("<<ComboboxSelected>>", self.on_phase_change)

        tk.Label(
            details_frame,
            text="GENDER",
            bg=self.bg_color,
            fg="white",
            font=("Space Mono", 14, "bold"),
        ).grid(row=0, column=3, padx=5, pady=5)
        self.gender_var = tk.StringVar(self.root, value="M")
        gender_toggle_btn = ttk.Button(
            details_frame,
            textvariable=self.gender_var,
            command=self.toggle_gender,
            style="Small.TButton",
        )
        gender_toggle_btn.grid(row=1, column=3, padx=5, pady=5, sticky="ew")

        tk.Label(
            details_frame,
            text="BIBS",
            bg=self.bg_color,
            fg="white",
            font=("Space Mono", 14, "bold"),
        ).grid(row=0, column=4, padx=5, pady=5)
        paddler_count_menu = ttk.Combobox(
            details_frame,
            textvariable=self.num_paddlers_var,
            values=[4, 3, 2],
            state="readonly",
            width=5,
        )
        paddler_count_menu.grid(row=1, column=4, padx=5, pady=5)
        paddler_count_menu.bind("<<ComboboxSelected>>", self.on_paddler_count_change)

        details_frame.columnconfigure((0, 1, 2, 3, 4), weight=1)

        paddlers_frame = ttk.Frame(self.root, style="TFrame")
        paddlers_frame.pack(pady=10, fill="x", padx=50)
        paddlers_frame.columnconfigure((0, 2), weight=1)
        paddlers_frame.columnconfigure(1, weight=2) 

        paddlers_buttons_frame = ttk.Frame(paddlers_frame, style="TFrame")
        paddlers_buttons_frame.grid(row=0, column=0, sticky="n")
        tk.Label(
            paddlers_buttons_frame,
            text="BIB",
            font=("Space Mono", 14, "bold"),
            bg=self.bg_color,
            fg="white",
        ).pack(pady=(0, 10))
        paddler_grid = ttk.Frame(paddlers_buttons_frame, style="TFrame")
        paddler_grid.pack()
        for i, (p_key, bib_info) in enumerate(self.bib_data.items()):
            color = bib_info["color"]
            name = bib_info["name"]
            lbl = tk.Label(
                paddler_grid,
                text=name,
                bg=color,
                fg="#010101",
                font=("Space Mono", 14, "bold"),
                relief=tk.RAISED,
                width=6,
                height=3,
                bd=2,
            )
            lbl.bind("<Button-1>", lambda e, n=p_key: self.on_paddler_press(n))
            lbl.grid(row=i // 2, column=i % 2, padx=5, pady=5)
            self.paddler_buttons[p_key] = lbl

        # --- MIDDLE COLUMN: Names & Actions ---
        middle_frame = ttk.Frame(paddlers_frame, style="TFrame")
        middle_frame.grid(row=0, column=1, sticky="ns", padx=20, pady=(10, 0))
        
        # Athlete Name Entry Section
        tk.Label(middle_frame, text="ATHLETE NAMES", font=("Space Mono", 14, "bold"), bg=self.bg_color, fg="white").pack(pady=(0,5))
        name_entry_frame = ttk.Frame(middle_frame, style="TFrame")
        name_entry_frame.pack(pady=5, fill="x")

        for p_key, bib_info in self.bib_data.items():
            name_frame = ttk.Frame(name_entry_frame, style="TFrame")
            name_frame.pack(fill="x", pady=2)
            
            lbl = tk.Label(name_frame, text=bib_info["name"], bg=bib_info["color"], fg="#010101", font=("Space Mono", 10, "bold"), width=6)
            lbl.pack(side="left", padx=(0, 5))

            combo = ttk.Combobox(name_frame, textvariable=self.athlete_name_vars[p_key], font=("Space Mono", 10))
            combo.pack(side="left", fill="x", expand=True)
            self.athlete_name_comboboxes[p_key] = combo

        # Action Buttons Section
        clear_frame = ttk.Frame(middle_frame, style="TFrame")
        clear_frame.pack(pady=15, fill="x")
        
        ttk.Button(clear_frame, text="CLEAR ALL", command=self.clear_all_assignments, style="Small.TButton").pack(pady=2, fill="x")
        ttk.Button(clear_frame, text="CLEAR CURRENT", command=self.clear_tag_selection, style="Small.TButton").pack(pady=2, fill="x")
        ttk.Button(clear_frame, text="UNDO LAST", command=self.undo_last_paddler, style="Small.TButton").pack(pady=2, fill="x")
        ttk.Button(clear_frame, text="DNS", command=self.save_dns_tag, style="Small.TButton").pack(pady=10, fill="x")
        ttk.Button(clear_frame, text="FAULT", command=self.save_fault_tag, style="Small.Red.TButton").pack(pady=2, fill="x")
        
        # Retroactive cleanup button
        cleanup_btn = ttk.Button(clear_frame, text="CLEAN UP CSV", command=self.cleanup_csv_data, style="Small.TButton")
        cleanup_btn.pack(pady=(20, 2), fill="x")


        position_frame = ttk.Frame(paddlers_frame, style="TFrame")
        position_frame.grid(row=0, column=2, sticky="n")
        tk.Label(
            position_frame,
            text="RAMP POSITION",
            font=("Space Mono", 14, "bold"),
            bg=self.bg_color,
            fg="white",
        ).pack(pady=(0, 10))
        position_grid = ttk.Frame(position_frame, style="TFrame")
        position_grid.pack()
        positions = [1, 2, 3, 4]
        for i, pos in enumerate(positions):
            lbl = tk.Label(
                position_grid,
                text=str(pos),
                width=6,
                height=3,
                font=("Space Mono", 14, "bold"),
                bg="#CCCCCC",
                fg="black",
                relief=tk.RAISED,
                bd=2,
            )
            lbl.bind("<Button-1>", lambda e, p=pos: self.assign_ramp_position(p))
            lbl.grid(row=i // 2, column=i % 2, padx=5, pady=5)
            self.ramp_position_buttons[pos] = lbl

        segments_frame = ttk.Frame(self.root, style="TFrame")
        segments_frame.pack(pady=5, fill="x", padx=50)
        tk.Label(
            segments_frame,
            text="GATE",
            font=("Space Mono", 14, "bold"),
            bg=self.bg_color,
            fg="white",
        ).pack(pady=(0, 10))
        segments_grid = ttk.Frame(segments_frame, style="TFrame")
        segments_grid.pack(fill="x")
        for i in range(1, 9):
            gate_name = f"Gate {i}"
            btn = ttk.Button(
                segments_grid,
                text=str(i),
                command=lambda s=gate_name: self.select_gate(s),
            )
            btn.grid(row=0, column=i - 1, padx=5, pady=5, sticky="ew")
            self.gate_buttons[gate_name] = btn
            segments_grid.columnconfigure(i - 1, weight=1)
        
        self.gate_order = [f"Gate {i}" for i in range(1, 9)]

        actions_frame = ttk.Frame(segments_frame, style="TFrame")
        actions_frame.pack(fill="x", pady=(10, 0))
        action_defs = [
            ("Roll", "ExtraSmall.TButton"),
            ("Through", "ExtraSmall.TButton"),
            ("Left", "ExtraSmall.TButton"),
            ("Right", "ExtraSmall.TButton"),
            ("Up", "ExtraSmall.Red.TButton"),
            ("Down", "ExtraSmall.Green.TButton"),
            ("Finish", "ExtraSmall.TButton"),
        ]
        for i, (name, style) in enumerate(action_defs):
            btn = ttk.Button(
                actions_frame,
                text=name,
                style=style,
                command=lambda a=name: self.select_action(a),
            )
            btn.grid(row=0, column=i, padx=2, pady=2, sticky="ew")
            self.action_buttons[name] = btn
            actions_frame.columnconfigure(i, weight=1)

        log_frame = ttk.Frame(self.root, style="TFrame")
        log_frame.pack(pady=10, fill="both", expand=True, padx=50)
        tk.Label(
            log_frame,
            text="TAGS",
            font=("Space Mono", 14, "bold"),
            bg=self.bg_color,
            fg="white",
        ).pack(pady=(0, 10))
        self.log_display = tk.Text(
            log_frame,
            height=8,
            state=tk.DISABLED,
            bg=self.bg_color,
            fg="white",
            bd=1,
            relief=tk.SOLID,
            font=("Space Mono", 10),
            insertbackground="white",
        )
        self.log_display.pack(fill="both", expand=True, pady=5)
        self.root.bind("<Key>", self.handle_keypress)

    def log_to_display(self, message):
        self.log_display.config(state=tk.NORMAL)
        self.log_display.insert(tk.END, message + "\n")
        self.log_display.see(tk.END)
        self.log_display.config(state=tk.DISABLED)

    def handle_keypress(self, event):
        if isinstance(event.widget, (ttk.Combobox, tk.Listbox)):
            return
        
        char = event.char
        keysym = event.keysym
        if char in ["1", "2", "3", "4"]:
            self.on_paddler_press(f"P{char}")
        elif keysym.startswith("F") and keysym[1:].isdigit():
            gate_num = int(keysym[1:])
            if 1 <= gate_num <= 8:
                self.select_gate(f"Gate {gate_num}")
        elif keysym == "BackSpace":
            self.save_dns_tag()
        elif char.lower() == "z":
            self.undo_last_paddler()
            
        elif char.lower() == "n":
            self.select_next_gate()
            
        elif char.lower() == "d":
            self.select_action("Down")
        elif char.lower() == "f":
            self.select_action("Finish")
        elif char.lower() == "l":
            self.select_action("Left")
        elif char.lower() == "o":
            self.select_action("Roll")
        elif char.lower() == "r":
            self.select_action("Right")
        elif char.lower() == "t":
            self.select_action("Through")
        elif char.lower() == "u":
            self.select_action("Up")
            
    def _update_athlete_name_dropdowns(self):
        """Updates the values in the athlete name comboboxes based on gender."""
        current_gender = self.gender_var.get()
        name_list = sorted(list(self.male_athlete_names)) if current_gender == "M" else sorted(list(self.female_athlete_names))
        for combo in self.athlete_name_comboboxes.values():
            combo['values'] = name_list


    def on_paddler_count_change(self, event=None):
        num_paddlers = self.num_paddlers_var.get()
        disabled_color = "#404040"

        for i in range(1, 5):
            p_key = f"P{i}"
            p_button = self.paddler_buttons[p_key]
            combo = self.athlete_name_comboboxes[p_key]

            if i <= num_paddlers:
                p_button.config(bg=self.bib_data[p_key]["color"])
                p_button.bind("<Button-1>", lambda e, n=p_key: self.on_paddler_press(n))
                combo.config(state="normal")
            else:
                p_button.config(bg=disabled_color)
                p_button.unbind("<Button-1>")
                self.athlete_name_vars[p_key].set("")
                combo.config(state="disabled")

        for i in range(1, 5):
            r_button = self.ramp_position_buttons[i]
            r_button.config(bg="#CCCCCC")
            r_button.bind("<Button-1>", lambda e, p=i: self.assign_ramp_position(p))

        self.clear_all_assignments()

    def on_phase_change(self, event=None):
        self.log_to_display(
            f"--- NEW PHASE: {self.phase_var.get()}. Clearing assignments. ---"
        )
        self.finish_line_sequence.clear()
        self.faulted_bibs.clear()
        self.phase_final_positions.clear()
        self.upstream_tactic_actions.clear()
        self.clear_all_assignments()
        for var in self.athlete_name_vars.values():
            var.set("")

    def _get_athlete_name_for_paddler(self, paddler_key):
        """Finds the athlete name for a given paddler from the initial 'Ramp' tag in the current race."""
        bib_csv_char = self.bib_data[paddler_key]["csv_char"]
        current_phase = self.phase_var.get()
        current_comp = self.comp_var.get()
        current_year = self.year_var.get()

        for row in reversed(self.tagged_data):
            if (row.get("Year") == current_year and
                row.get("Competition") == current_comp and
                row.get("Phase") == current_phase and
                row.get("BIB") == bib_csv_char and
                row.get("Athlete Name")):
                return row.get("Athlete Name")
        return "" 

    def _backfill_final_position(self, bib_key, final_pos):
        """Updates internal state tracking for a bib's final position. Does NOT backfill old CSV rows."""
        bib_csv_char = self.bib_data[bib_key]["csv_char"]
        self.phase_final_positions[bib_csv_char] = final_pos

    def _find_and_copy_extra_data(self, new_entry):
        for existing_row in reversed(self.tagged_data):
            is_match = (
                existing_row.get('Year') == new_entry.get('Year') and
                existing_row.get('Competition') == new_entry.get('Competition') and
                existing_row.get('Phase') == new_entry.get('Phase') and
                existing_row.get('BIB') == new_entry.get('BIB')
            )
            if is_match:
                for header in self.extra_headers:
                    if existing_row.get(header):
                        new_entry[header] = existing_row[header]
                return

    def on_paddler_press(self, name):
        if name in self.faulted_bibs:
            self.log_to_display(
                f"NOTE: {self.bib_data[name]['name']} has FAULTED and cannot be tagged further."
            )
            return

        num_paddlers = self.num_paddlers_var.get()
        if len(self.paddler_ramp_positions) < num_paddlers:
            self.select_paddler_for_setup(name)
        else:
            self.add_paddler_to_sequence(name)

    def select_paddler_for_setup(self, name):
        if self.selected_paddler_setup:
            if self.selected_paddler_setup not in self.paddler_ramp_positions:
                self.paddler_buttons[self.selected_paddler_setup].config(
                    relief=tk.RAISED
                )
        self.selected_paddler_setup = name
        self.paddler_buttons[name].config(relief=tk.SUNKEN)
        if self.athlete_name_comboboxes.get(name):
            self.athlete_name_comboboxes[name].focus_set()

    def assign_ramp_position(self, position):
        if not self.selected_paddler_setup:
            messagebox.showwarning(
                "No BIB Selected", "Please select a BIB before assigning a position."
            )
            return
        if position in self.disabled_positions:
            messagebox.showwarning(
                "Position Taken", f"Position {position} is already assigned."
            )
            return
        paddler_name = self.selected_paddler_setup

        if paddler_name in self.paddler_ramp_positions:
            old_pos = self.paddler_ramp_positions[paddler_name]
            self.ramp_position_buttons[old_pos].config(bg="#CCCCCC")
            self.disabled_positions.remove(old_pos)

        paddler_color = self.bib_data[paddler_name]["color"]
        self.paddler_ramp_positions[paddler_name] = position
        self.ramp_position_buttons[position].config(bg=paddler_color)
        self.ramp_position_buttons[position].unbind("<Button-1>")
        self.disabled_positions.add(position)
        self.paddler_buttons[paddler_name].config(relief=tk.RAISED)

        entry = {
            "Year": self.year_var.get(), "Competition": self.comp_var.get(),
            "Gender": self.gender_var.get(), "Phase": self.phase_var.get(),
            "Gate": "Ramp", "BIB": self.bib_data[paddler_name]["csv_char"],
            "Ramp Position": position, "Action": "Assigned", "Order": "",
            "Final Position": "", "Upstream Tactic": "",
        }
        
        athlete_name = self.athlete_name_vars[paddler_name].get().strip()
        if athlete_name:
            entry['Athlete Name'] = athlete_name
            gender = self.gender_var.get()
            if gender == "M":
                self.male_athlete_names.add(athlete_name)
            else:
                self.female_athlete_names.add(athlete_name)
            self._update_athlete_name_dropdowns()

        self.tagged_data.append(entry)
        self.autosave_csv()
        log_msg = f"--> SAVED: Ramp, {self.bib_data[paddler_name]['name']}[{entry['Ramp Position']}], {entry['Action']}"
        if athlete_name:
            log_msg += f" [{athlete_name}]"
        self.log_to_display(log_msg)
        self.selected_paddler_setup = None

    def clear_all_assignments(self):
        """Clears all ramp positions, selections, and resets faulted bibs."""
        self.paddler_ramp_positions.clear()
        self.selected_paddler_setup = None
        self.disabled_positions.clear()
        self.faulted_bibs.clear()
        self.dns_bibs.clear() 
        self.phase_final_positions.clear()
        self.finish_line_sequence.clear()
        self.upstream_tactic_actions.clear()

        for i in range(1, 5):
            p_key = f"P{i}"
            self.paddler_buttons[p_key].config(relief=tk.RAISED)
            pos_button = self.ramp_position_buttons[i]
            pos_button.config(bg="#CCCCCC")
            pos_button.bind("<Button-1>", lambda e, p=i: self.assign_ramp_position(p))

        self.on_paddler_count_change()
        self.clear_tag_selection()
        self.log_to_display("--- All ramp positions and states cleared ---")

    def select_gate(self, gate_name):
        if "Roll" in self.selected_actions:
            self.selected_actions.remove("Roll")
            self.action_buttons["Roll"].config(style="ExtraSmall.TButton")
            self.paddler_order_sequence.clear()
            self.update_paddler_sequence_ui()

        if self.selected_gate != gate_name:
            if self.selected_gate:
                self.gate_buttons[self.selected_gate].config(style="TButton")

            self.paddler_order_sequence.clear()
            self.update_paddler_sequence_ui()
            self.upstream_tactic_actions.clear()

            self.selected_gate = gate_name
            self.gate_buttons[gate_name].config(style="Active.TButton")
        else:
            self.log_to_display(f"Gate {gate_name.split(' ')[1]} is active.")

    def select_next_gate(self, event=None):
        if not self.gate_order: return
        if not self.selected_gate:
            next_gate_name = self.gate_order[0]
        else:
            try:
                current_index = self.gate_order.index(self.selected_gate)
                if current_index + 1 < len(self.gate_order):
                    next_gate_name = self.gate_order[current_index + 1]
                else: return
            except ValueError:
                next_gate_name = self.gate_order[0]
        self.select_gate(next_gate_name)

    def select_action(self, action_name):
        if action_name in ["Up", "Down"]:
            self.selected_actions.add(action_name)
            return

        current_style = self.action_buttons[action_name].cget("style")
        is_active = "Active" in current_style

        if is_active:
            self.selected_actions.remove(action_name)
            self.action_buttons[action_name].config(style=current_style.replace(".Active", ""))
            if action_name == "Roll":
                self.clear_tag_selection(keep_context=False)
        else:
            if action_name == "Roll":
                if self.selected_gate:
                    self.gate_buttons[self.selected_gate].config(style="TButton")
                    self.selected_gate = None
                for other_action in list(self.selected_actions):
                    if other_action not in ["Up", "Down"]:
                        style = self.action_buttons[other_action].cget("style").replace(".Active", "")
                        self.action_buttons[other_action].config(style=style)
                self.selected_actions.clear()
                self.paddler_order_sequence.clear()
                self.upstream_tactic_actions.clear()
                self.update_paddler_sequence_ui()
            else:
                if "Roll" in self.selected_actions:
                    self.selected_actions.remove("Roll")
                    self.action_buttons["Roll"].config(style="ExtraSmall.TButton")
                    self.paddler_order_sequence.clear()
                    self.update_paddler_sequence_ui()

            if "Up" in self.selected_actions:
                if action_name == "Left" and "Right" in self.selected_actions:
                    self.selected_actions.remove("Right")
                    style = self.action_buttons["Right"].cget("style").replace(".Active", "")
                    self.action_buttons["Right"].config(style=style)
                elif action_name == "Right" and "Left" in self.selected_actions:
                    self.selected_actions.remove("Left")
                    style = self.action_buttons["Left"].cget("style").replace(".Active", "")
                    self.action_buttons["Left"].config(style=style)

            self.selected_actions.add(action_name)
            self.action_buttons[action_name].config(style=current_style.replace(".TButton", ".Active.TButton"))

    def add_paddler_to_sequence(self, paddler_name):
        if not self.selected_actions:
            messagebox.showwarning("Action Needed", "Please select an action.")
            return

        is_roll_only = ("Roll" in self.selected_actions) and (len(self.selected_actions) == 1)
        is_finish = "Finish" in self.selected_actions
        is_upstream_gate = "Up" in self.selected_actions

        if not self.selected_gate and not is_roll_only and not is_finish:
            messagebox.showwarning("Gate Needed", "A gate must be selected.")
            return

        if paddler_name in self.paddler_order_sequence: return

        num_paddlers = self.num_paddlers_var.get()
        active_paddlers = num_paddlers - len(self.faulted_bibs)

        if len(self.paddler_order_sequence) >= active_paddlers: return

        self.paddler_order_sequence.append(paddler_name)
        order_num = len(self.paddler_order_sequence)

        action_string = "-".join(sorted(list(self.selected_actions)))
        gate_value = self.selected_gate if self.selected_gate else "Course"
        if is_finish: gate_value = "Finish"

        bib_csv_char = self.bib_data[paddler_name]["csv_char"]
        
        # --- NEW FINAL POSITION LOGIC ---
        final_pos = "" 
        if is_finish:
            if paddler_name not in self.finish_line_sequence:
                self.finish_line_sequence.append(paddler_name)
            final_pos = self.finish_line_sequence.index(paddler_name) + 1
            self._backfill_final_position(paddler_name, final_pos)

        upstream_tactic = ""
        if is_upstream_gate:
            if not self.upstream_tactic_actions:
                upstream_tactic = action_string
            else:
                previous_action = self.upstream_tactic_actions[-1]
                upstream_tactic = "FOLLOW" if action_string == previous_action else "SPLIT"
            self.upstream_tactic_actions.append(action_string)

        entry = {
            "Year": self.year_var.get(), "Competition": self.comp_var.get(),
            "Gender": self.gender_var.get(), "Phase": self.phase_var.get(),
            "Gate": gate_value, "BIB": bib_csv_char,
            "Ramp Position": self.paddler_ramp_positions.get(paddler_name, "N/A"),
            "Action": action_string, "Order": order_num,
            "Final Position": final_pos, "Upstream Tactic": upstream_tactic,
        }
        
        athlete_name = self._get_athlete_name_for_paddler(paddler_name)
        self._find_and_copy_extra_data(entry)
        self.tagged_data.append(entry)
        self.autosave_csv()

        self.update_paddler_sequence_ui()
        suffix = self.get_ordinal_suffix(order_num)
        log_msg = f"--> SAVED: {gate_value}, {self.bib_data[paddler_name]['name']} ({order_num}{suffix}), Action: {action_string}"
        if athlete_name: log_msg += f" [{athlete_name}]"
        if is_finish: log_msg += f" [Rank: {final_pos}]"
        self.log_to_display(log_msg)

        if len(self.paddler_order_sequence) >= active_paddlers:
            self.clear_tag_selection(keep_context=False)

    def undo_last_paddler(self):
        if not self.paddler_order_sequence: return
        last_paddler = self.paddler_order_sequence.pop()
        if self.tagged_data:
            self.tagged_data.pop()
            self.autosave_csv()
        if "Finish" in self.selected_actions and last_paddler in self.finish_line_sequence:
            self.finish_line_sequence.remove(last_paddler)
        if "Up" in self.selected_actions and self.upstream_tactic_actions:
            self.upstream_tactic_actions.pop()
        self.update_paddler_sequence_ui()

    def get_ordinal_suffix(self, num):
        if 10 <= num % 100 <= 20: return "th"
        return {1: "st", 2: "nd", 3: "rd"}.get(num % 10, "th")

    def update_paddler_sequence_ui(self):
        for p_key, bib_info in self.bib_data.items():
            if p_key in self.paddler_order_sequence:
                order_num = self.paddler_order_sequence.index(p_key) + 1
                self.paddler_buttons[p_key].config(text=f"{bib_info['name']}\n({order_num}{self.get_ordinal_suffix(order_num)})")
            else:
                self.paddler_buttons[p_key].config(text=bib_info["name"])

    def save_dns_tag(self):
        paddler_to_mark = self.selected_paddler_setup
        if not paddler_to_mark: return
        self.dns_bibs.add(paddler_to_mark)
        entry = {
            "Year": self.year_var.get(), "Competition": self.comp_var.get(),
            "Gender": self.gender_var.get(), "Phase": self.phase_var.get(),
            "Gate": "Start", "BIB": self.bib_data[paddler_to_mark]["csv_char"],
            "Ramp Position": self.paddler_ramp_positions.get(paddler_to_mark, "N/A"),
            "Action": "DNS", "Order": 0, "Final Position": "DNS", "Upstream Tactic": "",
        }
        entry['Athlete Name'] = self._get_athlete_name_for_paddler(paddler_to_mark)
        self.tagged_data.append(entry)
        self.autosave_csv()
        self.log_to_display(f"--> SAVED: {self.bib_data[paddler_to_mark]['name']} DNS")
        self.paddler_buttons[paddler_to_mark].config(relief=tk.RAISED)
        self.selected_paddler_setup = None

    def save_fault_tag(self):
        total = self.num_paddlers_var.get()
        expected = total - len(self.dns_bibs)
        if len(self.finish_line_sequence) < expected:
            messagebox.showwarning("Race Not Finished", f"Tag all finishers before faults.")
            return
        self.open_fault_selection_popup()

    def open_fault_selection_popup(self):
        popup = tk.Toplevel(self.root)
        popup.title("Fault Entry"); popup.config(bg=self.bg_color); popup.geometry("450x550")
        
        selected_bibs = set()
        if self.selected_paddler_setup: selected_bibs.add(self.selected_paddler_setup)

        bib_frame = ttk.Frame(popup, style="TFrame"); bib_frame.pack(fill="x", padx=20, pady=5)
        def toggle_bib(btn, k, c):
            if k in selected_bibs:
                selected_bibs.remove(k); btn.config(relief=tk.RAISED, text=self.bib_data[k]["name"])
            else:
                selected_bibs.add(k); btn.config(relief=tk.SUNKEN, text=f"✓ {self.bib_data[k]['name']}")

        for i in range(1, self.num_paddlers_var.get() + 1):
            k = f"P{i}"; c = self.bib_data[k]["color"]
            txt = f"✓ {self.bib_data[k]['name']}" if k in selected_bibs else self.bib_data[k]["name"]
            rlf = tk.SUNKEN if k in selected_bibs else tk.RAISED
            btn = tk.Label(bib_frame, text=txt, bg=c, fg="#010101", font=("Space Mono", 10, "bold"), relief=rlf, width=8, height=2, bd=3)
            btn.bind("<Button-1>", lambda e, b=btn, k=k, c=c: toggle_bib(b, k, c))
            btn.pack(side="left", fill="x", expand=True, padx=2)

        tk.Label(popup, text="SELECT FAULT LOCATION(S)", bg=self.bg_color, fg="white", font=("Space Mono", 12)).pack(pady=10)
        btn_frame = ttk.Frame(popup, style="TFrame"); btn_frame.pack(fill="both", expand=True, padx=20)
        selected_faults = set()
        def toggle_fault(btn, label):
            if label in selected_faults:
                selected_faults.remove(label); btn.config(style="ExtraSmall.TButton")
            else:
                selected_faults.add(label); btn.config(style="ExtraSmall.Red.TButton")

        options = [str(i) for i in range(1, 9)] + ["Roll", "Course"]
        for i, opt in enumerate(options):
            btn = ttk.Button(btn_frame, text=opt, style="ExtraSmall.TButton")
            btn.configure(command=lambda b=btn, o=opt: toggle_fault(b, o))
            btn.grid(row=i // 3, column=i % 3, sticky="ew", padx=5, pady=5)
            btn_frame.columnconfigure(i % 3, weight=1)

        ttk.Button(popup, text="CONFIRM", style="Small.Red.TButton", command=lambda: [self._finalize_fault_tag(list(selected_bibs), selected_faults), popup.destroy()]).pack(pady=20)

    def _finalize_fault_tag(self, bib_keys, selected_faults):
        num_paddlers = self.num_paddlers_var.get()
        current_phase = self.phase_var.get()
        for bib_key in bib_keys:
            self.faulted_bibs.add(bib_key)
            bib_csv_char = self.bib_data[bib_key]["csv_char"]
            final_pos = num_paddlers - (len(self.faulted_bibs) - 1)
            self._backfill_final_position(bib_key, final_pos)
            
            for fault_item in selected_faults:
                target_gate = f"Gate {fault_item}" if fault_item.isdigit() else fault_item
                fault_str = f"FLT {fault_item}" if fault_item.isdigit() else f"FLT R" if fault_item == "Roll" else "FLT Course"
                found_row = False
                for row in reversed(self.tagged_data):
                    if (row.get("Phase") == current_phase and row.get("BIB") == bib_csv_char):
                        if (target_gate and row.get("Gate") == target_gate) or (fault_item == "Roll" and "Roll" in row.get("Action", "")) or (fault_item == "Course" and row.get("Gate") == "Finish"):
                            row["Faults"] = (row.get("Faults", "") + ", " + fault_str).strip(", ")
                            row["Final Position"] = final_pos
                            found_row = True; break 
                if not found_row:
                    new_entry = {"Year": self.year_var.get(), "Competition": self.comp_var.get(), "Gender": self.gender_var.get(), "Phase": current_phase, "Gate": target_gate if target_gate != "Roll" else "Course", "BIB": bib_csv_char, "Ramp Position": self.paddler_ramp_positions.get(bib_key, "N/A"), "Action": "FLT", "Order": "", "Final Position": final_pos, "Faults": fault_str}
                    self._find_and_copy_extra_data(new_entry); self.tagged_data.append(new_entry)
            self.paddler_buttons[bib_key].config(bg="#808080"); self.paddler_buttons[bib_key].unbind("<Button-1>")
        self.autosave_csv(); self.selected_paddler_setup = None

    def clear_tag_selection(self, keep_context=False):
        self.paddler_order_sequence.clear(); self.update_paddler_sequence_ui()
        if self.selected_paddler_setup and not keep_context:
            if self.selected_paddler_setup not in self.faulted_bibs: self.paddler_buttons[self.selected_paddler_setup].config(relief=tk.RAISED)
            self.selected_paddler_setup = None
        if not keep_context:
            self.upstream_tactic_actions.clear()
            if self.selected_gate: self.gate_buttons[self.selected_gate].config(style="TButton")
            for a in list(self.selected_actions):
                if a not in ["Up", "Down"]: self.action_buttons[a].config(style=self.action_buttons[a].cget("style").replace(".Active", ""))
            self.selected_gate = None; self.selected_actions.clear()

    def toggle_gender(self):
        self.gender_var.set("W" if self.gender_var.get() == "M" else "M")
        self._update_athlete_name_dropdowns()


def main():
    root = tk.Tk(); app = KXTaggerApp(root); root.mainloop()


if __name__ == "__main__":
    main()
