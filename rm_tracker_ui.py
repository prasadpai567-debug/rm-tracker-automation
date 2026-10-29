from __future__ import annotations

import os
import queue
import sys
import threading
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable


APP_VERSION = "1.1.0"


def run_gui(
    *,
    generate: Callable[..., dict[str, Any]],
    read_tracker: Callable[[Path], Any],
    build_followups: Callable[..., Any],
    logger,
    script_path: Path,
) -> None:
    """Run the responsive Tkinter frontend without changing automation logic."""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    def resource_path(filename: str) -> Path:
        bundle_root = getattr(sys, "_MEIPASS", script_path.resolve().parent)
        return Path(bundle_root) / filename

    logger.info("Starting redesigned Tkinter GUI")
    root = tk.Tk()
    root.title("RM Follow-Up Automation")
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    width = min(1240, max(900, screen_width - 80))
    height = min(780, max(620, screen_height - 80))
    root.geometry(
        f"{width}x{height}+{max(0, (screen_width - width) // 2)}+"
        f"{max(0, (screen_height - height) // 2)}"
    )
    root.minsize(900, 620)
    root.option_add("*Font", ("Segoe UI", 10))

    colors = {
        "primary": "#2563EB",
        "primary_hover": "#1D4ED8",
        "primary_light": "#EFF6FF",
        "navy": "#0F172A",
        "header": "#0B1F3A",
        "bg": "#F8FAFC",
        "card": "#FFFFFF",
        "surface": "#F8FAFC",
        "border": "#E2E8F0",
        "muted": "#64748B",
        "subtle": "#94A3B8",
        "green": "#16A34A",
        "green_light": "#F0FDF4",
        "amber": "#D97706",
        "amber_light": "#FFF7ED",
        "red": "#DC2626",
    }
    root.configure(bg=colors["bg"])
    try:
        icon_path = resource_path("app_icon.ico")
        if icon_path.exists():
            root.iconbitmap(default=str(icon_path))
    except tk.TclError:
        pass

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure(
        "Primary.TButton", background=colors["primary"], foreground="#FFFFFF",
        font=("Segoe UI", 10, "bold"), padding=(16, 11), borderwidth=0,
        relief="flat", anchor="center",
    )
    style.map(
        "Primary.TButton",
        background=[("disabled", "#CBD5E1"), ("pressed", "#1E40AF"),
                    ("active", colors["primary_hover"]), ("focus", "#1D4ED8")],
        foreground=[("disabled", "#64748B")],
    )
    style.configure(
        "Secondary.TButton", background="#FFFFFF", foreground=colors["navy"],
        font=("Segoe UI", 10, "bold"), padding=(14, 10), borderwidth=1,
        relief="solid", anchor="center",
    )
    style.map(
        "Secondary.TButton",
        background=[("disabled", "#F1F5F9"), ("pressed", "#DBEAFE"),
                    ("active", colors["primary_light"]), ("focus", "#EFF6FF")],
        foreground=[("disabled", "#94A3B8")],
    )
    style.configure(
        "Link.TButton", background=colors["card"], foreground=colors["primary"],
        font=("Segoe UI", 9, "bold"), padding=(3, 3), borderwidth=0,
        relief="flat", anchor="w",
    )
    style.map(
        "Link.TButton",
        foreground=[("active", colors["primary_hover"]), ("disabled", colors["subtle"])],
    )
    style.configure(
        "App.Horizontal.TProgressbar", troughcolor="#E2E8F0",
        background=colors["primary"], borderwidth=0,
        lightcolor=colors["primary"], darkcolor=colors["primary"],
    )
    style.configure(
        "App.Vertical.TScrollbar", background="#CBD5E1", troughcolor=colors["bg"],
        borderwidth=0, arrowsize=13,
    )

    runtime_dir = (
        Path(sys.executable).resolve().parent
        if getattr(sys, "frozen", False)
        else script_path.resolve().parent
    )
    project_dir = runtime_dir
    if getattr(sys, "frozen", False):
        packaged_parent = runtime_dir.parent
        if not (runtime_dir / "Vendor Contact List.xlsx").exists() and (
            packaged_parent / "Vendor Contact List.xlsx"
        ).exists():
            project_dir = packaged_parent
    app_dir = project_dir
    default_tracker = app_dir / "RM Tracker.xlsx"
    default_contacts = app_dir / "Vendor Contact List.xlsx"
    default_output = app_dir / "RM_Tracker_Pivot_Analysis_Final.xlsx"

    tracker_var = tk.StringVar(value=str(default_tracker) if default_tracker.exists() else "")
    contacts_var = tk.StringVar(value=str(default_contacts))
    output_var = tk.StringVar(value=str(default_output))
    status_title_var = tk.StringVar(value="Ready")
    status_message_var = tk.StringVar(
        value="Select a tracker workbook, then choose an action."
    )
    details_var = tk.StringVar(value="No activity yet.")
    dashboard_date_var = tk.StringVar(value=f"As of {date.today():%d %b %Y}")
    no_confirmation_var = tk.StringVar(value="—")
    delivery_var = tk.StringVar(value="—")
    dashboard_message_var = tk.StringVar(
        value="Select a valid tracker to calculate daily follow-up totals."
    )

    events: queue.Queue = queue.Queue()
    interactive_buttons: list[ttk.Button] = []
    busy = False
    closing = False
    dashboard_request = 0
    dashboard_after_id = None

    class ToolTip:
        def __init__(self, widget, text_getter):
            self.widget = widget
            self.text_getter = text_getter
            self.window = None
            self.after_id = None
            widget.bind("<Enter>", self.schedule, add="+")
            widget.bind("<Leave>", self.hide, add="+")
            widget.bind("<ButtonPress>", self.hide, add="+")

        def schedule(self, _event=None):
            self.hide()
            self.after_id = self.widget.after(450, self.show)

        def show(self):
            value = self.text_getter()
            if not value:
                return
            self.window = tk.Toplevel(self.widget)
            self.window.wm_overrideredirect(True)
            self.window.wm_geometry(
                f"+{self.widget.winfo_rootx() + 10}+"
                f"{self.widget.winfo_rooty() + self.widget.winfo_height() + 6}"
            )
            tk.Label(
                self.window, text=value, bg=colors["navy"], fg="#FFFFFF",
                font=("Segoe UI", 9), padx=9, pady=6, relief="solid",
                borderwidth=1, wraplength=520, justify="left",
            ).pack()

        def hide(self, _event=None):
            if self.after_id is not None:
                self.widget.after_cancel(self.after_id)
                self.after_id = None
            if self.window is not None:
                self.window.destroy()
                self.window = None

    def shorten_path(value: str, limit: int = 68) -> str:
        value = value or "No file selected"
        if len(value) <= limit:
            return value
        path = Path(value)
        room = max(14, limit - len(path.name) - 5)
        return f"{str(path.parent)[:room]}...\\{path.name}"

    def choose_file(var: tk.StringVar, title: str, save: bool = False) -> None:
        current = Path(var.get()).expanduser() if var.get() else app_dir
        initial_dir = current.parent if current.suffix else current
        if not initial_dir.exists():
            initial_dir = app_dir
        common = {
            "title": title,
            "filetypes": [("Excel workbook", "*.xlsx *.xlsm")],
            "initialdir": str(initial_dir),
            "parent": root,
        }
        if save:
            selected = filedialog.asksaveasfilename(
                **common, defaultextension=".xlsx",
                initialfile=current.name if current.suffix else default_output.name,
            )
        else:
            selected = filedialog.askopenfilename(**common)
        if selected:
            var.set(selected)

    def open_output_folder() -> None:
        target = Path(output_var.get()).expanduser()
        folder = target.parent if target.suffix else target
        if not folder.exists():
            folder = app_dir
        try:
            os.startfile(str(folder))
        except OSError as exc:
            messagebox.showerror("Cannot open folder", str(exc), parent=root)

    shell = tk.Frame(root, bg=colors["bg"])
    shell.grid(row=0, column=0, sticky="nsew")
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)
    shell.grid_rowconfigure(1, weight=1)
    shell.grid_columnconfigure(0, weight=1)

    header = tk.Frame(shell, bg=colors["header"], height=84)
    header.grid(row=0, column=0, sticky="ew", padx=24, pady=(18, 0))
    header.grid_propagate(False)
    header.grid_columnconfigure(1, weight=1)

    header_logo = None
    try:
        header_logo = tk.PhotoImage(file=str(resource_path("app_logo.png")))
        factor = max(1, (max(header_logo.width(), header_logo.height()) + 53) // 54)
        if factor > 1:
            header_logo = header_logo.subsample(factor, factor)
    except tk.TclError:
        pass
    if header_logo is not None:
        logo_widget = tk.Label(header, image=header_logo, bg=colors["header"], borderwidth=0)
        logo_widget.image = header_logo
        logo_widget.grid(row=0, column=0, rowspan=2, padx=(18, 14), pady=10)
    else:
        fallback = tk.Canvas(header, width=52, height=52, bg=colors["header"], highlightthickness=0)
        fallback.create_oval(3, 3, 49, 49, outline="#60A5FA", width=2)
        fallback.create_text(26, 26, text="RM", fill="#FFFFFF",
                             font=("Segoe UI", 13, "bold"))
        fallback.grid(row=0, column=0, rowspan=2, padx=(18, 14), pady=14)

    brand = tk.Frame(header, bg=colors["header"])
    brand.grid(row=0, column=1, rowspan=2, sticky="w", pady=13)
    tk.Label(
        brand, text="RM Follow-Up Automation", bg=colors["header"], fg="#FFFFFF",
        font=("Segoe UI", 17, "bold"), anchor="w",
    ).pack(anchor="w")
    tk.Label(
        brand, text="PO automation · Track, follow up, and close",
        bg=colors["header"], fg="#BFDBFE", font=("Segoe UI", 9), anchor="w",
    ).pack(anchor="w", pady=(2, 0))
    header_meta = tk.Frame(header, bg=colors["header"])
    header_meta.grid(row=0, column=2, rowspan=2, sticky="e", padx=18)
    tk.Label(
        header_meta, text="●  Runs locally", bg=colors["header"], fg="#BBF7D0",
        font=("Segoe UI", 9, "bold"),
    ).pack(anchor="e")
    tk.Label(
        header_meta, text=f"Version {APP_VERSION}", bg=colors["header"], fg="#94A3B8",
        font=("Segoe UI", 8),
    ).pack(anchor="e", pady=(3, 0))

    viewport = tk.Frame(shell, bg=colors["bg"])
    viewport.grid(row=1, column=0, sticky="nsew")
    viewport.grid_rowconfigure(0, weight=1)
    viewport.grid_columnconfigure(0, weight=1)
    canvas = tk.Canvas(viewport, bg=colors["bg"], highlightthickness=0, borderwidth=0)
    scrollbar = ttk.Scrollbar(
        viewport, orient="vertical", command=canvas.yview, style="App.Vertical.TScrollbar"
    )
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.grid(row=0, column=0, sticky="nsew")
    scrollbar.grid(row=0, column=1, sticky="ns", pady=14)
    body = tk.Frame(canvas, bg=colors["bg"])
    body_window = canvas.create_window((0, 0), window=body, anchor="nw")
    canvas.bind("<Configure>", lambda event: canvas.itemconfigure(body_window, width=event.width))

    def scroll_workspace(event):
        if body.winfo_reqheight() > canvas.winfo_height():
            canvas.yview_scroll(int(-event.delta / 120), "units")

    root.bind_all("<MouseWheel>", scroll_workspace, add="+")

    dashboard_panel = tk.Frame(body, bg=colors["bg"])
    setup_card = tk.Frame(
        body, bg=colors["card"], highlightbackground=colors["border"], highlightthickness=1
    )

    tk.Label(
        dashboard_panel, text="Daily Follow-Up Dashboard", bg=colors["bg"],
        fg=colors["navy"], font=("Segoe UI", 22, "bold"), anchor="w",
    ).pack(fill="x")
    tk.Label(
        dashboard_panel, textvariable=dashboard_date_var, bg=colors["bg"],
        fg=colors["muted"], font=("Segoe UI", 10), anchor="w",
    ).pack(fill="x", pady=(5, 20))

    dashboard_cards = tk.Frame(dashboard_panel, bg=colors["bg"])
    dashboard_cards.pack(fill="x")
    dashboard_cards.grid_columnconfigure(0, weight=1, uniform="dashboard")
    dashboard_cards.grid_columnconfigure(1, weight=1, uniform="dashboard")

    def dashboard_card(parent, column, title, subtitle, value_var, accent, accent_light):
        card = tk.Frame(
            parent, bg=colors["card"], highlightbackground=colors["border"],
            highlightthickness=1,
        )
        card.grid(
            row=0, column=column, sticky="nsew",
            padx=(0 if column == 0 else 7, 7 if column == 0 else 0),
        )
        tk.Frame(card, bg=accent, height=5).pack(fill="x")
        title_row = tk.Frame(card, bg=colors["card"])
        title_row.pack(fill="x", padx=17, pady=(16, 0))
        tk.Label(
            title_row, text="●", bg=accent_light, fg=accent,
            font=("Segoe UI", 9, "bold"), padx=5, pady=2,
        ).pack(side="left", padx=(0, 8))
        tk.Label(
            title_row, text=title, bg=colors["card"], fg=colors["navy"],
            font=("Segoe UI", 10, "bold"), anchor="w",
        ).pack(side="left", fill="x", expand=True)
        tk.Label(
            card, textvariable=value_var, bg=colors["card"], fg=colors["navy"],
            font=("Segoe UI", 29, "bold"), anchor="w",
        ).pack(fill="x", padx=17, pady=(12, 0))
        tk.Label(
            card, text=subtitle, bg=colors["card"], fg=colors["muted"],
            font=("Segoe UI", 8), anchor="w",
        ).pack(fill="x", padx=17, pady=(1, 16))

    dashboard_card(
        dashboard_cards, 0, "No confirmation", "Follow-up items",
        no_confirmation_var, colors["amber"], colors["amber_light"],
    )
    dashboard_card(
        dashboard_cards, 1, "Delivery", "Follow-up items",
        delivery_var, colors["primary"], colors["primary_light"],
    )
    dashboard_message = tk.Label(
        dashboard_panel, textvariable=dashboard_message_var, bg=colors["bg"],
        fg=colors["muted"], font=("Segoe UI", 9), justify="left",
        anchor="w", wraplength=470,
    )
    dashboard_message.pack(fill="x", pady=(12, 18))

    workflow_card = tk.Frame(
        dashboard_panel, bg=colors["card"], highlightbackground=colors["border"],
        highlightthickness=1,
    )
    workflow_card.pack(fill="x", pady=(0, 4))
    tk.Label(
        workflow_card, text="Quick workflow", bg=colors["card"], fg=colors["navy"],
        font=("Segoe UI", 11, "bold"), anchor="w",
    ).pack(fill="x", padx=18, pady=(16, 10))
    for number, description in (
        ("1", "Choose the current RM tracker workbook."),
        ("2", "Generate the consolidated workbook and analysis."),
        ("3", "Optionally create drafts in Classic Outlook."),
    ):
        workflow_row = tk.Frame(workflow_card, bg=colors["card"])
        workflow_row.pack(fill="x", padx=18, pady=(0, 10))
        tk.Label(
            workflow_row, text=number, width=2, bg=colors["primary_light"],
            fg=colors["primary"], font=("Segoe UI", 9, "bold"), padx=2, pady=2,
        ).pack(side="left", padx=(0, 9))
        tk.Label(
            workflow_row, text=description, bg=colors["card"], fg=colors["muted"],
            font=("Segoe UI", 9), anchor="w", justify="left",
        ).pack(side="left", fill="x", expand=True)

    setup_inner = tk.Frame(setup_card, bg=colors["card"])
    setup_inner.pack(fill="both", expand=True, padx=24, pady=22)
    tk.Label(
        setup_inner, text="Workbook Setup", bg=colors["card"], fg=colors["navy"],
        font=("Segoe UI", 18, "bold"), anchor="w",
    ).pack(fill="x")
    tk.Label(
        setup_inner,
        text="Choose the tracker and output location, then start the automation.",
        bg=colors["card"], fg=colors["muted"], font=("Segoe UI", 9),
        anchor="w", justify="left",
    ).pack(fill="x", pady=(5, 18))

    def file_selector(parent, marker, title, description, var, button_text, command):
        box = tk.Frame(
            parent, bg="#FBFDFF", highlightbackground=colors["border"],
            highlightthickness=1,
        )
        box.pack(fill="x", pady=(0, 11))
        tk.Label(
            box, text=marker, width=4, bg=colors["primary_light"],
            fg=colors["primary"], font=("Segoe UI", 8, "bold"), padx=3, pady=5,
        ).pack(side="left", padx=(13, 11), pady=14)
        text_area = tk.Frame(box, bg="#FBFDFF")
        text_area.pack(side="left", fill="x", expand=True, pady=11)
        tk.Label(
            text_area, text=title, bg="#FBFDFF", fg=colors["navy"],
            font=("Segoe UI", 10, "bold"), anchor="w",
        ).pack(fill="x")
        path_label = tk.Label(
            text_area, text=shorten_path(var.get()), bg="#FBFDFF", fg=colors["muted"],
            font=("Segoe UI", 9), anchor="w", justify="left",
        )
        path_label.pack(fill="x", pady=(3, 1))
        tk.Label(
            text_area, text=description, bg="#FBFDFF", fg=colors["subtle"],
            font=("Segoe UI", 8), anchor="w",
        ).pack(fill="x")
        button = ttk.Button(
            box, text=button_text, style="Secondary.TButton", command=command, takefocus=True
        )
        button.pack(side="right", padx=12)
        interactive_buttons.append(button)
        var.trace_add(
            "write", lambda *_: path_label.configure(text=shorten_path(var.get()))
        )
        box.bind(
            "<Configure>",
            lambda event: path_label.configure(wraplength=max(210, event.width - 175)),
            add="+",
        )
        ToolTip(path_label, var.get)
        return button

    browse_button = file_selector(
        setup_inner, "IN", "Source tracker", "Excel workbook · .xlsx or .xlsm",
        tracker_var, "Browse…",
        lambda: choose_file(tracker_var, "Choose tracker workbook"),
    )
    file_selector(
        setup_inner, "OUT", "Output workbook", "Generated report location",
        output_var, "Change…",
        lambda: choose_file(output_var, "Choose output workbook", True),
    )

    contacts_path = Path(contacts_var.get())
    contact_tone = colors["green"] if contacts_path.exists() else colors["red"]
    contact_state = "Connected" if contacts_path.exists() else "File not found"
    contacts_row = tk.Frame(setup_inner, bg=colors["card"])
    contacts_row.pack(fill="x", pady=(0, 13))
    tk.Label(
        contacts_row, text="●", bg=colors["card"], fg=contact_tone,
        font=("Segoe UI", 9, "bold"),
    ).pack(side="left")
    tk.Label(
        contacts_row, text=f"Vendor contacts: {contact_state}", bg=colors["card"],
        fg=colors["muted"], font=("Segoe UI", 8),
    ).pack(side="left", padx=(5, 0))

    action_row = tk.Frame(setup_inner, bg=colors["card"])
    action_row.pack(fill="x", pady=(0, 9))
    action_row.grid_columnconfigure(0, weight=1, uniform="actions")
    action_row.grid_columnconfigure(1, weight=1, uniform="actions")
    generate_button = ttk.Button(
        action_row, text="Generate Workbook", style="Primary.TButton", takefocus=True
    )
    generate_button.grid(row=0, column=0, sticky="ew", padx=(0, 5))
    drafts_button = ttk.Button(
        action_row, text="Generate + Create Outlook Drafts",
        style="Secondary.TButton", takefocus=True,
    )
    drafts_button.grid(row=0, column=1, sticky="ew", padx=(5, 0))
    interactive_buttons.extend([generate_button, drafts_button])

    open_folder_button = ttk.Button(
        setup_inner, text="Open output folder  ↗", style="Link.TButton",
        command=open_output_folder, takefocus=True,
    )
    open_folder_button.pack(anchor="w", pady=(1, 13))
    interactive_buttons.append(open_folder_button)

    status_panel = tk.Frame(
        setup_inner, bg=colors["surface"], highlightbackground=colors["border"],
        highlightthickness=1,
    )
    status_panel.pack(fill="x")
    status_dot = tk.Canvas(
        status_panel, width=14, height=14, bg=colors["surface"], highlightthickness=0
    )
    status_indicator = status_dot.create_oval(
        2, 2, 12, 12, fill=colors["primary"], outline=""
    )
    status_dot.pack(side="left", padx=(15, 9), pady=16, anchor="n")
    status_text = tk.Frame(status_panel, bg=colors["surface"])
    status_text.pack(side="left", fill="x", expand=True, padx=(0, 14), pady=11)
    tk.Label(
        status_text, textvariable=status_title_var, bg=colors["surface"],
        fg=colors["navy"], font=("Segoe UI", 10, "bold"), anchor="w",
    ).pack(fill="x")
    status_message_label = tk.Label(
        status_text, textvariable=status_message_var, bg=colors["surface"],
        fg=colors["muted"], font=("Segoe UI", 9), anchor="w",
        justify="left", wraplength=460,
    )
    status_message_label.pack(fill="x", pady=(2, 0))

    progress = ttk.Progressbar(
        setup_inner, style="App.Horizontal.TProgressbar", mode="indeterminate"
    )
    progress.pack(fill="x", pady=(10, 0))
    progress.pack_forget()
    details_button = ttk.Button(
        setup_inner, text="Show activity details  ▾", style="Link.TButton", takefocus=True
    )
    details_button.pack(anchor="w", pady=(11, 0))
    details_container = tk.Frame(
        setup_inner, bg="#FBFDFF", highlightbackground=colors["border"],
        highlightthickness=1,
    )
    details_label = tk.Label(
        details_container, textvariable=details_var, bg="#FBFDFF", fg=colors["muted"],
        font=("Segoe UI", 8), justify="left", anchor="w", wraplength=520,
        padx=12, pady=10,
    )
    details_label.pack(fill="x")

    def set_status(title: str, message: str, tone: str = "ready") -> None:
        status_title_var.set(title)
        status_message_var.set(message)
        status_dot.itemconfigure(
            status_indicator,
            fill={
                "ready": colors["primary"], "success": colors["green"],
                "warning": colors["amber"], "error": colors["red"],
                "processing": colors["primary"],
            }.get(tone, colors["primary"]),
        )

    def toggle_details() -> None:
        if details_container.winfo_ismapped():
            details_container.pack_forget()
            details_button.configure(text="Show activity details  ▾")
        else:
            details_container.pack(fill="x", pady=(7, 0))
            details_button.configure(text="Hide activity details  ▴")

    details_button.configure(command=toggle_details)

    def set_busy(value: bool) -> None:
        nonlocal busy
        busy = value
        for button in interactive_buttons:
            button.state(["disabled"] if value else ["!disabled"])
        if value:
            progress.pack(fill="x", pady=(10, 0), before=details_button)
            progress.start(12)
            root.configure(cursor="wait")
        else:
            progress.stop()
            progress.pack_forget()
            root.configure(cursor="")

    def queue_dashboard_refresh(*_args) -> None:
        nonlocal dashboard_request, dashboard_after_id
        dashboard_request += 1
        request_id = dashboard_request
        dashboard_date_var.set(f"As of {date.today():%d %b %Y}")
        if dashboard_after_id is not None:
            root.after_cancel(dashboard_after_id)
        dashboard_after_id = root.after(
            250, lambda: start_dashboard_refresh(request_id)
        )

    def start_dashboard_refresh(request_id: int) -> None:
        nonlocal dashboard_after_id
        dashboard_after_id = None
        tracker_path = Path(tracker_var.get()).expanduser() if tracker_var.get() else None
        if not tracker_path or not tracker_path.exists():
            no_confirmation_var.set("—")
            delivery_var.set("—")
            dashboard_message_var.set(
                "Select a valid tracker to calculate daily follow-up totals."
            )
            return
        no_confirmation_var.set("…")
        delivery_var.set("…")
        dashboard_message_var.set("Calculating follow-up totals…")

        def calculate_dashboard() -> None:
            try:
                combined = read_tracker(tracker_path)
                no_total = len(build_followups(combined, {}, "no_confirmation"))
                delivery_total = len(build_followups(combined, {}, "delivery"))
                events.put(("dashboard_success", request_id, no_total, delivery_total))
            except Exception as exc:
                events.put(("dashboard_error", request_id, str(exc)))

        threading.Thread(
            target=calculate_dashboard, name="rm-dashboard", daemon=True
        ).start()

    def start_generation(
        tracker_path: Path,
        vendor_contacts: Path,
        output_path: Path,
        with_drafts: bool,
    ) -> None:
        action = (
            "Creating the workbook and Outlook drafts"
            if with_drafts
            else "Creating the workbook"
        )
        set_status(
            "Processing",
            f"{action}. You can keep this window open while the work completes.",
            "processing",
        )
        details_var.set(f"Source: {tracker_path}\nOutput: {output_path}")
        set_busy(True)

        def generate_in_background() -> None:
            try:
                result = generate(
                    tracker_path, vendor_contacts, output_path, with_drafts
                )
                events.put(("generation_success", result))
            except PermissionError as exc:
                events.put(
                    (
                        "generation_permission_error",
                        tracker_path,
                        vendor_contacts,
                        output_path,
                        with_drafts,
                        str(exc),
                    )
                )
            except Exception as exc:
                logger.exception("Generation failed")
                events.put(("generation_error", str(exc)))

        threading.Thread(
            target=generate_in_background, name="rm-generation", daemon=True
        ).start()

    def validate_and_start(with_drafts: bool) -> None:
        if busy:
            return
        tracker_path = Path(tracker_var.get()).expanduser() if tracker_var.get() else None
        vendor_contacts = (
            Path(contacts_var.get()).expanduser() if contacts_var.get() else None
        )
        if not tracker_path or not vendor_contacts:
            set_status(
                "Files required",
                "Choose both the tracker and vendor contact workbooks.",
                "warning",
            )
            messagebox.showwarning(
                "Missing files", "Choose both workbooks first.", parent=root
            )
            return
        if (
            not tracker_path.exists()
            or tracker_path.suffix.lower() not in {".xlsx", ".xlsm"}
        ):
            set_status(
                "Invalid tracker",
                "Select an existing .xlsx or .xlsm tracker workbook.",
                "error",
            )
            messagebox.showerror(
                "Invalid tracker",
                "The tracker file does not exist or is not an Excel workbook.",
                parent=root,
            )
            return
        if (
            not vendor_contacts.exists()
            or vendor_contacts.suffix.lower() not in {".xlsx", ".xlsm"}
        ):
            set_status(
                "Invalid contact list",
                "Select an existing .xlsx or .xlsm contact workbook.",
                "error",
            )
            messagebox.showerror(
                "Invalid contact list",
                "The vendor contact file does not exist or is not an Excel workbook.",
                parent=root,
            )
            return

        output_text = output_var.get().strip()
        output_path = Path(output_text).expanduser() if output_text else default_output
        if output_path.suffix.lower() not in {".xlsx", ".xlsm"}:
            output_path = output_path.with_suffix(".xlsx")
            output_var.set(str(output_path))
        if output_path.exists():
            overwrite = messagebox.askyesno(
                "Output workbook already exists",
                f"{output_path.name} already exists.\n\n"
                "Choose Yes to replace it. Close the workbook in Excel first.\n"
                "Choose No to create a new timestamped workbook.",
                parent=root,
            )
            if not overwrite:
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = output_path.with_name(
                    f"{output_path.stem}_{stamp}{output_path.suffix}"
                )
                output_var.set(str(output_path))
        start_generation(tracker_path, vendor_contacts, output_path, with_drafts)

    def finish_generation(result: dict[str, Any]) -> None:
        set_busy(False)
        set_status(
            "Completed",
            f"{result['no_confirmation']} no-confirmation and "
            f"{result['delivery']} delivery follow-ups prepared.",
            "success",
        )
        details_var.set(
            f"Output: {result['output']}\n"
            f"Outlook drafts created: {result['drafts_created']}  ·  "
            f"Missing contacts: {result['drafts_skipped']}"
        )
        messagebox.showinfo(
            "Completed",
            f"Workbook created:\n{result['output']}\n\n"
            f"Outlook drafts created: {result['drafts_created']}\n"
            f"Missing contacts: {result['drafts_skipped']}",
            parent=root,
        )

    def handle_permission_error(
        tracker_path: Path,
        vendor_contacts: Path,
        output_path: Path,
        with_drafts: bool,
        error_text: str,
    ) -> None:
        set_busy(False)
        set_status(
            "Output unavailable",
            "The selected workbook is locked or unavailable.",
            "warning",
        )
        details_var.set(error_text)
        alternative = output_path.with_name(
            f"{output_path.stem}_{datetime.now():%Y%m%d_%H%M%S}{output_path.suffix}"
        )
        use_alternative = messagebox.askyesno(
            "Cannot write output workbook",
            "Windows denied access to the selected workbook.\n\n"
            "It may be open in Excel, Preview, or another process.\n\n"
            f"Create a new file instead?\n{alternative.name}",
            parent=root,
        )
        if use_alternative:
            output_var.set(str(alternative))
            start_generation(
                tracker_path, vendor_contacts, alternative, with_drafts
            )

    def poll_events() -> None:
        if closing:
            return
        try:
            while True:
                event = events.get_nowait()
                event_name = event[0]
                if event_name == "dashboard_success":
                    _, request_id, no_total, delivery_total = event
                    if request_id == dashboard_request:
                        no_confirmation_var.set(f"{no_total:,}")
                        delivery_var.set(f"{delivery_total:,}")
                        dashboard_message_var.set(
                            "Counts include one row per consolidated PO line item."
                        )
                elif event_name == "dashboard_error":
                    _, request_id, error_text = event
                    if request_id == dashboard_request:
                        no_confirmation_var.set("—")
                        delivery_var.set("—")
                        dashboard_message_var.set(
                            f"Tracker could not be read: {error_text}"
                        )
                elif event_name == "generation_success":
                    finish_generation(event[1])
                elif event_name == "generation_permission_error":
                    handle_permission_error(*event[1:])
                elif event_name == "generation_error":
                    set_busy(False)
                    set_status("Processing failed", event[1], "error")
                    details_var.set(event[1])
                    messagebox.showerror(
                        "RM Tracker Automation", event[1], parent=root
                    )
        except queue.Empty:
            pass
        root.after(100, poll_events)

    generate_button.configure(command=lambda: validate_and_start(False))
    drafts_button.configure(command=lambda: validate_and_start(True))

    layout_mode = None

    def arrange_panels(available_width: int) -> None:
        nonlocal layout_mode
        target = "stacked" if available_width < 1160 else "wide"
        if target == layout_mode:
            return
        layout_mode = target
        dashboard_panel.grid_forget()
        setup_card.grid_forget()
        if target == "stacked":
            body.grid_columnconfigure(0, weight=1)
            body.grid_columnconfigure(1, weight=0)
            dashboard_panel.grid(
                row=0, column=0, sticky="ew", padx=24, pady=(24, 13)
            )
            setup_card.grid(
                row=1, column=0, sticky="ew", padx=24, pady=(0, 24)
            )
        else:
            body.grid_columnconfigure(0, weight=5, uniform="workspace")
            body.grid_columnconfigure(1, weight=6, uniform="workspace")
            dashboard_panel.grid(
                row=0, column=0, sticky="nsew", padx=(24, 11), pady=24
            )
            setup_card.grid(
                row=0, column=1, sticky="nsew", padx=(11, 24), pady=24
            )

    def update_workspace(event) -> None:
        canvas.configure(scrollregion=canvas.bbox("all"))
        arrange_panels(event.width)
        dashboard_message.configure(
            wraplength=max(300, dashboard_panel.winfo_width() - 16)
        )
        status_message_label.configure(
            wraplength=max(280, setup_card.winfo_width() - 110)
        )
        details_label.configure(
            wraplength=max(300, setup_card.winfo_width() - 90)
        )

    body.bind("<Configure>", update_workspace, add="+")

    footer = tk.Frame(
        shell, bg=colors["card"], highlightbackground=colors["border"],
        highlightthickness=1,
    )
    footer.grid(row=2, column=0, sticky="ew")
    tk.Label(
        footer, text="Created by Prasad Pai", bg=colors["card"], fg=colors["muted"],
        font=("Segoe UI", 8),
    ).pack(side="left", padx=24, pady=8)
    tk.Label(
        footer, text="Ctrl+O Browse  ·  Ctrl+G Generate", bg=colors["card"],
        fg=colors["subtle"], font=("Segoe UI", 8),
    ).pack(side="right", padx=24, pady=8)

    def keyboard_browse(_event=None):
        if not busy:
            choose_file(tracker_var, "Choose tracker workbook")
        return "break"

    def keyboard_generate(_event=None):
        validate_and_start(False)
        return "break"

    def keyboard_generate_drafts(_event=None):
        validate_and_start(True)
        return "break"

    root.bind("<Control-o>", keyboard_browse)
    root.bind("<Control-g>", keyboard_generate)
    root.bind("<Control-Shift-G>", keyboard_generate_drafts)

    def on_close() -> None:
        nonlocal closing
        if busy and not messagebox.askyesno(
            "Processing is still running",
            "The current operation has not finished. Close the application anyway?",
            parent=root,
        ):
            return
        closing = True
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    tracker_var.trace_add("write", queue_dashboard_refresh)
    queue_dashboard_refresh()
    root.after(0, lambda: arrange_panels(canvas.winfo_width()))
    root.after(100, poll_events)
    browse_button.focus_set()
    logger.info("Redesigned Tkinter GUI initialized; entering mainloop")
    root.mainloop()
