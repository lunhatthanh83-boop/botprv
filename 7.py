import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import queue
import concurrent.futures
import os
import re
import webbrowser
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

try:
    from loader import OutlookChecker
except ImportError:
    from hotmail import OutlookChecker

class RoundedScrollbar(tk.Canvas):
    def __init__(self, parent, command=None, bg="#111111", thumb="#3A3A3A", width=12, radius=6):
        super().__init__(parent, width=width, bg=bg, highlightthickness=0, bd=0)
        self.command = command
        self.thumb_color = thumb
        self.radius = radius
        self.first = 0
        self.last = 1
        self.thumb = None
        self.bind("<Button-1>", self.on_click)
        self.bind("<B1-Motion>", self.on_drag)
        self.bind("<Configure>", lambda e: self.redraw())
    def set(self, first, last):
        self.first = float(first)
        self.last = float(last)
        self.redraw()
    def yview(self, *args):
        if self.command:
            self.command(*args)
    def on_click(self, event):
        if self.thumb is None:
            return
        x1, y1, x2, y2 = self.coords(self.thumb)
        if not (y1 <= event.y <= y2):
            h = self.winfo_height()
            t = max(min(event.y / h, 1), 0)
            self.yview("moveto", t)
    def on_drag(self, event):
        h = self.winfo_height()
        thumb_h = max(int((self.last - self.first) * h), 20)
        top = max(min(event.y - thumb_h // 2, h - thumb_h), 0)
        t = 0 if h == thumb_h else top / (h - thumb_h)
        self.yview("moveto", t)
    def draw_round_rect(self, x1, y1, x2, y2, r, fill):
        points = [x1+r, y1, x2-r, y1, x2, y1, x2, y1+r, x2, y2-r, x2, y2, x2-r, y2, x1+r, y2, x1, y2, x1, y2-r, x1, y1+r, x1, y1]
        return self.create_polygon(points, smooth=True, fill=fill, outline=fill)
    def redraw(self):
        self.delete("all")
        h = self.winfo_height()
        w = self.winfo_width()
        self.draw_round_rect(2, 2, w-2, h-2, self.radius, "#1A1A1A")
        thumb_h = max(int((self.last - self.first) * h), 28)
        top = int(self.first * (h - thumb_h))
        shadow = self.draw_round_rect(2, top+3, w-2, top+thumb_h+3, self.radius, "#2A2A2A")
        self.itemconfigure(shadow, stipple="gray25")
        self.thumb = self.draw_round_rect(2, top, w-2, top+thumb_h, self.radius, self.thumb_color)

class RoundedTreeview(ttk.Treeview):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.style = ttk.Style()
        self.configure_style()
    def configure_style(self):
        self.style.configure(
            "Custom.Treeview",
            background="#1A1A1A",
            fieldbackground="#1A1A1A",
            rowheight=24,
            bordercolor="#1A1A1A",
            borderwidth=0,
            relief="flat",
            font=("Segoe UI", 10)
        )
        self.style.map("Custom.Treeview", background=[("selected", "#00FF7F")], foreground=[("selected", "#021109")])
        self.style.layout("Custom.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])

class ScrollableFrame(tk.Frame):
    def __init__(self, parent, bg="#111111"):
        super().__init__(parent, bg=bg, highlightthickness=0, bd=0)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self.vsb = RoundedScrollbar(self, command=self.on_scroll, bg=bg, thumb="#3A3A3A", width=12, radius=6)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")
        self.inner = tk.Frame(self.canvas, bg=bg, highlightthickness=0, bd=0)
        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.window_id, width=e.width))
    def on_scroll(self, *args):
        self.canvas.yview(*args)

class HotmailCloudAouto:
    def __init__(self, root):
        self.root = root
        self.root.title("Hotmail Cloud - Auto Checker")
        icon_img = tk.PhotoImage(file="logo.png")
        self.root.iconphoto(True, icon_img)
        self.icon_img = icon_img
        self.root.configure(bg="#050505")
        self.root.geometry("1300x700")
        self.root.minsize(1000, 600)
        self.hotmail_file_path = None
        self.keyword_file_path = None
        self.total_lines = 0
        self.checked_count = 0
        self.good_count = 0
        self.found_count = 0
        self.bad_count = 0
        self.locked_count = 0
        self.excluded_count = 0
        self.error_count = 0
        self.host_not_found_count = 0
        self.multipassword_count = 0
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_percent_var = tk.StringVar(value="0%")
        self.total_lines_var = tk.StringVar(value="0")
        self.checked_var = tk.StringVar(value="0")
        self.threads_var = tk.IntVar(value=10)
        self.good_var = tk.StringVar(value="0")
        self.found_var = tk.StringVar(value="0")
        self.bad_var = tk.StringVar(value="0")
        self.locked_var = tk.StringVar(value="0")
        self.excluded_var = tk.StringVar(value="0")
        self.errors_var = tk.StringVar(value="0")
        self.host_not_found_var = tk.StringVar(value="0")
        self.multipassword_var = tk.StringVar(value="0")
        self.show_password_var = tk.BooleanVar(value=False)
        self.max_retry_var = tk.IntVar(value=3)
        self.database_label_var = tk.StringVar(value="No hotmail file selected")
        self.keyword_label_var = tk.StringVar(value="No keyword file selected")
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.on_search_change)
        self.result_queue = queue.Queue()
        self.executor = None
        self.running = False
        self.paused = False
        self.row_data = {}
        self._hover_iid = None
        self.remaining_accounts = []
        self.current_accounts = []
        self.current_active_tab = "global"
        self.all_tree_items = []
        self.pending_results = {}
        self.next_display_index = 1
        self.idx_to_item = {}
        self.item_to_idx = {}
        self.status_icons = {
            "good": "G",
            "found": "F",
            "bad": "B",
            "locked": "L",
            "excluded": "E",
            "error": "E",
            "host": "H",
            "multipassword": "M",
            "info": "I",
            "neutral": "•"
        }
        self.status_colors = {
            "good": "#00E676",
            "found": "#00E5FF",
            "bad": "#FF5252",
            "locked": "#FFB74D",
            "excluded": "#FFD54F",
            "error": "#FF4081",
            "host": "#40C4FF",
            "multipassword": "#80D8FF",
            "info": "#00B0FF",
            "neutral": "#E6E6E6"
        }
        self.status_bg_colors = {
            "good": "#00351F",
            "found": "#00323A",
            "bad": "#3A0000",
            "locked": "#3A2400",
            "excluded": "#3A2E00",
            "error": "#3A0015",
            "host": "#00293A",
            "multipassword": "#00263A",
            "info": "#002538",
            "neutral": "#1A1A1A"
        }
        self.masked_password_length = 12
        self.results_data = {"good": [], "found": [], "bad": [], "locked": [], "error": []}
        self.current_status_filter = "all"
        self.build_style()
        self.build_layout()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(50, self.process_queue)

    def build_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background="#050505")
        style.configure("Sidebar.TFrame", background="#121212")
        style.configure("Center.TFrame", background="#121212")
        style.configure("Right.TFrame", background="#121212")
        style.configure("Neon.TLabel", background="#050505", foreground="#39FF14", font=("Segoe UI", 22, "bold"))
        style.configure("SidebarLabel.TLabel", background="#141414", foreground="#E6E6E6", font=("Segoe UI", 10))
        style.configure("Glass.TFrame", background="#141414", borderwidth=0)
        style.configure("TNotebook", background="#141414", tabmargins=[2, 2, 2, 0])
        style.configure("TNotebook.Tab", background="#181818", foreground="#E6E6E6", padding=[12, 6], font=("Segoe UI", 10))
        style.map("TNotebook.Tab", background=[("selected", "#202020")], foreground=[("selected", "#39FF14")])
        style.configure("SettingsLabel.TLabel", background="#141414", foreground="#E6E6E6", font=("Segoe UI", 10))
        style.configure("SettingsValue.TLabel", background="#141414", foreground="#9A9A9A", font=("Segoe UI", 9))
        style.configure(
            "Custom.Treeview",
            background="#1A1A1A",
            foreground="#E6E6E6",
            fieldbackground="#1A1A1A",
            rowheight=40,
            bordercolor="#1A1A1A",
            borderwidth=0,
            relief="flat",
            font=("Segoe UI", 10)
        )
        style.map("Custom.Treeview", background=[("selected", "#00FF7F")], foreground=[("selected", "#021109")])
        style.layout("Custom.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        style.configure("Custom.Treeview.Heading", background="#141414", foreground="#E6E6E6", font=("Segoe UI", 10, "bold"), relief="flat", borderwidth=0)
        style.map("Custom.Treeview.Heading", background=[("active", "#141414")])
        style.configure(
            "Modern.Horizontal.TProgressbar",
            troughcolor="#0A0A0A",
            bordercolor="#0A0A0A",
            background="#00FF7F",
            lightcolor="#00FF7F",
            darkcolor="#00CC66",
            thickness=10
        )

    def draw_round_rect(self, canvas, x1, y1, x2, y2, r, **kwargs):
        points = [x1+r, y1, x2-r, y1, x2, y1, x2, y1+r, x2, y2-r, x2, y2, x2-r, y2, x1+r, y2, x1, y2, x1, y2-r, x1, y1+r, x1, y1]
        return canvas.create_polygon(points, smooth=True, **kwargs)

    def create_rounded_frame(self, parent, bg_color, radius=12, **kwargs):
        outer = tk.Frame(parent, bg=parent.cget('bg'), **kwargs)
        canvas = tk.Canvas(outer, bg=parent.cget('bg'), highlightthickness=0, bd=0)
        canvas.pack(fill="both", expand=True)
        inner = tk.Frame(outer, bg=bg_color, highlightthickness=0, bd=0)
        def configure_canvas(event):
            canvas.delete("all")
            self.draw_round_rect(canvas, 0, 0, event.width, event.height, radius, fill=bg_color, outline=bg_color)
            inner.place(x=8, y=8, width=max(event.width - 16, 0), height=max(event.height - 16, 0))
        canvas.bind("<Configure>", configure_canvas)
        return outer, inner

    def create_custom_treeview(self, parent, columns):
        container = tk.Frame(parent, bg="#050505", highlightthickness=0, bd=0)
        tree_frame = tk.Frame(container, bg="#050505", highlightthickness=0, bd=0)
        tree_frame.pack(fill="both", expand=True, padx=0, pady=0)
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse", style="Custom.Treeview")
        vsb = RoundedScrollbar(tree_frame, command=tree.yview, bg="#050505", thumb="#3A3A3A", width=12, radius=6)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        return tree, container

    def create_rounded_button(self, parent, text, command, width=120, height=35, bg_color="#00FF7F", fg_color="#021109", hover_color="#39FF14", radius=14):
        frame = tk.Frame(parent, bg=parent.cget('bg'))
        canvas = tk.Canvas(frame, width=width, height=height, bg=parent.cget('bg'), highlightthickness=0, bd=0)
        canvas.pack()
        canvas.configure(cursor="hand2")
        btn_bg = self.draw_round_rect(canvas, 2, 2, width-2, height-2, radius, fill=bg_color, outline=bg_color)
        btn_text = canvas.create_text(width//2, height//2, text=text, fill=fg_color, font=("Segoe UI", 10, "bold"))
        def on_enter(_):
            canvas.itemconfig(btn_bg, fill=hover_color)
        def on_leave(_):
            canvas.itemconfig(btn_bg, fill=bg_color)
        def on_press(_):
            canvas.itemconfig(btn_bg, fill=hover_color)
        def on_release(_):
            canvas.itemconfig(btn_bg, fill=hover_color)
            command()
        canvas.bind("<Enter>", on_enter)
        canvas.bind("<Leave>", on_leave)
        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<ButtonRelease-1>", on_release)
        return frame

    def create_rounded_tab_button(self, parent, text, tab_type, width=140, height=40, radius=8):
        frame = tk.Frame(parent, bg=parent.cget('bg'))
        canvas = tk.Canvas(frame, width=width, height=height, bg=parent.cget('bg'), highlightthickness=0, bd=0)
        canvas.pack()
        def update_appearance():
            if self.current_active_tab == tab_type:
                canvas.itemconfig(btn_bg, fill="#00FF7F")
                canvas.itemconfig(btn_text, fill="#021109")
            else:
                canvas.itemconfig(btn_bg, fill="#181818")
                canvas.itemconfig(btn_text, fill="#E6E6E6")
        def on_enter(_):
            if self.current_active_tab != tab_type:
                canvas.itemconfig(btn_bg, fill="#2A2A2A")
        def on_leave(_):
            if self.current_active_tab != tab_type:
                canvas.itemconfig(btn_bg, fill="#181818")
        def on_click(_):
            if tab_type == "global":
                self.show_global_settings()
            else:
                self.show_mail_settings()
        btn_bg = self.draw_round_rect(canvas, 2, 2, width-2, height-2, radius, fill="#181818", outline="#181818")
        btn_text = canvas.create_text(width//2, height//2, text=text, fill="#E6E6E6", font=("Segoe UI", 10, "bold"))
        canvas.bind("<Enter>", on_enter)
        canvas.bind("<Leave>", on_leave)
        canvas.bind("<Button-1>", on_click)
        update_appearance()
        return frame, update_appearance

    def create_toggle_button(self, parent, text, variable, command, width=120, height=32):
        frame = tk.Frame(parent, bg=parent.cget('bg'))
        canvas = tk.Canvas(frame, width=width, height=height, bg=parent.cget('bg'), highlightthickness=0, bd=0)
        canvas.pack()
        def update_appearance():
            if variable.get():
                canvas.itemconfig(btn_bg, fill="#00FF7F")
                canvas.itemconfig(btn_text, text=f"{text} ON", fill="#021109")
            else:
                canvas.itemconfig(btn_bg, fill="#666666")
                canvas.itemconfig(btn_text, text=f"{text} OFF", fill="#E6E6E6")
        def toggle():
            variable.set(not variable.get())
            update_appearance()
            command()
        btn_bg = self.draw_round_rect(canvas, 2, 2, width-2, height-2, 10, fill="#666666", outline="#666666")
        btn_text = canvas.create_text(width//2, height//2, text=f"{text} OFF", fill="#E6E6E6", font=("Segoe UI", 9, "bold"))
        canvas.bind("<Button-1>", lambda e: toggle())
        update_appearance()
        return frame

    def build_layout(self):
        content_frame = ttk.Frame(self.root, style="TFrame")
        content_frame.pack(fill="both", expand=True, padx=10, pady=10)
        sidebar_container = tk.Frame(content_frame, bg="#050505", highlightthickness=0, bd=0, width=260)
        sidebar_container.pack(side="left", fill="y", padx=(0, 8))
        right_container = tk.Frame(content_frame, bg="#050505", highlightthickness=0, bd=0, width=320)
        right_container.pack(side="right", fill="y", padx=(8, 0))
        center_container = tk.Frame(content_frame, bg="#050505", highlightthickness=0, bd=0)
        center_container.pack(side="left", fill="both", expand=True, padx=0)
        self.build_sidebar(sidebar_container)
        self.build_center(center_container)
        self.build_right(right_container)

    def add_stat_row(self, parent, label_text, var, color_key="neutral"):
        row = tk.Frame(parent, bg="#141414", highlightthickness=0, bd=0)
        row.pack(fill="x", padx=10, pady=1)
        icon = tk.Label(row, text=self.status_icons.get(color_key, "•"), bg="#141414", fg=self.status_colors.get(color_key, "#E6E6E6"), font=("Segoe UI", 12, "bold"))
        icon.pack(side="left", padx=(0, 6))
        label = ttk.Label(row, text=label_text, style="SidebarLabel.TLabel")
        label.pack(side="left")
        value = tk.Label(row, textvariable=var, bg="#141414", fg=self.status_colors.get(color_key, "#E6E6E6"), font=("Segoe UI", 11, "bold"))
        value.pack(side="right")

    def build_sidebar(self, parent):
        card = tk.Frame(parent, bg="#111111", highlightthickness=0, bd=0)
        card.pack(fill="both", expand=True, padx=0, pady=0)
        sidebar_top = tk.Frame(card, bg="#111111", highlightthickness=0, bd=0)
        sidebar_top.pack(fill="x", padx=8, pady=8)
        progress_container = tk.Frame(sidebar_top, bg="#101010", bd=0, highlightthickness=0)
        progress_container.pack(fill="x", pady=(8, 10))
        circle_frame = tk.Frame(progress_container, bg="#101010")
        circle_frame.pack(pady=(4, 10))
        circle_canvas = tk.Canvas(circle_frame, width=135, height=135, bg="#101010", highlightthickness=0, bd=0)
        circle_canvas.pack()
        circle_canvas.create_oval(10, 10, 125, 125, outline="#004422", width=5)
        circle_canvas.create_oval(16, 16, 119, 119, outline="#00FF7F", width=4)
        circle_canvas.create_oval(22, 22, 113, 113, outline="#021109", width=2)
        self.progress_text = circle_canvas.create_text(67, 67, text="0%", fill="#39FF14", font=("Segoe UI", 18, "bold"))
        self.progress_canvas = circle_canvas
        bar_outer = tk.Frame(progress_container, bg="#101010")
        bar_outer.pack(fill="x", padx=6, pady=(0, 10))
        bar_bg = tk.Frame(bar_outer, bg="#0A0A0A")
        bar_bg.pack(fill="x", pady=2)
        progress_bar = ttk.Progressbar(bar_bg, maximum=100, variable=self.progress_var, style="Modern.Horizontal.TProgressbar")
        progress_bar.pack(fill="x", padx=3, pady=3)
        buttons_container = tk.Frame(progress_container, bg="#101010")
        buttons_container.pack(fill="x", padx=14, pady=(0, 12))
        self.start_button = self.create_rounded_button(buttons_container, "Start", self.start_scan, width=95, height=35, bg_color="#00FF7F", fg_color="#021109", hover_color="#39FF14")
        self.start_button.pack(side="left", padx=(0, 8))
        self.stop_button = self.create_rounded_button(buttons_container, "Stop", self.stop_scan, width=95, height=35, bg_color="#FF5252", fg_color="#FFFFFF", hover_color="#FF7979")
        self.stop_button.pack(side="left")
        self.continue_button = self.create_rounded_button(buttons_container, "Continue", self.continue_scan, width=95, height=35, bg_color="#2196F3", fg_color="#FFFFFF", hover_color="#64B5F6")
        self.continue_button.pack(side="left", padx=(8, 0))
        self.continue_button.pack_forget()
        stats_container = tk.Frame(card, bg="#111111")
        stats_container.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        base_card = tk.Frame(stats_container, bg="#141414", bd=0, highlightthickness=0)
        base_card.pack(fill="x", pady=(0, 6))
        header = tk.Label(base_card, text="BASE", bg="#141414", fg="#9A9A9A", font=("Segoe UI", 9, "bold"))
        header.pack(anchor="w", padx=10, pady=(8, 4))
        self.add_stat_row(base_card, "Total lines", self.total_lines_var, "neutral")
        self.add_stat_row(base_card, "Checked", self.checked_var, "info")
        self.threads_display_var = tk.StringVar(value=str(self.threads_var.get()))
        self.add_stat_row(base_card, "Threads", self.threads_display_var, "neutral")
        status_card = tk.Frame(stats_container, bg="#141414", bd=0, highlightthickness=0)
        status_card.pack(fill="x", pady=(4, 0))
        header2 = tk.Label(status_card, text="STATUS", bg="#141414", fg="#9A9A9A", font=("Segoe UI", 9, "bold"))
        header2.pack(anchor="w", padx=10, pady=(8, 4))
        self.add_stat_row(status_card, "Good", self.good_var, "good")
        self.add_stat_row(status_card, "Found", self.found_var, "found")
        self.add_stat_row(status_card, "Bad", self.bad_var, "bad")
        self.add_stat_row(status_card, "Locked", self.locked_var, "locked")
        self.add_stat_row(status_card, "Excluded", self.excluded_var, "excluded")
        self.add_stat_row(status_card, "Errors", self.errors_var, "error")
        self.add_stat_row(status_card, "Host not found", self.host_not_found_var, "host")
        self.add_stat_row(status_card, "Multipassword", self.multipassword_var, "multipassword")
        self.sidebar_threads_label = status_card

    def update_threads_display(self):
        self.threads_display_var.set(str(self.threads_var.get()))

    def build_center(self, parent):
        search_card = tk.Frame(parent, bg="#050505", bd=0, highlightthickness=0)
        search_card.pack(fill="x", padx=0, pady=(4, 6))
        inner = tk.Frame(search_card, bg="#050505")
        inner.pack(fill="x", padx=0, pady=(8, 8))
        search_frame = tk.Frame(inner, bg="#050505")
        search_frame.pack(side="left", fill="x", expand=True)
        search_container = tk.Frame(search_frame, bg="#1A1A1A", bd=0, highlightthickness=0)
        search_container.pack(fill="x", padx=(0, 6))
        self.search_entry = tk.Entry(search_container, textvariable=self.search_var, bg="#1A1A1A", fg="#E6E6E6", insertbackground="#E6E6E6", relief="flat", font=("Segoe UI", 10), highlightthickness=1, highlightbackground="#2A2A2A", highlightcolor="#39FF14", bd=0)
        self.search_entry.pack(fill="x", expand=True, ipady=6, padx=8, pady=4)
        search_button_frame = tk.Frame(inner, bg="#050505")
        search_button_frame.pack(side="left")
        search_button = self.create_rounded_button(search_button_frame, "Search", self.apply_search, width=80, height=30, bg_color="#00FF7F", fg_color="#021109", hover_color="#39FF14")
        search_button.pack()
        status_filter_frame = tk.Frame(inner, bg="#050505")
        status_filter_frame.pack(fill="x", pady=(6, 6))

        def create_status_button(text, key):
            return self.create_rounded_button(
                status_filter_frame,
                text,
                lambda: self.apply_status_filter(key),
                width=70, height=28,
                bg_color="#181818", fg_color="#E6E6E6",
                hover_color="#333333"
            )

        self.all_button = create_status_button("All", "all")
        self.all_button.pack(side="left", padx=3)
        self.good_button = create_status_button("Good", "good")
        self.good_button.pack(side="left", padx=3)
        self.found_button = create_status_button("Found", "found")
        self.found_button.pack(side="left", padx=3)
        self.bad_button = create_status_button("Bad", "bad")
        self.bad_button.pack(side="left", padx=3)
        self.locked_button = create_status_button("Locked", "locked")
        self.locked_button.pack(side="left", padx=3)
        self.error_button = create_status_button("Error", "error")
        self.error_button.pack(side="left", padx=3)
        table_outer = tk.Frame(parent, bg="#050505", highlightthickness=0, bd=0)
        table_outer.pack(fill="both", expand=True, padx=0, pady=(0, 4))
        rounded_frame, content_frame = self.create_rounded_frame(table_outer, bg_color="#141414", radius=10)
        rounded_frame.pack(fill="both", expand=True, padx=4, pady=0)
        tree_container = tk.Frame(content_frame, bg="#141414", highlightthickness=0, bd=0)
        tree_container.pack(fill="both", expand=True, padx=0, pady=0)
        columns = ("email", "password", "mails", "keywords", "actions")
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings", selectmode="browse", style="Custom.Treeview")
        self.tree.heading("email", text="Email")
        self.tree.heading("password", text="Password")
        self.tree.heading("mails", text="Mails")
        self.tree.heading("keywords", text="Keywords")
        self.tree.heading("actions", text="Actions")
        self.tree.column("email", anchor="w")
        self.tree.column("password", anchor="center")
        self.tree.column("mails", anchor="center")
        self.tree.column("keywords", anchor="center")
        self.tree.column("actions", anchor="center")
        self.vsb = RoundedScrollbar(tree_container, command=self.tree.yview, bg="#141414", thumb="#3A3A3A", width=12, radius=6)
        self.tree.configure(yscrollcommand=self.vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")
        self.tree.tag_configure("good", foreground=self.status_colors["good"], background=self.status_bg_colors["good"])
        self.tree.tag_configure("found", foreground=self.status_colors["found"], background=self.status_bg_colors["found"])
        self.tree.tag_configure("bad", foreground=self.status_colors["bad"], background=self.status_bg_colors["bad"])
        self.tree.tag_configure("locked", foreground=self.status_colors["locked"], background=self.status_bg_colors["locked"])
        self.tree.tag_configure("error", foreground=self.status_colors["error"], background=self.status_bg_colors["error"])
        self.tree.tag_configure("excluded", foreground=self.status_colors["excluded"], background=self.status_bg_colors["excluded"])
        self.tree.tag_configure("hover", background="#242424")
        self.tree.tag_configure("rounded", background=self.status_bg_colors["neutral"], foreground=self.status_colors["neutral"])
        self.tree.bind("<Motion>", self.on_tree_hover)
        self.tree.bind("<Button-1>", self.on_tree_click)
        tree_container.bind("<Configure>", self.adjust_tree_columns)

    def adjust_tree_columns(self, event=None):
        if not hasattr(self, "tree"):
            return
        total = self.tree.winfo_width()
        if total <= 0:
            return
        columns = ("email", "password", "mails", "keywords", "actions")
        fractions = (0.30, 0.22, 0.12, 0.21, 0.15)
        for col, frac in zip(columns, fractions):
            self.tree.column(col, width=int(total * frac))

    def build_right(self, parent):
        card = tk.Frame(parent, bg="#111111", highlightthickness=0, bd=0)
        card.pack(fill="both", expand=True, padx=0, pady=0)
        tab_buttons_container = tk.Frame(card, bg="#111111")
        tab_buttons_container.pack(fill="x", padx=8, pady=(8, 4))
        self.global_tab_button, self.update_global_tab = self.create_rounded_tab_button(tab_buttons_container, "Global Settings", "global", width=150, height=40)
        self.global_tab_button.pack(side="left", padx=(0, 8))
        self.mail_tab_button, self.update_mail_tab = self.create_rounded_tab_button(tab_buttons_container, "Mail Settings", "mail", width=150, height=40)
        self.mail_tab_button.pack(side="left")
        self.settings_container = tk.Frame(card, bg="#111111")
        self.settings_container.pack(fill="both", expand=True, padx=8, pady=8)
        self.global_frame = tk.Frame(self.settings_container, bg="#111111")
        self.mail_frame = tk.Frame(self.settings_container, bg="#111111")
        self.build_global_settings(self.global_frame)
        self.build_mail_settings(self.mail_frame)
        self.show_global_settings()
        bottom_bar = tk.Frame(card, bg="#111111")
        bottom_bar.pack(side="bottom", fill="x", padx=8, pady=(0, 8))
        try:
            original_icon = tk.PhotoImage(file="telegram.png")
            self.telegram_icon = original_icon.subsample(12, 12)
        except Exception:
            self.telegram_icon = None
        if self.telegram_icon:
            telegram_label = tk.Label(bottom_bar, image=self.telegram_icon, bg="#111111", cursor="hand2")
            telegram_label.pack(side="right", padx=(0, 4), pady=(0, 4))
            telegram_label.bind("<Button-1>", lambda e: webbrowser.open("https://t.me/+WabQmkgw2641Njdl"))

    def show_global_settings(self):
        self.current_active_tab = "global"
        self.mail_frame.pack_forget()
        self.global_frame.pack(fill="both", expand=True)
        self.update_global_tab()
        self.update_mail_tab()

    def show_mail_settings(self):
        self.current_active_tab = "mail"
        self.global_frame.pack_forget()
        self.mail_frame.pack(fill="both", expand=True)
        self.update_global_tab()
        self.update_mail_tab()

    def build_global_settings(self, parent):
        parent.columnconfigure(0, weight=1)

        def create_glass_input(parent, var, min_v, max_v):
            outer = tk.Frame(parent, bg="#141414")
            box = tk.Frame(outer, bg="#1A1A1A")
            box.pack(fill="x", padx=0, pady=0)

            box.configure(
                highlightthickness=2,
                highlightbackground="#252525",
                highlightcolor="#39FF14"
            )

            sv = tk.StringVar(value=str(var.get()))
            
            def update_threads_immediately(*args):
                try:
                    v = int(sv.get())
                    if v < min_v:
                        v = min_v
                    if v > max_v:
                        v = max_v
                    var.set(v)
                    self.update_threads_display()
                except:
                    pass

            sv.trace_add("write", update_threads_immediately)

            def clamp():
                try:
                    v = int(sv.get())
                except:
                    v = min_v
                if v < min_v:
                    v = min_v
                if v > max_v:
                    v = max_v
                sv.set(str(v))
                var.set(v)
                self.update_threads_display()

            entry = tk.Entry(
                box,
                textvariable=sv,
                bg="#1A1A1A",
                fg="#E6E6E6",
                insertbackground="#E6E6E6",
                relief="flat",
                font=("Segoe UI", 10)
            )
            entry.pack(fill="x", padx=10, pady=8)

            def focus_in(e):
                box.config(highlightbackground="#39FF14")

            def focus_out(e):
                box.config(highlightbackground="#252525")
                clamp()

            entry.bind("<FocusIn>", focus_in)
            entry.bind("<FocusOut>", focus_out)

            def scroll(event):
                try:
                    v = int(sv.get())
                except:
                    v = min_v
                if event.delta > 0:
                    v += 1
                else:
                    v -= 1
                if v < min_v:
                    v = min_v
                if v > max_v:
                    v = max_v
                sv.set(str(v))
                var.set(v)
                self.update_threads_display()

            entry.bind("<MouseWheel>", scroll)

            return outer

        thread_card = tk.Frame(parent, bg="#141414", bd=0)
        thread_card.grid(row=0, column=0, sticky="ew", padx=8, pady=(10, 4))
        ttk.Label(thread_card, text="Thread Count", style="SettingsLabel.TLabel").pack(anchor="w", padx=6, pady=(6, 2))
        create_glass_input(thread_card, self.threads_var, 1, 50).pack(anchor="w", padx=6, pady=(0, 6))

        display_card = tk.Frame(parent, bg="#141414", bd=0)
        display_card.grid(row=1, column=0, sticky="ew", padx=8, pady=4)
        ttk.Label(display_card, text="Display Settings", style="SettingsLabel.TLabel").pack(anchor="w", padx=6, pady=(6, 2))
        self.create_toggle_button(display_card, "Show Password", self.show_password_var, self.toggle_password_display, width=140, height=32).pack(anchor="w", padx=6, pady=(0, 6))

        retry_card = tk.Frame(parent, bg="#141414", bd=0)
        retry_card.grid(row=2, column=0, sticky="ew", padx=8, pady=4)
        ttk.Label(retry_card, text="Max Retry Count", style="SettingsLabel.TLabel").pack(anchor="w", padx=6, pady=(6, 2))
        create_glass_input(retry_card, self.max_retry_var, 1, 10).pack(anchor="w", padx=6, pady=(0, 6))

    def build_mail_settings(self, parent):
        parent.columnconfigure(0, weight=1)
        hotmail_card = tk.Frame(parent, bg="#141414", bd=0, highlightthickness=0)
        hotmail_card.grid(row=0, column=0, sticky="ew", padx=8, pady=(10, 4))
        hotmail_label = ttk.Label(hotmail_card, text="Hotmail file", style="SettingsLabel.TLabel")
        hotmail_label.pack(anchor="w", padx=6, pady=(6, 2))
        hotmail_button = self.create_rounded_button(hotmail_card, "Upload hotmail.txt", self.select_hotmail_file, width=160, height=32, bg_color="#00FF7F", fg_color="#021109", hover_color="#39FF14")
        hotmail_button.pack(anchor="w", padx=6, pady=(2, 2))
        hotmail_name = ttk.Label(hotmail_card, textvariable=self.database_label_var, style="SettingsValue.TLabel", wraplength=280)
        hotmail_name.pack(anchor="w", padx=6, pady=(4, 6))
        keyword_card = tk.Frame(parent, bg="#141414", bd=0, highlightthickness=0)
        keyword_card.grid(row=1, column=0, sticky="ew", padx=8, pady=4)
        keyword_label = ttk.Label(keyword_card, text="Keyword file", style="SettingsLabel.TLabel")
        keyword_label.pack(anchor="w", padx=6, pady=(6, 2))
        keyword_button = self.create_rounded_button(keyword_card, "Upload keyword file", self.select_keyword_file, width=160, height=32, bg_color="#00FF7F", fg_color="#021109", hover_color="#39FF14")
        keyword_button.pack(anchor="w", padx=6, pady=(2, 2))
        keyword_name = ttk.Label(keyword_card, textvariable=self.keyword_label_var, style="SettingsValue.TLabel", wraplength=280)
        keyword_name.pack(anchor="w", padx=6, pady=(4, 6))

    def select_hotmail_file(self):
        path = filedialog.askopenfilename(title="Select hotmail.txt", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if not path:
            return
        self.hotmail_file_path = path
        self.database_label_var.set(os.path.basename(path))

    def select_keyword_file(self):
        path = filedialog.askopenfilename(title="Select keyword file", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if not path:
            return
        self.keyword_file_path = path
        self.keyword_label_var.set(os.path.basename(path))

    def reset_stats(self):
        self.total_lines = 0
        self.checked_count = 0
        self.good_count = 0
        self.found_count = 0
        self.bad_count = 0
        self.locked_count = 0
        self.excluded_count = 0
        self.error_count = 0
        self.host_not_found_count = 0
        self.multipassword_count = 0
        self.total_lines_var.set("0")
        self.checked_var.set("0")
        self.good_var.set("0")
        self.found_var.set("0")
        self.bad_var.set("0")
        self.locked_var.set("0")
        self.excluded_var.set("0")
        self.errors_var.set("0")
        self.host_not_found_var.set("0")
        self.multipassword_var.set("0")
        self.progress_var.set(0)
        self.progress_percent_var.set("0%")
        self.results_data = {"good": [], "found": [], "bad": [], "locked": [], "error": []}
        self.remaining_accounts = []
        self.current_accounts = []
        self.pending_results = {}
        self.next_display_index = 1
        self.row_data = {}
        self.all_tree_items = []
        self.idx_to_item = {}
        self.item_to_idx = {}
        self.update_progress_text()

    def get_masked_password(self, password):
        if self.show_password_var.get():
            return password
        else:
            return "●" * self.masked_password_length

    def is_valid_hotmail(self, email):
        if not email or "@" not in email:
            return False
        local, domain = email.rsplit("@", 1)
        if not local or not domain:
            return False
        if not re.match(r"^[A-Za-z0-9._%+-]+$", local):
            return False
        if not re.match(r"^[A-Za-z0-9.-]+\.[A-Za-z]{2,}$", domain):
            return False
        domain = domain.lower()
        parts = domain.split(".")
        if len(parts) < 2:
            return False
        provider = parts[0]
        if provider not in {"hotmail", "outlook", "live", "msn", "passport", "windowslive"}:
            return False
        return True    

    def start_scan(self):
        if self.executor:
            try:
                self.executor.shutdown(wait=False)
            except Exception:
                pass
            self.executor = None
        self.running = False
        self.paused = False
        if not self.hotmail_file_path or not os.path.exists(self.hotmail_file_path):
            messagebox.showerror("Error", "Please select a valid hotmail.txt file.")
            return
        self.reset_stats()
        self.tree.delete(*self.tree.get_children())
        self._hover_iid = None
        self.row_data = {}
        self.result_queue = queue.Queue()
        self.all_tree_items = []
        try:
            with open(self.hotmail_file_path, "r", encoding="utf-8") as f:
                lines = f.readlines()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        accounts = []
        index = 0
        for raw in lines:
            line = raw.strip()
            if not line:
                continue
            if ":" not in line:
                self.excluded_count += 1
                continue
            parts = line.split(":")
            if len(parts) > 2:
                self.multipassword_count += 1
            email = parts[0].strip()
            password = ":".join(parts[1:]).strip()
            if not email or not password:
                self.excluded_count += 1
                continue
            if not self.is_valid_hotmail(email):
                self.excluded_count += 1
                continue
            index += 1
            accounts.append((index, email, password))
        if not accounts:
            messagebox.showwarning("Notice", "No valid accounts found in the file.")
            self.update_sidebar_stats()
            return
        self.total_lines = len(accounts)
        self.total_lines_var.set(str(self.total_lines))
        self.excluded_var.set(str(self.excluded_count))
        self.multipassword_var.set(str(self.multipassword_count))
        for idx, email, password in accounts:
            masked = self.get_masked_password(password)
            item_id = self.tree.insert("", "end", values=(email, masked, "", "", "Open in browser"), tags=("rounded",))
            self.row_data[idx] = {"email": email, "password": password, "mails": "", "keywords": "", "result": "", "status": ""}
            self.all_tree_items.append(item_id)
            self.idx_to_item[idx] = item_id
            self.item_to_idx[item_id] = idx
        self.pending_results = {}
        self.next_display_index = 1
        self.remaining_accounts = accounts.copy()
        self.current_accounts = accounts.copy()
        self.start_scan_execution()

    def start_scan_execution(self):
        threads = self.threads_var.get()
        if threads < 1:
            threads = 1
            self.threads_var.set(1)
        max_retry = self.max_retry_var.get()
        if max_retry < 1:
            max_retry = 1
            self.max_retry_var.set(1)
        self.running = True
        self.paused = False
        self.executor = concurrent.futures.ThreadPoolExecutor(max_workers=threads)
        for idx, email, password in self.remaining_accounts:
            self.executor.submit(self.worker_task, idx, email, password, max_retry)
        self.start_button.pack_forget()
        self.stop_button.pack(side="left")
        self.continue_button.pack_forget()

    def stop_scan(self):
        if not self.running:
            return
        self.running = False
        self.paused = True
        if self.executor:
            self.executor.shutdown(wait=False)
            self.executor = None
        self.start_button.pack_forget()
        self.stop_button.pack_forget()
        self.continue_button.pack(side="left", padx=(8, 0))
        messagebox.showinfo("Paused", "Scan has been paused. Click Continue to resume.")

    def continue_scan(self):
        if self.running:
            return
        if not self.remaining_accounts:
            messagebox.showinfo("Info", "No remaining accounts to check.")
            return
        self.start_scan_execution()
        messagebox.showinfo("Resumed", "Scan has been resumed.")

    def worker_task(self, idx, email, password, max_retry):
        if not self.running:
            return
        keyword_file = self.keyword_file_path if self.keyword_file_path else None
        result = "❌ ERROR: Unknown error"
        for attempt in range(max_retry):
            if not self.running:
                return
            try:
                checker = OutlookChecker(keyword_file=self.keyword_file_path, debug=False)
                result = checker.check(email, password)
                if any(x in result for x in ["✅ HIT", "🆓 FREE", "❌ BAD", "Locked", "Need Verify", "Timeout"]):
                    break
                elif "Request Error" in result or "ERROR" in result:
                    if attempt + 1 >= max_retry:
                        break
                    time.sleep(1)
                else:
                    break
            except Exception as e:
                result = f"❌ ERROR: {str(e)}"
                if attempt + 1 >= max_retry:
                    break
                time.sleep(1)
        if self.running:
            self.result_queue.put((idx, email, password, result))

    def process_queue(self):
        try:
            while True:
                idx, email, password, result = self.result_queue.get_nowait()
                self.pending_results[idx] = (email, password, result)
                self.remaining_accounts = [acc for acc in self.remaining_accounts if acc[0] != idx]
        except queue.Empty:
            pass
        while self.next_display_index in self.pending_results:
            email, password, result = self.pending_results.pop(self.next_display_index)
            self.handle_result(self.next_display_index, email, password, result)
            self.next_display_index += 1
        self.root.after(50, self.process_queue)

    def parse_mailbox_count(self, result):
        if "✅ HIT" in result or "🆓 FREE" in result:
            if "Found:" in result:
                match = re.search(r"Found:.*?\((\d+)\)", result)
                if match:
                    return match.group(1)
            return "1+"
        return "0"

    def parse_keyword_count(self, result):
        if "✅ HIT" in result and "Found:" in result:
            found_part = result.split("Found:", 1)[1].split("|", 1)[0]
            matches = re.findall(r'([^\(\),]+?)\s*\((\d+)\)', found_part)
            if matches:
                return f"{len(matches)} keywords"
            return "0 keywords"
        elif "✅ HIT" in result:
            return "1+ keywords"
        elif "🆓 FREE" in result:
            return "0 keywords"
        return ""

    def extract_keywords_from_result(self, result):
        keywords = []
        if "✅ HIT" in result and "Found:" in result:
            found_part = result.split("Found:", 1)[1].split("|", 1)[0]
            matches = re.findall(r'([^\(\),]+?)\s*\((\d+)\)', found_part)
            keywords = [m[0].strip() for m in matches if m[0].strip()]
        return keywords

    def extract_profile_info(self, result):
        name = ""
        country = ""
        birthdate = ""
        if "|" in result:
            parts = result.split("|")
            for part in parts:
                part = part.strip()
                if not any(x in part for x in ["✅ HIT", "🆓 FREE", "❌ BAD", "Found:", "Locked", "Need Verify"]):
                    if not name and len(part) > 0 and not part.isdigit() and "keywords" not in part.lower():
                        name = part
                    elif not country and len(part) == 2:
                        country = part
                    elif not birthdate and re.match(r'\d{1,2}-\d{1,2}-\d{4}', part):
                        birthdate = part
        return name, country, birthdate

    def handle_result(self, idx, email, password, result):
        self.checked_count += 1
        self.checked_var.set(str(self.checked_count))
        status = "bad"
        keywords_found = self.extract_keywords_from_result(result)
        keyword_text = ", ".join(keywords_found) if keywords_found else ""
        name, country, birthdate = self.extract_profile_info(result)
        profile_info = []
        if name:
            profile_info.append(f"Name: {name}")
        if country:
            profile_info.append(f"Country: {country}")
        if birthdate:
            profile_info.append(f"Birth: {birthdate}")
        profile_text = " | ".join(profile_info)
        account_data = f"{email}:{password}"
        if profile_text:
            account_data += f" | {profile_text}"
        if keyword_text:
            account_data += f" | Keywords: {keyword_text}"
        
        if "✅ HIT" in result:
            status = "found"
            self.found_count += 1
            self.results_data["found"].append(account_data)
        elif "🆓 FREE" in result:
            status = "good"
            self.good_count += 1
            self.results_data["good"].append(account_data)
        elif "❌ BAD" in result:
            self.bad_count += 1
            self.results_data["bad"].append(account_data)
            if "Locked" in result or "Abuse" in result:
                self.locked_count += 1
                self.results_data["locked"].append(account_data)
                status = "locked"
            elif "Need Verify" in result:
                self.error_count += 1
                self.results_data["error"].append(account_data)
                status = "error"
            elif "Timeout" in result or "Request Error" in result:
                self.error_count += 1
                self.results_data["error"].append(account_data)
                status = "error"
        elif "ERROR" in result:
            self.error_count += 1
            self.results_data["error"].append(account_data)
            status = "error"
        else:
            self.bad_count += 1
            self.results_data["bad"].append(account_data)
        
        if "Host not found" in result or "HOST NOT FOUND" in result or "Host Not Found" in result:
            self.host_not_found_count += 1
        self.update_sidebar_stats()
        mails = self.parse_mailbox_count(result)
        keyword_count = self.parse_keyword_count(result)
        if idx in self.row_data:
            self.row_data[idx]["mails"] = mails
            self.row_data[idx]["keywords"] = keyword_text if keyword_text else keyword_count
            self.row_data[idx]["result"] = result
            self.row_data[idx]["status"] = status
        self.refresh_row(idx)
        percent = (self.checked_count / self.total_lines) * 100 if self.total_lines > 0 else 0
        percent = min(percent, 100)
        self.progress_var.set(percent)
        self.progress_percent_var.set(f"{int(percent)}%")
        self.update_progress_text()
        if self.checked_count >= self.total_lines and self.running:
            self.running = False
            if self.executor:
                self.executor.shutdown(wait=False)
            self.export_separated_results_auto()
            self.start_button.pack(side="left", padx=(0, 8))
            self.stop_button.pack_forget()
            self.continue_button.pack_forget()
            messagebox.showinfo("Completed", "Scan completed successfully!")

    def update_sidebar_stats(self):
        self.good_var.set(str(self.good_count))
        self.found_var.set(str(self.found_count))
        self.bad_var.set(str(self.bad_count))
        self.locked_var.set(str(self.locked_count))
        self.excluded_var.set(str(self.excluded_count))
        self.errors_var.set(str(self.error_count))
        self.host_not_found_var.set(str(self.host_not_found_count))
        self.multipassword_var.set(str(self.multipassword_count))

    def update_progress_text(self):
        self.progress_canvas.itemconfigure(self.progress_text, text=self.progress_percent_var.get())

    def refresh_row(self, idx):
        item_id = self.idx_to_item.get(idx)
        if not item_id or not self.tree.exists(item_id):
            return
        data = self.row_data.get(idx)
        if not data:
            return
        email = data["email"]
        password = data["password"]
        mails = data["mails"]
        keywords = data["keywords"]
        status = data.get("status", "")
        masked = self.get_masked_password(password)
        icon = self.status_icons.get(status, "")
        tags = ("rounded",)
        if status:
            tags = ("rounded", status)
        self.tree.item(item_id, values=(f"{icon} {email}", masked, mails, keywords, "Open in browser"), tags=tags)

    def on_tree_hover(self, event):
        row_id = self.tree.identify_row(event.y)
        if self._hover_iid and self._hover_iid != row_id:
            if self.tree.exists(self._hover_iid):
                old_tags = list(self.tree.item(self._hover_iid, "tags") or ())
                if "hover" in old_tags:
                    old_tags.remove("hover")
                    self.tree.item(self._hover_iid, tags=tuple(old_tags))
            self._hover_iid = None
        if row_id and self.tree.exists(row_id):
            tags = list(self.tree.item(row_id, "tags") or ())
            if "hover" not in tags:
                tags.append("hover")
                self.tree.item(row_id, tags=tuple(tags))
            self._hover_iid = row_id

    def on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not row_id:
            return
        if col == "#5":
            idx = self.item_to_idx.get(row_id)
            if idx is None:
                return
            email = self.row_data[idx]["email"]
            webbrowser.open(f"https://outlook.live.com/mail/0/inbox?login_hint={email}")

    def toggle_password_display(self):
        for iid, data in self.row_data.items():
            self.refresh_row(iid)

    def on_search_change(self, *args):
        self.apply_search()

    def apply_search(self):
        q = self.search_var.get().strip().lower()
        if not q:
            for item in self.all_tree_items:
                self.tree.reattach(item, "", "end")
            return
        for item in self.all_tree_items:
            vals = self.tree.item(item, "values")
            show = False
            if vals:
                show = any(q in str(v).lower() for v in vals[:4])
            if show:
                self.tree.reattach(item, "", "end")
            else:
                self.tree.detach(item)

    def apply_status_filter(self, status):
        self.current_status_filter = status
        for item in self.all_tree_items:
            idx = self.item_to_idx.get(item)
            data = self.row_data.get(idx, {}) if idx is not None else {}
            item_status = data.get("status", "")
            if status == "all":
                self.tree.reattach(item, "", "end")
            elif status == "good" and item_status == "good":
                self.tree.reattach(item, "", "end")
            elif status == "found" and item_status == "found":
                self.tree.reattach(item, "", "end")
            elif status == "bad" and item_status == "bad":
                self.tree.reattach(item, "", "end")
            elif status == "locked" and item_status == "locked":
                self.tree.reattach(item, "", "end")
            elif status == "error" and item_status == "error":
                self.tree.reattach(item, "", "end")
            else:
                self.tree.detach(item)

    def export_separated_results_auto(self):
        base_dir = os.path.dirname(self.hotmail_file_path) if self.hotmail_file_path else os.getcwd()
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        files_created = []
        if self.results_data["good"]:
            good_file = os.path.join(base_dir, f"hotmail_good_{timestamp}.txt")
            with open(good_file, "w", encoding="utf-8") as f:
                f.write("\n".join(self.results_data["good"]))
            files_created.append(f"Good accounts: {good_file}")
        if self.results_data["found"]:
            found_file = os.path.join(base_dir, f"hotmail_found_{timestamp}.txt")
            with open(found_file, "w", encoding="utf-8") as f:
                f.write("\n".join(self.results_data["found"]))
            files_created.append(f"Found accounts: {found_file}")
        if self.results_data["bad"]:
            bad_file = os.path.join(base_dir, f"hotmail_bad_{timestamp}.txt")
            with open(bad_file, "w", encoding="utf-8") as f:
                f.write("\n".join(self.results_data["bad"]))
            files_created.append(f"Bad accounts: {bad_file}")
        if self.results_data["locked"]:
            locked_file = os.path.join(base_dir, f"hotmail_locked_{timestamp}.txt")
            with open(locked_file, "w", encoding="utf-8") as f:
                f.write("\n".join(self.results_data["locked"]))
            files_created.append(f"Locked accounts: {locked_file}")
        if self.results_data["error"]:
            error_file = os.path.join(base_dir, f"hotmail_error_{timestamp}.txt")
            with open(error_file, "w", encoding="utf-8") as f:
                f.write("\n".join(self.results_data["error"]))
            files_created.append(f"Error accounts: {error_file}")
        if files_created:
            messagebox.showinfo("Export Completed", "Results exported to separate files:\n\n" + "\n".join(files_created))

    def export_separated_results(self):
        if not any(self.results_data.values()):
            messagebox.showinfo("Info", "No results to export.")
            return
        base_dir = filedialog.askdirectory(title="Select folder to save results")
        if not base_dir:
            return
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        files_created = []
        if self.results_data["good"]:
            good_file = os.path.join(base_dir, f"hotmail_good_{timestamp}.txt")
            with open(good_file, "w", encoding="utf-8") as f:
                f.write("\n".join(self.results_data["good"]))
            files_created.append(f"Good accounts: {good_file}")
        if self.results_data["found"]:
            found_file = os.path.join(base_dir, f"hotmail_found_{timestamp}.txt")
            with open(found_file, "w", encoding="utf-8") as f:
                f.write("\n".join(self.results_data["found"]))
            files_created.append(f"Found accounts: {found_file}")
        if self.results_data["bad"]:
            bad_file = os.path.join(base_dir, f"hotmail_bad_{timestamp}.txt")
            with open(bad_file, "w", encoding="utf-8") as f:
                f.write("\n".join(self.results_data["bad"]))
            files_created.append(f"Bad accounts: {bad_file}")
        if self.results_data["locked"]:
            locked_file = os.path.join(base_dir, f"hotmail_locked_{timestamp}.txt")
            with open(locked_file, "w", encoding="utf-8") as f:
                f.write("\n".join(self.results_data["locked"]))
            files_created.append(f"Locked accounts: {locked_file}")
        if self.results_data["error"]:
            error_file = os.path.join(base_dir, f"hotmail_error_{timestamp}.txt")
            with open(error_file, "w", encoding="utf-8") as f:
                f.write("\n".join(self.results_data["error"]))
            files_created.append(f"Error accounts: {error_file}")
        messagebox.showinfo("Export Completed", "Results exported to separate files:\n\n" + "\n".join(files_created))

    def on_close(self):
        try:
            if self.executor:
                self.executor.shutdown(wait=False)
        except Exception:
            pass
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = HotmailCloudAouto(root)
    root.mainloop()
