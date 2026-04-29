# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
import json
import threading
import urllib.request
from datetime import datetime

try:
    from version import VERSION
except ImportError:
    VERSION = "dev"

GITHUB_REPO = "DJAVDH/Excellent-App"

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

import supabase_client as sc
from labelMaker import LabelMakerComponent
from assistentieBon import AssistentieBonComponent
from kastBon import KastBonComponent
from login import LoginWindow

# ── Colour palette ───────────────────────────────────────────────────────────
BG_ROOT      = "#F0F2F5"   # lichtgrijs pagina-achtergrond
BG_SIDEBAR   = "#FFFFFF"   # wit sidebar
BG_CONTENT   = "#FFFFFF"   # wit content-kaart
BG_CARD      = "#FFFFFF"   # wit kaart
BG_DARK_CARD = "#1A1A2E"   # donkere samenvattings-kaart
ACCENT       = "#1A1A2E"   # donker navy
ACCENT_LIGHT = "#E8F5E9"   # lichtgroen tint (hover/actief)
GREEN        = "#4CAF50"   # groen accent
BLUE         = "#2196F3"
ORANGE       = "#FF9800"
PURPLE       = "#9C27B0"
TEXT_DARK    = "#1A1A2E"
TEXT_MUTED   = "#8A8FA8"
TEXT_LIGHT   = "#FFFFFF"
BORDER       = "#E8EAF0"
FONT         = "Segoe UI"

# ── Unicode icons ────────────────────────────────────────────────────────────
ICON_DASHBOARD   = "⊞"
ICON_LABEL       = "🏷"
ICON_ASSISTENTIE = "📋"
ICON_KAST        = "📦"

# ── Kaart-helper ─────────────────────────────────────────────────────────────
def make_card(parent, bg=BG_CARD, accent_color=None, padx=16, pady=14):
    """Maak een witte kaart met optionele gekleurde top-balk."""
    outer = tk.Frame(parent, bg=bg,
                     highlightbackground=BORDER, highlightthickness=1)
    if accent_color:
        tk.Frame(outer, bg=accent_color, height=4).pack(fill=tk.X)
    inner = tk.Frame(outer, bg=bg)
    inner.pack(fill=tk.BOTH, expand=True, padx=padx, pady=pady)
    return outer, inner


class ExcellentApp:
    def __init__(self, root):
        self.root = root
        root.title("Excellent Packing and Moving B.V.")
        root.configure(bg=BG_ROOT)
        root.geometry("1100x720")
        root.minsize(860, 560)

        # ── App-icoon ────────────────────────────────────────────────────────
        ico_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "logo.ico")
        if os.path.exists(ico_path):
            root.iconbitmap(ico_path)

        # ── ttk-stijl ────────────────────────────────────────────────────────
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TCombobox",
                        fieldbackground="#F7F9FC", background="#F7F9FC",
                        foreground=TEXT_DARK, bordercolor=BORDER,
                        arrowcolor=TEXT_MUTED, font=(FONT, 10))
        style.configure("TEntry",
                        fieldbackground="#F7F9FC", foreground=TEXT_DARK,
                        bordercolor=BORDER, font=(FONT, 10))
        style.configure("Vertical.TScrollbar",
                        background=BORDER, troughcolor=BG_CONTENT,
                        bordercolor=BORDER, arrowcolor=TEXT_MUTED)

        # ── Buitenste wrapper ────────────────────────────────────────────────
        outer = tk.Frame(root, bg=BG_ROOT)
        outer.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # ════════════════════════════════════════════════════════════════════
        # SIDEBAR  (vaste breedte 230px)
        # ════════════════════════════════════════════════════════════════════
        self.sidebar = tk.Frame(outer, bg=BG_SIDEBAR, width=230,
                                highlightbackground=BORDER, highlightthickness=1)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 16))
        self.sidebar.pack_propagate(False)

        self._build_sidebar()

        # ════════════════════════════════════════════════════════════════════
        # HOOFD-CONTENT  (vult de rest)
        # ════════════════════════════════════════════════════════════════════
        right = tk.Frame(outer, bg=BG_ROOT)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Top-balk
        topbar = tk.Frame(right, bg=BG_CONTENT, height=58,
                          highlightbackground=BORDER, highlightthickness=1)
        topbar.pack(fill=tk.X, pady=(0, 16))
        topbar.pack_propagate(False)

        self.page_title_var = tk.StringVar(value="Dashboard")
        tk.Label(topbar, textvariable=self.page_title_var,
                 bg=BG_CONTENT, fg=TEXT_DARK,
                 font=(FONT, 15, "bold"),
                 anchor="w", padx=24).pack(side=tk.LEFT, fill=tk.Y)

        # Content-frame (wordt gewist bij elke paginawissel)
        self.content_frame = tk.Frame(right, bg=BG_ROOT)
        self.content_frame.pack(fill=tk.BOTH, expand=True)

        self._go_dashboard()

    # ════════════════════════════════════════════════════════════════════════
    # SIDEBAR OPBOUW
    # ════════════════════════════════════════════════════════════════════════

    def _build_sidebar(self):
        sb = self.sidebar

        # ── Logo banner ──────────────────────────────────────────────────────
        logo_frame = tk.Frame(sb, bg=BG_SIDEBAR)
        logo_frame.pack(fill=tk.X, pady=(0, 0))

        banner_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "logobanner.png")
        self._logo_img = None

        if PIL_AVAILABLE and os.path.exists(banner_path):
            try:
                img = Image.open(banner_path)
                if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
                    base = Image.new("RGB", img.size, (255, 255, 255))
                    if img.mode != "RGBA":
                        img = img.convert("RGBA")
                    base.paste(img, mask=img.split()[3])
                    img = base
                else:
                    img = img.convert("RGB")
                target_w = 230
                ratio = target_w / img.width
                target_h = min(int(img.height * ratio), 90)
                img = img.resize((target_w, target_h), Image.LANCZOS)
                self._logo_img = ImageTk.PhotoImage(img)
                tk.Label(logo_frame, image=self._logo_img,
                         bg=BG_SIDEBAR, anchor="center").pack(fill=tk.X)
            except Exception:
                self._logo_img = None

        if self._logo_img is None:
            tk.Label(logo_frame, text="Excellent",
                     bg=BG_SIDEBAR, fg=ACCENT,
                     font=(FONT, 15, "bold"),
                     anchor="center", pady=22).pack(fill=tk.X)

        # ── Bedrijfsnaam ─────────────────────────────────────────────────────
        tk.Label(sb, text="Excellent Packing and Moving B.V.",
                 bg=BG_SIDEBAR, fg=TEXT_MUTED,
                 font=(FONT, 7), anchor="center",
                 wraplength=210).pack(fill=tk.X, padx=12, pady=(4, 0))

        # ── Scheidingslijn ───────────────────────────────────────────────────
        tk.Frame(sb, bg=BORDER, height=1).pack(fill=tk.X, padx=16, pady=10)

        # ── Menu-label ───────────────────────────────────────────────────────
        tk.Label(sb, text="NAVIGATIE", bg=BG_SIDEBAR, fg=TEXT_MUTED,
                 font=(FONT, 7, "bold"), anchor="w",
                 padx=20).pack(fill=tk.X, pady=(4, 4))

        # ── Nav-knoppen ──────────────────────────────────────────────────────
        self._active_nav = None
        self._nav_buttons = {}

        self._make_nav_btn(ICON_DASHBOARD,   "Dashboard",       self._go_dashboard)
        self._make_nav_btn(ICON_LABEL,       "Label Maker",     self.show_label_maker)
        self._make_nav_btn(ICON_ASSISTENTIE, "Assistentie Bon", self.show_assistentie_bon)
        self._make_nav_btn(ICON_KAST,        "123kast Bon",     self.show_kast_bon)
        if sc.is_admin():
            self._make_nav_btn("⚙", "Admin", self.show_admin_page)

        # ── Spacer ───────────────────────────────────────────────────────────
        tk.Frame(sb, bg=BG_SIDEBAR).pack(fill=tk.BOTH, expand=True)

        # ── Altijd vooraan toggle ────────────────────────────────────────────
        tk.Frame(sb, bg=BORDER, height=1).pack(fill=tk.X, padx=16)

        topmost_row = tk.Frame(sb, bg=BG_SIDEBAR)
        topmost_row.pack(fill=tk.X, padx=16, pady=(8, 4))

        self._topmost_var = tk.BooleanVar(value=False)

        self._topmost_indicator = tk.Canvas(topmost_row, width=14, height=14,
                                            bg=BG_SIDEBAR, highlightthickness=0)
        self._topmost_indicator.pack(side=tk.LEFT, padx=(4, 6))
        self._draw_topmost_indicator(False)

        tk.Checkbutton(
            topmost_row,
            text="Altijd vooraan",
            variable=self._topmost_var,
            command=self._toggle_topmost,
            bg=BG_SIDEBAR, fg=TEXT_MUTED,
            activebackground=BG_SIDEBAR, activeforeground=TEXT_DARK,
            selectcolor=BG_SIDEBAR,
            font=(FONT, 8),
            bd=0, relief=tk.FLAT,
            cursor="hand2",
            anchor="w"
        ).pack(side=tk.LEFT)

        # ── Scheidingslijn ───────────────────────────────────────────────────
        tk.Frame(sb, bg=BORDER, height=1).pack(fill=tk.X, padx=16)

        # ── Gebruiker / avatar + uitloggen ───────────────────────────────────
        email = sc.get_user_email()
        initials = (email[:2].upper() if email else "EP")

        user_row = tk.Frame(sb, bg=BG_SIDEBAR)
        user_row.pack(fill=tk.X, padx=16, pady=(10, 6))

        av = tk.Canvas(user_row, width=36, height=36,
                       bg=BG_SIDEBAR, highlightthickness=0)
        av.pack(side=tk.LEFT)
        av.create_oval(2, 2, 34, 34, fill=ACCENT, outline="")
        av.create_text(18, 18, text=initials, fill=TEXT_LIGHT,
                       font=(FONT, 9, "bold"))

        ui = tk.Frame(user_row, bg=BG_SIDEBAR)
        ui.pack(side=tk.LEFT, padx=(10, 0))
        display = email if email else "Gebruiker"
        tk.Label(ui, text=display, bg=BG_SIDEBAR, fg=TEXT_DARK,
                 font=(FONT, 9, "bold"), anchor="w",
                 wraplength=130).pack(anchor="w")
        tk.Label(ui, text="Ingelogd", bg=BG_SIDEBAR, fg=TEXT_MUTED,
                 font=(FONT, 7), anchor="w").pack(anchor="w")

        logout_btn = tk.Button(sb, text="↩  Uitloggen", command=self._do_logout,
                               bg="#FFF0F0", fg="#C62828",
                               activebackground="#FFCDD2", activeforeground="#B71C1C",
                               font=(FONT, 9), bd=0, relief=tk.FLAT,
                               pady=7, cursor="hand2")
        logout_btn.pack(fill=tk.X, padx=16, pady=(8, 10))

        # ── Versie ───────────────────────────────────────────────────────────
        tk.Frame(sb, bg=BORDER, height=1).pack(fill=tk.X, padx=16)
        tk.Label(sb, text=f"{VERSION}  •  Excellent App",
                 bg=BG_SIDEBAR, fg=TEXT_MUTED,
                 font=(FONT, 7), anchor="w",
                 padx=20).pack(fill=tk.X, pady=8)

    def _make_nav_btn(self, icon, label, command):
        full = f"{icon}   {label}"
        row = tk.Frame(self.sidebar, bg=BG_SIDEBAR, cursor="hand2")
        row.pack(fill=tk.X, padx=10, pady=2)

        bar = tk.Frame(row, bg=BG_SIDEBAR, width=4)
        bar.pack(side=tk.LEFT, fill=tk.Y)

        btn = tk.Button(row, text=full,
                        bg=BG_SIDEBAR, fg=TEXT_MUTED,
                        activebackground=ACCENT_LIGHT, activeforeground=ACCENT,
                        font=(FONT, 10), anchor="w",
                        padx=12, pady=10, bd=0, relief=tk.FLAT, cursor="hand2",
                        command=lambda lbl=label, cmd=command: self._nav_click(lbl, cmd))
        btn.pack(side=tk.LEFT, fill=tk.X, expand=True)

        def enter(e, b=btn, r=row, ab=bar):
            if b is not self._active_nav:
                b.config(bg=ACCENT_LIGHT, fg=ACCENT)
                r.config(bg=ACCENT_LIGHT)
                ab.config(bg=ACCENT_LIGHT)

        def leave(e, b=btn, r=row, ab=bar):
            if b is not self._active_nav:
                b.config(bg=BG_SIDEBAR, fg=TEXT_MUTED)
                r.config(bg=BG_SIDEBAR)
                ab.config(bg=BG_SIDEBAR)

        btn.bind("<Enter>", enter)
        btn.bind("<Leave>", leave)
        row.bind("<Enter>", enter)
        row.bind("<Leave>", leave)

        self._nav_buttons[label] = (btn, row, bar)

    def _draw_topmost_indicator(self, active):
        """Teken een klein bolletje: groen als actief, grijs als inactief."""
        c = self._topmost_indicator
        c.delete("all")
        color = GREEN if active else BORDER
        c.create_oval(2, 2, 12, 12, fill=color, outline="")

    def _do_logout(self):
        sc.logout()
        self.root.destroy()
        _start_login()

    def _toggle_topmost(self):
        """Zet het venster altijd vooraan of niet, voor de hele app."""
        active = self._topmost_var.get()
        self.root.attributes("-topmost", active)
        self._draw_topmost_indicator(active)

    def _nav_click(self, label, cmd):
        for lbl, (b, r, ab) in self._nav_buttons.items():
            b.config(bg=BG_SIDEBAR, fg=TEXT_MUTED, font=(FONT, 10))
            r.config(bg=BG_SIDEBAR)
            ab.config(bg=BG_SIDEBAR)
        btn, row, bar = self._nav_buttons[label]
        btn.config(bg=ACCENT_LIGHT, fg=ACCENT, font=(FONT, 10, "bold"))
        row.config(bg=ACCENT_LIGHT)
        bar.config(bg=GREEN)
        self._active_nav = btn
        cmd()

    # ════════════════════════════════════════════════════════════════════════
    # STATISTIEKEN  (lezen / schrijven naar app_stats.json)
    # ════════════════════════════════════════════════════════════════════════

    STATS_FILE = os.path.join(os.path.abspath(os.path.dirname(__file__)),
                              "app_stats.json")
    _DEFAULT_STATS = {
        "totaal_labels_gegenereerd": 0,
        "totaal_assistentie_bonnen": 0,
        "totaal_kast_bonnen":        0,
        "laatste_label_datum": "",
        "laatste_bon_datum":   "",
        "laatste_kast_datum":  "",
    }

    def load_stats(self):
        """Lees stats uit Supabase; val terug op lokale JSON bij geen verbinding."""
        try:
            res = sc.get_client().table("app_stats").select("*").eq("id", 1).execute()
            if res.data:
                row = res.data[0]
                return {
                    "totaal_labels_gegenereerd": row.get("totaal_labels_gegenereerd", 0),
                    "totaal_assistentie_bonnen": row.get("totaal_assistentie_bonnen", 0),
                    "totaal_kast_bonnen":        row.get("totaal_kast_bonnen", 0),
                    "laatste_label_datum":       row.get("laatste_label_datum", ""),
                    "laatste_bon_datum":         row.get("laatste_bon_datum", ""),
                    "laatste_kast_datum":        row.get("laatste_kast_datum", ""),
                }
        except Exception:
            pass
        # Fallback: lokale JSON
        if not os.path.exists(self.STATS_FILE):
            self.save_stats(self._DEFAULT_STATS.copy())
        try:
            with open(self.STATS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            for k, v in self._DEFAULT_STATS.items():
                data.setdefault(k, v)
            return data
        except Exception:
            return self._DEFAULT_STATS.copy()

    def save_stats(self, data):
        """Schrijf stats naar lokale JSON (alleen als fallback)."""
        try:
            with open(self.STATS_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
        except Exception:
            pass

    def increment_labels(self, aantal):
        """Verhoog label-teller atomisch via Supabase RPC."""
        try:
            sc.get_client().rpc("increment_labels", {"aantal": aantal}).execute()
            sc.log_activity("label", aantal)
            return
        except Exception:
            pass
        data = self.load_stats()
        data["totaal_labels_gegenereerd"] += aantal
        data["laatste_label_datum"] = datetime.now().strftime("%d-%m-%Y %H:%M")
        self.save_stats(data)

    def increment_bons(self):
        """Verhoog bon-teller atomisch via Supabase RPC."""
        try:
            sc.get_client().rpc("increment_bons", {}).execute()
            sc.log_activity("bon", 1)
            return
        except Exception:
            pass
        data = self.load_stats()
        data["totaal_assistentie_bonnen"] += 1
        data["laatste_bon_datum"] = datetime.now().strftime("%d-%m-%Y %H:%M")
        self.save_stats(data)

    def increment_kast_bonnen(self):
        """Verhoog 123kast bon-teller atomisch via Supabase RPC."""
        try:
            sc.get_client().rpc("increment_kast_bonnen", {}).execute()
            sc.log_activity("kast_bon", 1)
            return
        except Exception:
            pass
        data = self.load_stats()
        data["totaal_kast_bonnen"] += 1
        data["laatste_kast_datum"] = datetime.now().strftime("%d-%m-%Y %H:%M")
        self.save_stats(data)

    # ════════════════════════════════════════════════════════════════════════
    # CONTENT HELPERS
    # ════════════════════════════════════════════════════════════════════════

    def _clear_content(self):
        for w in self.content_frame.winfo_children():
            w.destroy()

    def _scrollable(self):
        """Maak een scrollbare frame en geef de inner frame terug."""
        wrap = tk.Frame(self.content_frame, bg=BG_ROOT)
        wrap.pack(fill=tk.BOTH, expand=True)

        cv = tk.Canvas(wrap, bg=BG_ROOT, highlightthickness=0)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=cv.yview)
        inner = tk.Frame(cv, bg=BG_ROOT)

        inner.bind("<Configure>",
                   lambda e: cv.configure(scrollregion=cv.bbox("all")))
        win_id = cv.create_window((0, 0), window=inner, anchor="nw")
        cv.configure(yscrollcommand=sb.set)

        cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        cv.bind("<Configure>", lambda e: cv.itemconfig(win_id, width=e.width))

        def _scroll(e):
            cv.yview_scroll(int(-1 * (e.delta / 120)), "units")
        cv.bind_all("<MouseWheel>", _scroll)

        return inner

    # ════════════════════════════════════════════════════════════════════════
    # DASHBOARD
    # ════════════════════════════════════════════════════════════════════════

    def _go_dashboard(self):
        self.page_title_var.set("Dashboard")
        self._clear_content()
        page = self._scrollable()

        # ── Rij A: Welkom-banner ─────────────────────────────────────────────
        banner, bi = make_card(page, bg="#F7F9FC", padx=24, pady=18)
        banner.pack(fill=tk.X, pady=(0, 16))
        tk.Label(bi, text="Welkom terug \U0001f44b",
                 bg="#F7F9FC", fg=TEXT_DARK,
                 font=(FONT, 14, "bold"), anchor="w").pack(anchor="w")
        tk.Label(bi,
                 text="Kies een module in het menu aan de linkerkant om te beginnen.",
                 bg="#F7F9FC", fg=TEXT_MUTED,
                 font=(FONT, 10), anchor="w").pack(anchor="w", pady=(4, 0))

        # ── Rij B: Snelkoppelings-tegels ─────────────────────────────────────
        tiles_row = tk.Frame(page, bg=BG_ROOT)
        tiles_row.pack(fill=tk.X, pady=(0, 16))
        for i in range(3):
            tiles_row.columnconfigure(i, weight=1)

        self._tile(tiles_row, ICON_LABEL, "Label Maker",
                   "Genereer en print verhuislabels",
                   GREEN, self.show_label_maker, col=0)
        self._tile(tiles_row, ICON_ASSISTENTIE, "Assistentie Bon",
                   "Maak een assistentie bon PDF",
                   BLUE, self.show_assistentie_bon, col=1)
        self._tile(tiles_row, ICON_KAST, "123kast Bon",
                   "Leveringsbon & tevredenheidsonderzoek",
                   ORANGE, self.show_kast_bon, col=2)

        # ── Rij C: Statistieken (echte data uit JSON) ────────────────────────
        stats = self.load_stats()
        labels_totaal  = stats["totaal_labels_gegenereerd"]
        bons_totaal    = stats["totaal_assistentie_bonnen"]
        kastbon_totaal = stats["totaal_kast_bonnen"]
        laatste_label  = stats["laatste_label_datum"] or "—"
        laatste_bon    = stats["laatste_bon_datum"]   or "—"
        laatste_kast   = stats["laatste_kast_datum"]  or "—"

        stats_row = tk.Frame(page, bg=BG_ROOT)
        stats_row.pack(fill=tk.X, pady=(0, 16))
        for i in range(4):
            stats_row.columnconfigure(i, weight=1)

        self._stat_card(stats_row, "Gegenereerde Labels",
                        str(labels_totaal),
                        f"Laatste: {laatste_label}", GREEN, col=0)
        self._stat_card(stats_row, "Assistentie Bons",
                        str(bons_totaal),
                        f"Laatste: {laatste_bon}", BLUE, col=1)
        self._stat_card(stats_row, "123kast Bons",
                        str(kastbon_totaal),
                        f"Laatste: {laatste_kast}", ORANGE, col=2)
        self._totaal_card(stats_row, labels_totaal, bons_totaal, kastbon_totaal, col=3)

        # ── Rij D: Activiteit + Module gebruik ──────────────────────────────
        bottom_row = tk.Frame(page, bg=BG_ROOT)
        bottom_row.pack(fill=tk.X, pady=(0, 8))
        bottom_row.columnconfigure(0, weight=3)
        bottom_row.columnconfigure(1, weight=2)

        self._activiteit_card(bottom_row, laatste_label, laatste_bon, laatste_kast, col=0)
        self._progress_card(bottom_row, labels_totaal, bons_totaal, col=1)

    # ── Dashboard kaart-helpers ──────────────────────────────────────────────

    def _tile(self, parent, icon, title, subtitle, color, command, col):
        card, inner = make_card(parent, accent_color=color, padx=20, pady=16)
        card.grid(row=0, column=col, sticky="nsew",
                  padx=(0, 12) if col == 0 else 0, pady=0)
        card.config(cursor="hand2")

        top = tk.Frame(inner, bg=BG_CARD)
        top.pack(fill=tk.X)
        tk.Label(top, text=icon, bg=BG_CARD, fg=color,
                 font=(FONT, 20)).pack(side=tk.LEFT)

        tk.Label(inner, text=title, bg=BG_CARD, fg=TEXT_DARK,
                 font=(FONT, 12, "bold"), anchor="w").pack(anchor="w", pady=(8, 2))
        tk.Label(inner, text=subtitle, bg=BG_CARD, fg=TEXT_MUTED,
                 font=(FONT, 9), anchor="w").pack(anchor="w")

        def _bind(w):
            w.bind("<Button-1>", lambda e: command())
            w.bind("<Enter>", lambda e: card.config(bg="#F7F9FC",
                   highlightbackground=color))
            w.bind("<Leave>", lambda e: card.config(bg=BG_CARD,
                   highlightbackground=BORDER))
        _bind(card)
        for child in card.winfo_children():
            _bind(child)
            for gc in child.winfo_children():
                _bind(gc)

    def _stat_card(self, parent, title, value, subtitle, color, col):
        card, inner = make_card(parent, accent_color=color, padx=18, pady=16)
        card.grid(row=0, column=col, sticky="nsew",
                  padx=(0, 12) if col < 3 else 0)

        tk.Label(inner, text=title, bg=BG_CARD, fg=TEXT_MUTED,
                 font=(FONT, 9), anchor="w").pack(anchor="w")
        tk.Label(inner, text=value, bg=BG_CARD, fg=TEXT_DARK,
                 font=(FONT, 28, "bold"), anchor="w").pack(anchor="w", pady=(4, 2))
        tk.Label(inner, text=subtitle, bg=BG_CARD, fg=color,
                 font=(FONT, 8), anchor="w").pack(anchor="w")

        # Mini bar-grafiek
        bf = tk.Frame(inner, bg=BG_CARD)
        bf.pack(anchor="w", pady=(12, 0))
        for h in [12, 20, 16, 28, 18, 24]:
            b = tk.Frame(bf, bg=color, width=9, height=h)
            b.pack(side=tk.LEFT, padx=2)
            b.pack_propagate(False)

    def _totaal_card(self, parent, labels_totaal, bons_totaal, kastbon_totaal, col):
        """Donkere samenvattings-kaart met echte totalen."""
        card = tk.Frame(parent, bg=BG_DARK_CARD,
                        highlightbackground=BORDER, highlightthickness=1)
        card.grid(row=0, column=col, sticky="nsew")

        inner = tk.Frame(card, bg=BG_DARK_CARD)
        inner.pack(fill=tk.BOTH, expand=True, padx=18, pady=16)

        tk.Label(inner, text="Totaal activiteit", bg=BG_DARK_CARD, fg="#8A8FA8",
                 font=(FONT, 9), anchor="w").pack(anchor="w")
        totaal = labels_totaal + bons_totaal + kastbon_totaal
        tk.Label(inner, text=str(totaal), bg=BG_DARK_CARD, fg=TEXT_LIGHT,
                 font=(FONT, 28, "bold"), anchor="w").pack(anchor="w", pady=(4, 2))
        tk.Label(inner, text="TOTAAL", bg=GREEN, fg=TEXT_LIGHT,
                 font=(FONT, 7, "bold"), padx=6).pack(anchor="w", pady=(0, 10))

        tk.Label(inner, text=f"Labels: {labels_totaal}",
                 bg=BG_DARK_CARD, fg="#8A8FA8",
                 font=(FONT, 8), anchor="w").pack(anchor="w", pady=(2, 0))
        tk.Label(inner, text=f"Bons: {bons_totaal}",
                 bg=BG_DARK_CARD, fg="#8A8FA8",
                 font=(FONT, 8), anchor="w").pack(anchor="w", pady=(2, 0))
        tk.Label(inner, text=f"123kast: {kastbon_totaal}",
                 bg=BG_DARK_CARD, fg="#8A8FA8",
                 font=(FONT, 8), anchor="w").pack(anchor="w", pady=(2, 4))

        # Progress-balk (verhouding labels vs bons)
        if totaal > 0:
            pct = labels_totaal / totaal
        else:
            pct = 0.5
        track = tk.Frame(inner, bg="#2E2E4E", height=8)
        track.pack(fill=tk.X)
        track.pack_propagate(False)
        bar_cv = tk.Canvas(track, bg="#2E2E4E", height=8, highlightthickness=0)
        bar_cv.pack(fill=tk.BOTH, expand=True)

        def draw(event=None, c=bar_cv, p=pct):
            w = c.winfo_width() or 160
            c.delete("all")
            c.create_rectangle(0, 0, int(w * p), 8, fill=GREEN, outline="")
        bar_cv.bind("<Configure>", draw)
        bar_cv.after(120, draw)

    def _activiteit_card(self, parent, laatste_label, laatste_bon, laatste_kast, col):
        """Kaart met recente activiteiten per gebruiker uit Supabase."""
        card, inner = make_card(parent, accent_color=GREEN, padx=18, pady=16)
        card.grid(row=0, column=col, sticky="nsew", padx=(0, 12))

        tk.Label(inner, text="Recente activiteit per gebruiker", bg=BG_CARD,
                 fg=TEXT_MUTED, font=(FONT, 9), anchor="w").pack(anchor="w")
        tk.Label(inner, text="Laatste acties", bg=BG_CARD, fg=TEXT_DARK,
                 font=(FONT, 12, "bold"), anchor="w").pack(anchor="w", pady=(2, 12))

        # Probeer activiteiten uit Supabase te halen
        log_items = []
        try:
            res = sc.get_client().table("activity_log") \
                .select("user_email, action_type, aantal, created_at") \
                .order("created_at", desc=True).limit(8).execute()
            log_items = res.data or []
        except Exception:
            pass

        if log_items:
            for entry in log_items:
                action = entry.get("action_type", "")
                email  = entry.get("user_email", "onbekend")
                aantal = entry.get("aantal", 1)
                ts     = entry.get("created_at", "")[:16].replace("T", " ")
                color  = GREEN if action == "label" else (ORANGE if action == "kast_bon" else BLUE)
                icon   = ICON_LABEL if action == "label" else (ICON_KAST if action == "kast_bon" else ICON_ASSISTENTIE)
                label  = f"{aantal}× Label" if action == "label" else ("123kast Bon" if action == "kast_bon" else "Assistentie Bon")

                row = tk.Frame(inner, bg=BG_CARD,
                               highlightbackground=BORDER, highlightthickness=1)
                row.pack(fill=tk.X, pady=(0, 6))
                tk.Frame(row, bg=color, width=4).pack(side=tk.LEFT, fill=tk.Y)
                info = tk.Frame(row, bg=BG_CARD)
                info.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10, pady=6)
                tk.Label(info, text=f"{icon}  {label}", bg=BG_CARD, fg=TEXT_DARK,
                         font=(FONT, 9, "bold"), anchor="w").pack(anchor="w")
                tk.Label(info, text=f"{email}  •  {ts}", bg=BG_CARD, fg=TEXT_MUTED,
                         font=(FONT, 8), anchor="w").pack(anchor="w")
        else:
            # Fallback: toon laatste datums uit stats
            for icon, lbl, datum, color in [
                (ICON_LABEL,       "Label gegenereerd",       laatste_label, GREEN),
                (ICON_ASSISTENTIE, "Assistentie Bon gemaakt", laatste_bon,   BLUE),
                (ICON_KAST,        "123kast Bon gemaakt",     laatste_kast,  ORANGE),
            ]:
                row = tk.Frame(inner, bg=BG_CARD,
                               highlightbackground=BORDER, highlightthickness=1)
                row.pack(fill=tk.X, pady=(0, 8))
                tk.Frame(row, bg=color, width=4).pack(side=tk.LEFT, fill=tk.Y)
                info = tk.Frame(row, bg=BG_CARD)
                info.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10, pady=8)
                tk.Label(info, text=f"{icon}  {lbl}", bg=BG_CARD, fg=TEXT_DARK,
                         font=(FONT, 9, "bold"), anchor="w").pack(anchor="w")
                tk.Label(info, text=datum, bg=BG_CARD, fg=TEXT_MUTED,
                         font=(FONT, 8), anchor="w").pack(anchor="w")

    def _dark_summary_card(self, parent, col):
        """Donkere samenvattings-kaart met witte tekst en progress-bar."""
        card = tk.Frame(parent, bg=BG_DARK_CARD,
                        highlightbackground=BORDER, highlightthickness=1)
        card.grid(row=0, column=col, sticky="nsew")

        inner = tk.Frame(card, bg=BG_DARK_CARD)
        inner.pack(fill=tk.BOTH, expand=True, padx=18, pady=16)

        tk.Label(inner, text="Voortgang", bg=BG_DARK_CARD, fg="#8A8FA8",
                 font=(FONT, 9), anchor="w").pack(anchor="w")
        tk.Label(inner, text="4751", bg=BG_DARK_CARD, fg=TEXT_LIGHT,
                 font=(FONT, 28, "bold"), anchor="w").pack(anchor="w", pady=(4, 2))
        tk.Label(inner, text="PLAN", bg=GREEN, fg=TEXT_LIGHT,
                 font=(FONT, 7, "bold"),
                 padx=6).pack(anchor="w", pady=(0, 10))

        # Progress-balk
        tk.Label(inner, text="Maandelijks doel: 78%",
                 bg=BG_DARK_CARD, fg="#8A8FA8",
                 font=(FONT, 8), anchor="w").pack(anchor="w", pady=(4, 4))

        track = tk.Frame(inner, bg="#2E2E4E", height=8)
        track.pack(fill=tk.X)
        track.pack_propagate(False)

        bar_cv = tk.Canvas(track, bg="#2E2E4E", height=8, highlightthickness=0)
        bar_cv.pack(fill=tk.BOTH, expand=True)

        def draw(event=None, c=bar_cv):
            w = c.winfo_width() or 160
            c.delete("all")
            c.create_rectangle(0, 0, int(w * 0.78), 8, fill=GREEN, outline="")
        bar_cv.bind("<Configure>", draw)
        bar_cv.after(120, draw)

    def _chart_card(self, parent, col):
        card, inner = make_card(parent, accent_color=GREEN, padx=18, pady=16)
        card.grid(row=0, column=col, sticky="nsew", padx=(0, 12))

        tk.Label(inner, text="Activiteit overzicht", bg=BG_CARD, fg=TEXT_MUTED,
                 font=(FONT, 9), anchor="w").pack(anchor="w")
        tk.Label(inner, text="Afgelopen 7 dagen", bg=BG_CARD, fg=TEXT_DARK,
                 font=(FONT, 12, "bold"), anchor="w").pack(anchor="w", pady=(2, 10))

        chart = tk.Canvas(inner, bg=BG_CARD, height=80, highlightthickness=0)
        chart.pack(fill=tk.X)

        pts_y = [65, 48, 58, 28, 42, 22, 35]

        def draw_chart(event=None, c=chart, pts=pts_y):
            w = c.winfo_width() or 300
            step = w / (len(pts) - 1)
            coords = []
            for i, y in enumerate(pts):
                coords.extend([i * step, y])
            c.delete("all")
            c.create_line(*coords, fill="#C8E6C9", width=5, smooth=True)
            c.create_line(*coords, fill=GREEN, width=2, smooth=True)
            for i, y in enumerate(pts):
                x = i * step
                c.create_oval(x-4, y-4, x+4, y+4,
                              fill=GREEN, outline=BG_CARD, width=2)

        chart.bind("<Configure>", draw_chart)
        chart.after(120, draw_chart)

    def _progress_card(self, parent, labels_totaal, bons_totaal, col):
        """Progress-kaart met echte verhoudingen op basis van opgeslagen data."""
        card, inner = make_card(parent, accent_color=BLUE, padx=18, pady=16)
        card.grid(row=0, column=col, sticky="nsew")

        tk.Label(inner, text="Module gebruik", bg=BG_CARD, fg=TEXT_MUTED,
                 font=(FONT, 9), anchor="w").pack(anchor="w")
        tk.Label(inner, text="Verdeling", bg=BG_CARD, fg=TEXT_DARK,
                 font=(FONT, 12, "bold"), anchor="w").pack(anchor="w", pady=(2, 10))

        stats = self.load_stats()
        kastbon_totaal = stats.get("totaal_kast_bonnen", 0)

        totaal = labels_totaal + bons_totaal + kastbon_totaal
        if totaal > 0:
            label_pct   = int(labels_totaal  / totaal * 100)
            bon_pct     = int(bons_totaal    / totaal * 100)
            kastbon_pct = int(kastbon_totaal / totaal * 100)
        else:
            label_pct   = 34
            bon_pct     = 33
            kastbon_pct = 33

        cats = [
            ("Label Maker",     label_pct,   GREEN),
            ("Assistentie Bon", bon_pct,     BLUE),
            ("123kast Bon",     kastbon_pct, ORANGE),
        ]
        for name, pct, color in cats:
            row = tk.Frame(inner, bg=BG_CARD)
            row.pack(fill=tk.X, pady=4)

            tk.Label(row, text=name, bg=BG_CARD, fg=TEXT_DARK,
                     font=(FONT, 9), width=14, anchor="w").pack(side=tk.LEFT)
            tk.Label(row, text=f"{pct}%", bg=BG_CARD, fg=TEXT_MUTED,
                     font=(FONT, 8), width=4, anchor="e").pack(side=tk.RIGHT)

            track = tk.Frame(row, bg=BORDER, height=7)
            track.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 8))
            track.pack_propagate(False)

            bc = tk.Canvas(track, bg=BORDER, height=7, highlightthickness=0)
            bc.pack(fill=tk.BOTH, expand=True)

            def draw_bar(event=None, c=bc, p=pct, clr=color):
                w = c.winfo_width() or 180
                c.delete("all")
                c.create_rectangle(0, 0, int(w * p / 100), 7,
                                   fill=clr, outline="")
            bc.bind("<Configure>", draw_bar)
            bc.after(150, draw_bar)

    # ════════════════════════════════════════════════════════════════════════
    # ADMIN PAGINA
    # ════════════════════════════════════════════════════════════════════════

    def show_admin_page(self):
        if not sc.check_admin_role():
            messagebox.showerror("Geen toegang", "Je hebt geen beheerdersrechten.")
            self._go_dashboard()
            return
        self.page_title_var.set("Admin")
        self._clear_content()
        page = self._scrollable()

        # ── Gebruikers statistieken uit activity_log ─────────────────────────
        user_stats = {}
        log_rows   = []
        try:
            res = sc.get_client().table("activity_log") \
                .select("user_email, action_type, aantal, created_at") \
                .order("created_at", desc=True).limit(200).execute()
            for r in (res.data or []):
                e = r["user_email"]
                if e not in user_stats:
                    user_stats[e] = {"labels": 0, "bons": 0, "kastbons": 0, "laatste": r["created_at"][:16].replace("T", " ")}
                if r["action_type"] == "label":
                    user_stats[e]["labels"]   += r.get("aantal", 0)
                elif r["action_type"] == "kast_bon":
                    user_stats[e]["kastbons"] += r.get("aantal", 1)
                else:
                    user_stats[e]["bons"]     += r.get("aantal", 1)
            log_rows = (res.data or [])[:50]
        except Exception:
            pass

        # ── Rij A: Overzichtskaarten ──────────────────────────────────────────
        top_row = tk.Frame(page, bg=BG_ROOT)
        top_row.pack(fill=tk.X, pady=(0, 16))
        for i in range(4):
            top_row.columnconfigure(i, weight=1)

        stats_global = self.load_stats()
        self._stat_card(top_row, "Actieve gebruikers",
                        str(len(user_stats)), "uit activity log", PURPLE, col=0)
        self._stat_card(top_row, "Totaal labels",
                        str(stats_global["totaal_labels_gegenereerd"]), "alle gebruikers", GREEN, col=1)
        self._stat_card(top_row, "Totaal bons",
                        str(stats_global["totaal_assistentie_bonnen"]), "alle gebruikers", BLUE, col=2)
        self._stat_card(top_row, "Totaal 123kast Bons",
                        str(stats_global["totaal_kast_bonnen"]), "alle gebruikers", ORANGE, col=3)

        # ── Rij B: Per-gebruiker tabel ────────────────────────────────────────
        tbl_card, tbl_inner = make_card(page, accent_color=PURPLE, padx=18, pady=16)
        tbl_card.pack(fill=tk.X, pady=(0, 16))
        tk.Label(tbl_inner, text="Activiteit per gebruiker", bg=BG_CARD, fg=TEXT_MUTED,
                 font=(FONT, 9), anchor="w").pack(anchor="w")
        tk.Label(tbl_inner, text="Overzicht", bg=BG_CARD, fg=TEXT_DARK,
                 font=(FONT, 12, "bold"), anchor="w").pack(anchor="w", pady=(2, 10))

        # Tabelheader
        hdr = tk.Frame(tbl_inner, bg=ACCENT)
        hdr.pack(fill=tk.X)
        for txt, w in [("Gebruiker", 220), ("Labels", 70), ("Bons", 70), ("123kast", 70), ("Laatste actie", 150)]:
            tk.Label(hdr, text=txt, bg=ACCENT, fg=TEXT_LIGHT,
                     font=(FONT, 9, "bold"), width=w//8, anchor="w",
                     padx=8, pady=5).pack(side=tk.LEFT)

        if user_stats:
            for i, (email, d) in enumerate(sorted(user_stats.items())):
                row_bg = BG_CARD if i % 2 == 0 else "#F7F9FC"
                row = tk.Frame(tbl_inner, bg=row_bg, highlightbackground=BORDER, highlightthickness=1)
                row.pack(fill=tk.X)
                for txt, w in [(email, 220), (str(d["labels"]), 70),
                               (str(d["bons"]), 70), (str(d.get("kastbons", 0)), 70),
                               (d["laatste"], 150)]:
                    tk.Label(row, text=txt, bg=row_bg, fg=TEXT_DARK,
                             font=(FONT, 9), width=w//8, anchor="w",
                             padx=8, pady=5).pack(side=tk.LEFT)
        else:
            tk.Label(tbl_inner, text="Geen data beschikbaar (activity_log tabel leeg of niet aangemaakt).",
                     bg=BG_CARD, fg=TEXT_MUTED, font=(FONT, 9)).pack(anchor="w", pady=8)

        # ── Rij C: Log + Beheer ───────────────────────────────────────────────
        bottom = tk.Frame(page, bg=BG_ROOT)
        bottom.pack(fill=tk.X, pady=(0, 16))
        bottom.columnconfigure(0, weight=3)
        bottom.columnconfigure(1, weight=2)

        # Recente activiteiten log
        log_card, log_inner = make_card(bottom, accent_color=GREEN, padx=18, pady=16)
        log_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        tk.Label(log_inner, text="Recente activiteiten", bg=BG_CARD, fg=TEXT_MUTED,
                 font=(FONT, 9), anchor="w").pack(anchor="w")
        tk.Label(log_inner, text="Laatste 50 acties", bg=BG_CARD, fg=TEXT_DARK,
                 font=(FONT, 12, "bold"), anchor="w").pack(anchor="w", pady=(2, 10))

        for entry in log_rows:
            action = entry.get("action_type", "")
            email  = entry.get("user_email", "?")
            aantal = entry.get("aantal", 1)
            ts     = entry.get("created_at", "")[:16].replace("T", " ")
            color  = GREEN if action == "label" else (ORANGE if action == "kast_bon" else BLUE)
            icon   = ICON_LABEL if action == "label" else (ICON_KAST if action == "kast_bon" else ICON_ASSISTENTIE)
            lbl    = f"{aantal}× Label" if action == "label" else ("123kast Bon" if action == "kast_bon" else "Assistentie Bon")

            r = tk.Frame(log_inner, bg=BG_CARD, highlightbackground=BORDER, highlightthickness=1)
            r.pack(fill=tk.X, pady=(0, 4))
            tk.Frame(r, bg=color, width=4).pack(side=tk.LEFT, fill=tk.Y)
            info = tk.Frame(r, bg=BG_CARD)
            info.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8, pady=4)
            tk.Label(info, text=f"{icon}  {lbl}  —  {email}", bg=BG_CARD, fg=TEXT_DARK,
                     font=(FONT, 9), anchor="w").pack(anchor="w")
            tk.Label(info, text=ts, bg=BG_CARD, fg=TEXT_MUTED,
                     font=(FONT, 8), anchor="w").pack(anchor="w")

        if not log_rows:
            tk.Label(log_inner, text="Geen activiteiten gevonden.",
                     bg=BG_CARD, fg=TEXT_MUTED, font=(FONT, 9)).pack(anchor="w")

        # Beheer kaart
        mgmt_card, mgmt_inner = make_card(bottom, accent_color=ORANGE, padx=18, pady=16)
        mgmt_card.grid(row=0, column=1, sticky="nsew")
        tk.Label(mgmt_inner, text="Beheer", bg=BG_CARD, fg=TEXT_MUTED,
                 font=(FONT, 9), anchor="w").pack(anchor="w")
        tk.Label(mgmt_inner, text="Acties", bg=BG_CARD, fg=TEXT_DARK,
                 font=(FONT, 12, "bold"), anchor="w").pack(anchor="w", pady=(2, 16))

        btn_cfg = dict(bd=0, relief=tk.FLAT, pady=9, cursor="hand2",
                       font=(FONT, 10), anchor="w", padx=14)

        tk.Button(mgmt_inner, text="📥  Exporteer CSV",
                  bg="#E8F5E9", fg=TEXT_DARK, activebackground=BORDER,
                  command=lambda: self._admin_export_csv(log_rows),
                  **btn_cfg).pack(fill=tk.X, pady=(0, 8))

        tk.Button(mgmt_inner, text="🔄  Herlaad pagina",
                  bg="#E3F2FD", fg=TEXT_DARK, activebackground=BORDER,
                  command=self.show_admin_page,
                  **btn_cfg).pack(fill=tk.X, pady=(0, 8))

        tk.Frame(mgmt_inner, bg=BORDER, height=1).pack(fill=tk.X, pady=8)

        tk.Button(mgmt_inner, text="⚠  Reset statistieken",
                  bg="#FFEBEE", fg="#C62828", activebackground="#FFCDD2",
                  command=self._admin_reset_stats,
                  **btn_cfg).pack(fill=tk.X)

        # ── Rij D: Gebruikersbeheer ───────────────────────────────────────────
        usr_card, usr_inner = make_card(page, accent_color=PURPLE, padx=18, pady=16)
        usr_card.pack(fill=tk.X, pady=(0, 16))

        hdr_row = tk.Frame(usr_inner, bg=BG_CARD)
        hdr_row.pack(fill=tk.X, pady=(0, 12))
        tk.Label(hdr_row, text="Gebruikersbeheer", bg=BG_CARD, fg=TEXT_DARK,
                 font=(FONT, 12, "bold"), anchor="w").pack(side=tk.LEFT)
        tk.Button(hdr_row, text="+ Nieuw account",
                  command=self._admin_new_user_dialog,
                  bg=ACCENT, fg=TEXT_LIGHT, activebackground="#2E2E4E",
                  font=(FONT, 9), bd=0, relief=tk.FLAT,
                  padx=10, pady=4, cursor="hand2").pack(side=tk.RIGHT)

        # Gebruikerslijst ophalen
        users = sc.admin_list_users()
        roles_map = {}
        try:
            res = sc.get_admin_client().table("user_roles").select("user_email, is_admin").execute()
            roles_map = {r["user_email"]: r["is_admin"] for r in (res.data or [])}
        except Exception:
            pass

        if users:
            # Tabelheader
            col_hdr = tk.Frame(usr_inner, bg=ACCENT)
            col_hdr.pack(fill=tk.X)
            for txt, w in [("E-mailadres", 300), ("Admin", 80), ("Acties", 200)]:
                tk.Label(col_hdr, text=txt, bg=ACCENT, fg=TEXT_LIGHT,
                         font=(FONT, 9, "bold"), width=w//8, anchor="w",
                         padx=8, pady=5).pack(side=tk.LEFT)

            for i, user in enumerate(users):
                uid   = user.id
                email = user.email or "—"
                is_ad = roles_map.get(email, False)
                row_bg = BG_CARD if i % 2 == 0 else "#F7F9FC"
                row = tk.Frame(usr_inner, bg=row_bg,
                               highlightbackground=BORDER, highlightthickness=1)
                row.pack(fill=tk.X)

                tk.Label(row, text=email, bg=row_bg, fg=TEXT_DARK,
                         font=(FONT, 9), width=300//8, anchor="w",
                         padx=8, pady=6).pack(side=tk.LEFT)
                tk.Label(row, text="✓" if is_ad else "—", bg=row_bg,
                         fg=GREEN if is_ad else TEXT_MUTED,
                         font=(FONT, 9, "bold"), width=80//8, anchor="w",
                         padx=8).pack(side=tk.LEFT)

                act = tk.Frame(row, bg=row_bg)
                act.pack(side=tk.LEFT, padx=8)

                toggle_txt = "Admin verwijderen" if is_ad else "Maak admin"
                toggle_fg  = "#C62828" if is_ad else GREEN
                tk.Button(act, text=toggle_txt, fg=toggle_fg, bg=row_bg,
                          activebackground=BORDER, bd=0, relief=tk.FLAT,
                          font=(FONT, 8, "underline"), cursor="hand2",
                          command=lambda e=email, a=is_ad: self._admin_toggle_role(e, a)
                          ).pack(side=tk.LEFT, padx=(0, 8))

                tk.Button(act, text="Reset wachtwoord", fg="#1565C0", bg=row_bg,
                          activebackground=BORDER, bd=0, relief=tk.FLAT,
                          font=(FONT, 8, "underline"), cursor="hand2",
                          command=lambda e=email: self._admin_reset_user_pw(e)
                          ).pack(side=tk.LEFT, padx=(0, 8))

                if email != sc.get_user_email():
                    tk.Button(act, text="Verwijder", fg="#C62828", bg=row_bg,
                              activebackground=BORDER, bd=0, relief=tk.FLAT,
                              font=(FONT, 8, "underline"), cursor="hand2",
                              command=lambda uid=uid, e=email: self._admin_delete_user(uid, e)
                              ).pack(side=tk.LEFT)
        else:
            tk.Label(usr_inner, text="Geen gebruikers gevonden.",
                     bg=BG_CARD, fg=TEXT_MUTED, font=(FONT, 9)).pack(anchor="w")

    def _admin_new_user_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Nieuw account aanmaken")
        dlg.configure(bg="#FFFFFF")
        dlg.geometry("380x300")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        pad = tk.Frame(dlg, bg="#FFFFFF", padx=24, pady=20)
        pad.pack(fill=tk.BOTH, expand=True)

        tk.Label(pad, text="E-mailadres", bg="#FFFFFF", fg=TEXT_MUTED,
                 font=(FONT, 9), anchor="w").pack(fill=tk.X, pady=(0, 2))
        email_var = tk.StringVar()
        tk.Entry(pad, textvariable=email_var, bg="#F7F9FC", fg=TEXT_DARK,
                 relief=tk.FLAT, highlightbackground=BORDER, highlightthickness=1,
                 font=(FONT, 10)).pack(fill=tk.X, ipady=6)

        tk.Label(pad, text="Wachtwoord", bg="#FFFFFF", fg=TEXT_MUTED,
                 font=(FONT, 9), anchor="w").pack(fill=tk.X, pady=(12, 2))
        pass_var = tk.StringVar()
        tk.Entry(pad, textvariable=pass_var, show="•", bg="#F7F9FC", fg=TEXT_DARK,
                 relief=tk.FLAT, highlightbackground=BORDER, highlightthickness=1,
                 font=(FONT, 10)).pack(fill=tk.X, ipady=6)

        admin_var = tk.BooleanVar(value=False)
        tk.Checkbutton(pad, text="Admin-rechten geven", variable=admin_var,
                       bg="#FFFFFF", fg=TEXT_DARK, font=(FONT, 9),
                       activebackground="#FFFFFF").pack(anchor="w", pady=(12, 0))

        err_var = tk.StringVar()
        tk.Label(pad, textvariable=err_var, bg="#FFFFFF", fg="#E53935",
                 font=(FONT, 8), anchor="w", wraplength=330).pack(fill=tk.X, pady=(6, 0))

        def _do_create():
            import re
            email = email_var.get().strip()
            password = pass_var.get()
            if not email or not password:
                err_var.set("Vul e-mailadres en wachtwoord in.")
                return
            if not re.match(r'^[^@\s]+@[^@\s]+\.[^@\s]+$', email):
                err_var.set("Ongeldig e-mailadres.")
                return
            try:
                sc.admin_create_user(email, password)
                if admin_var.get():
                    sc.admin_set_admin_role(email, True)
                dlg.destroy()
                self.show_admin_page()
            except Exception as e:
                err_var.set(f"Fout: {e}")

        tk.Button(pad, text="Account aanmaken", command=_do_create,
                  bg=ACCENT, fg="#FFFFFF", activebackground="#2E2E4E",
                  activeforeground="#FFFFFF", font=(FONT, 10, "bold"),
                  bd=0, relief=tk.FLAT, pady=8, cursor="hand2").pack(fill=tk.X, pady=(12, 0))

    def _admin_toggle_role(self, email, current_is_admin):
        try:
            sc.admin_set_admin_role(email, not current_is_admin)
            self.show_admin_page()
        except Exception as e:
            messagebox.showerror("Fout", str(e))

    def _admin_delete_user(self, user_id, email):
        if not sc.check_admin_role():
            messagebox.showerror("Geen toegang", "Je hebt geen beheerdersrechten.")
            return
        if not messagebox.askyesno(
            "Gebruiker verwijderen",
            f"Weet je zeker dat je '{email}' wilt verwijderen?\nDit kan niet ongedaan worden gemaakt.",
            icon="warning"
        ):
            return
        try:
            sc.admin_delete_user(user_id)
            try:
                sc.get_admin_client().table("user_roles").delete().eq("user_email", email).execute()
            except Exception:
                pass
            self.show_admin_page()
        except Exception as e:
            messagebox.showerror("Fout", str(e))

    def _admin_reset_user_pw(self, email):
        if not messagebox.askyesno("Reset wachtwoord",
                                   f"Reset-link sturen naar:\n{email}?"):
            return
        try:
            sc.reset_password(email)
            messagebox.showinfo("Verstuurd", f"Reset-link verstuurd naar {email}.")
        except Exception as e:
            messagebox.showerror("Fout", str(e))

    def _admin_export_csv(self, rows):
        if not sc.check_admin_role():
            messagebox.showerror("Geen toegang", "Je hebt geen beheerdersrechten.")
            return
        import csv
        from tkinter import filedialog
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV bestand", "*.csv")],
            initialfile="activiteiten_export.csv",
            title="Sla export op")
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.DictWriter(f, fieldnames=["user_email", "action_type", "aantal", "created_at"])
                w.writeheader()
                w.writerows(rows)
            messagebox.showinfo("Succes", f"Geëxporteerd naar:\n{path}")
        except Exception as e:
            messagebox.showerror("Fout", str(e))

    def _admin_reset_stats(self):
        if not sc.check_admin_role():
            messagebox.showerror("Geen toegang", "Je hebt geen beheerdersrechten.")
            return
        if not messagebox.askyesno("Reset statistieken",
                                   "Weet je zeker dat je alle statistieken wilt resetten?\n"
                                   "Dit verwijdert ook de volledige activiteiten log.",
                                   icon="warning"):
            return
        try:
            sc.get_client().table("app_stats").update({
                "totaal_labels_gegenereerd": 0,
                "totaal_assistentie_bonnen": 0,
                "laatste_label_datum": "",
                "laatste_bon_datum": "",
            }).eq("id", 1).execute()
            sc.get_client().table("activity_log").delete().neq("id", 0).execute()
            messagebox.showinfo("Gereset", "Statistieken en activiteiten log zijn gereset.")
            self.show_admin_page()
        except Exception as e:
            messagebox.showerror("Fout", str(e))

    # ════════════════════════════════════════════════════════════════════════
    # PUBLIEKE NAVIGATIE (signaturen ongewijzigd)
    # ════════════════════════════════════════════════════════════════════════

    def show_main_menu(self):
        self._nav_click("Dashboard", self._go_dashboard)

    def show_label_maker(self):
        self.page_title_var.set("Label Maker")
        self._clear_content()
        # Wrap in scrollbare kaart
        wrap = tk.Frame(self.content_frame, bg=BG_ROOT)
        wrap.pack(fill=tk.BOTH, expand=True)

        cv = tk.Canvas(wrap, bg=BG_ROOT, highlightthickness=0)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=cv.yview)
        page = tk.Frame(cv, bg=BG_ROOT)
        page.bind("<Configure>",
                  lambda e: cv.configure(scrollregion=cv.bbox("all")))
        lm_win_id = cv.create_window((0, 0), window=page, anchor="nw")
        cv.configure(yscrollcommand=sb.set)
        cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        cv.bind("<Configure>", lambda e: cv.itemconfig(lm_win_id, width=e.width))
        cv.bind_all("<MouseWheel>",
                    lambda e: cv.yview_scroll(int(-1*(e.delta/120)), "units"))

        # Kaart-wrapper voor het formulier
        form_card = tk.Frame(page, bg=BG_CARD,
                             highlightbackground=BORDER, highlightthickness=1)
        form_card.pack(fill=tk.X, pady=(0, 16))
        tk.Frame(form_card, bg=GREEN, height=4).pack(fill=tk.X)

        self.label_maker = LabelMakerComponent(form_card, self)

    def show_assistentie_bon(self):
        self.page_title_var.set("Assistentie Bon")
        self._clear_content()
        wrap = tk.Frame(self.content_frame, bg=BG_ROOT)
        wrap.pack(fill=tk.BOTH, expand=True)

        cv = tk.Canvas(wrap, bg=BG_ROOT, highlightthickness=0)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=cv.yview)
        page = tk.Frame(cv, bg=BG_ROOT)
        page.bind("<Configure>",
                  lambda e: cv.configure(scrollregion=cv.bbox("all")))
        ab_win_id = cv.create_window((0, 0), window=page, anchor="nw")
        cv.configure(yscrollcommand=sb.set)
        cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        cv.bind("<Configure>", lambda e: cv.itemconfig(ab_win_id, width=e.width))
        cv.bind_all("<MouseWheel>",
                    lambda e: cv.yview_scroll(int(-1*(e.delta/120)), "units"))

        form_card = tk.Frame(page, bg=BG_CARD,
                             highlightbackground=BORDER, highlightthickness=1)
        form_card.pack(fill=tk.X, pady=(0, 16))
        tk.Frame(form_card, bg=BLUE, height=4).pack(fill=tk.X)

        self.assistentie_bon = AssistentieBonComponent(form_card, self)

    def show_kast_bon(self):
        self.page_title_var.set("123kast Bon")
        self._clear_content()
        wrap = tk.Frame(self.content_frame, bg=BG_ROOT)
        wrap.pack(fill=tk.BOTH, expand=True)

        cv = tk.Canvas(wrap, bg=BG_ROOT, highlightthickness=0)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=cv.yview)
        page = tk.Frame(cv, bg=BG_ROOT)
        page.bind("<Configure>",
                  lambda e: cv.configure(scrollregion=cv.bbox("all")))
        kb_win_id = cv.create_window((0, 0), window=page, anchor="nw")
        cv.configure(yscrollcommand=sb.set)
        cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        cv.bind("<Configure>", lambda e: cv.itemconfig(kb_win_id, width=e.width))
        cv.bind_all("<MouseWheel>",
                    lambda e: cv.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        form_card = tk.Frame(page, bg=BG_CARD,
                             highlightbackground=BORDER, highlightthickness=1)
        form_card.pack(fill=tk.X, pady=(0, 16))
        tk.Frame(form_card, bg="#FF6B35", height=4).pack(fill=tk.X)

        self.kast_bon = KastBonComponent(form_card, self)


def _check_for_update(root: tk.Tk):
    """Controleer op de achtergrond op updates en toon popup bij nieuwe versie."""
    def _worker():
        try:
            url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            req = urllib.request.Request(url, headers={"User-Agent": "Excellent-App"})
            with urllib.request.urlopen(req, timeout=5) as resp:
                data = json.loads(resp.read())
            latest = data.get("tag_name", "")
            if not latest or latest == VERSION or VERSION == "dev":
                return
            assets = data.get("assets", [])
            asset_url = next(
                (a["browser_download_url"] for a in assets
                 if a["name"].endswith(".exe")), None)
            if not asset_url:
                return
            sha256_url = next(
                (a["browser_download_url"] for a in assets
                 if a["name"].endswith(".sha256")), None)
            root.after(0, lambda: _show_update_popup(root, latest, asset_url, sha256_url))
        except Exception:
            pass
    threading.Thread(target=_worker, daemon=True).start()


def _show_update_popup(root: tk.Tk, latest: str, asset_url: str, sha256_url=None):
    import subprocess, tempfile, hashlib

    overlay = tk.Frame(root, bg="#000000")
    overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
    overlay.configure(cursor="arrow")

    card = tk.Frame(overlay, bg="#FFFFFF",
                    highlightbackground="#E8EAF0", highlightthickness=1)
    card.place(relx=0.5, rely=0.5, anchor="center", width=420, height=270)

    header = tk.Frame(card, bg=ACCENT, height=56)
    header.pack(fill=tk.X)
    header.pack_propagate(False)
    tk.Label(header, text="🔄  Update beschikbaar",
             bg=ACCENT, fg="#FFFFFF", font=(FONT, 13, "bold")).pack(expand=True)

    body = tk.Frame(card, bg="#FFFFFF", padx=28, pady=18)
    body.pack(fill=tk.BOTH, expand=True)

    tk.Label(body, text="Er is een nieuwe versie van Excellent App beschikbaar.",
             bg="#FFFFFF", fg="#1A1A2E", font=(FONT, 10)).pack(anchor="w")

    ver_row = tk.Frame(body, bg="#FFFFFF")
    ver_row.pack(anchor="w", pady=(10, 0))
    tk.Label(ver_row, text="Huidige versie:", bg="#FFFFFF", fg="#8A8FA8",
             font=(FONT, 9)).pack(side=tk.LEFT)
    tk.Label(ver_row, text=f"  {VERSION}", bg="#FFFFFF", fg="#1A1A2E",
             font=(FONT, 9, "bold")).pack(side=tk.LEFT)

    ver_row2 = tk.Frame(body, bg="#FFFFFF")
    ver_row2.pack(anchor="w", pady=(3, 0))
    tk.Label(ver_row2, text="Nieuwe versie: ", bg="#FFFFFF", fg="#8A8FA8",
             font=(FONT, 9)).pack(side=tk.LEFT)
    tk.Label(ver_row2, text=latest, bg="#FFFFFF", fg="#4CAF50",
             font=(FONT, 9, "bold")).pack(side=tk.LEFT)

    prog_track = tk.Frame(body, bg="#E8EAF0", height=6)
    prog_track.pack(fill=tk.X, pady=(14, 0))
    prog_track.pack_propagate(False)
    prog_bar = tk.Frame(prog_track, bg=GREEN, width=0, height=6)
    prog_bar.place(x=0, y=0, relheight=1)

    status_var = tk.StringVar(value="Klik op 'Update installeren' om te starten.")
    tk.Label(body, textvariable=status_var, bg="#FFFFFF", fg="#8A8FA8",
             font=(FONT, 8)).pack(anchor="w", pady=(6, 0))

    def _set_progress(pct):
        try:
            w = prog_track.winfo_width()
            prog_bar.place(x=0, y=0, relheight=1, width=int(w * pct / 100))
        except Exception:
            pass

    def _show_error(msg):
        status_var.set(msg)
        update_btn.config(text="Opnieuw proberen", state=tk.NORMAL)

    def _do_update():
        update_btn.config(text="Bezig...", state=tk.DISABLED)
        status_var.set("Download wordt gestart...")

        def _download():
            try:
                tmp = os.path.join(tempfile.gettempdir(), "ExcellentApp_update.exe")

                req = urllib.request.urlopen(
                    urllib.request.Request(asset_url, headers={"User-Agent": "Excellent-App"}),
                    timeout=120)
                total = int(req.headers.get("Content-Length", 0))
                downloaded = 0
                h = hashlib.sha256()
                with open(tmp, "wb") as f:
                    while True:
                        chunk = req.read(65536)
                        if not chunk:
                            break
                        f.write(chunk)
                        h.update(chunk)
                        downloaded += len(chunk)
                        if total > 0:
                            pct = int(downloaded / total * 100)
                            root.after(0, lambda p=pct: (
                                status_var.set(f"Downloaden... {p}%"),
                                _set_progress(p)))

                if sha256_url:
                    try:
                        with urllib.request.urlopen(sha256_url, timeout=10) as resp:
                            expected = resp.read().decode().strip().split()[0].lower()
                        if h.hexdigest().lower() != expected:
                            try:
                                os.remove(tmp)
                            except Exception:
                                pass
                            root.after(0, lambda: _show_error(
                                "Hashverificatie mislukt — update geannuleerd."))
                            return
                    except Exception:
                        pass

                if not getattr(sys, "frozen", False):
                    root.after(0, lambda: _show_error(
                        "Update werkt alleen vanuit de .exe versie."))
                    return

                current_exe = sys.executable
                exe_name = os.path.basename(current_exe)
                bat = os.path.join(tempfile.gettempdir(), "excellent_update.bat")

                with open(bat, "w") as f:
                    f.write(
                        "@echo off\n"
                        # Wacht tot app zichzelf sluit (1.5s) + extra marge
                        "timeout /t 5 /nobreak > nul\n"
                        # Forceer afsluiten voor de zekerheid
                        f'taskkill /f /im "{exe_name}" > nul 2>&1\n'
                        "timeout /t 2 /nobreak > nul\n"
                        # Kopieer met retry (max 15 pogingen)
                        "set COUNT=0\n"
                        ":RETRY\n"
                        "set /a COUNT=COUNT+1\n"
                        "if %COUNT% gtr 15 goto FAIL\n"
                        f'copy /Y "{tmp}" "{current_exe}" > nul 2>&1\n'
                        "if errorlevel 1 (\n"
                        "    timeout /t 2 /nobreak > nul\n"
                        "    goto RETRY\n"
                        ")\n"
                        # Wacht even voor schijf flush
                        "timeout /t 2 /nobreak > nul\n"
                        f'start "" "{current_exe}"\n'
                        f'del /f /q "{tmp}" > nul 2>&1\n'
                        'del /f /q "%~f0"\n'
                        "exit /b 0\n"
                        ":FAIL\n"
                        f'del /f /q "{tmp}" > nul 2>&1\n'
                        'del /f /q "%~f0"\n'
                        "exit /b 1\n"
                    )

                def _apply():
                    try:
                        status_var.set("Installeren... app start zo opnieuw op.")
                        _set_progress(100)
                        overlay.update()
                        # CREATE_NEW_PROCESS_GROUP zorgt dat het proces de parent overleeft
                        # DEVNULL handles zijn verplicht voor windowed apps zonder console
                        subprocess.Popen(
                            ["cmd", "/c", bat],
                            creationflags=(subprocess.CREATE_NEW_PROCESS_GROUP |
                                           subprocess.CREATE_NO_WINDOW),
                            stdin=subprocess.DEVNULL,
                            stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL,
                            close_fds=True,
                        )
                        # Geef het bat-script 1.5s om te starten voor we sluiten
                        root.after(1500, root.destroy)
                    except Exception as e:
                        root.after(0, lambda err=str(e): _show_error(
                            f"Kon update niet starten: {err}"))

                root.after(0, _apply)

            except Exception as e:
                root.after(0, lambda err=str(e): _show_error(f"Download mislukt: {err}"))

        threading.Thread(target=_download, daemon=True).start()

    btn_row = tk.Frame(body, bg="#FFFFFF")
    btn_row.pack(fill=tk.X, pady=(10, 0))

    update_btn = tk.Button(btn_row, text="Update installeren", command=_do_update,
                           bg=ACCENT, fg="#FFFFFF",
                           activebackground="#2E2E4E", activeforeground="#FFFFFF",
                           font=(FONT, 9, "bold"), bd=0, relief=tk.FLAT,
                           pady=9, cursor="hand2")
    update_btn.pack(fill=tk.X)


def _launch_app():
    sc.check_admin_role()
    main_root = tk.Tk()
    ExcellentApp(main_root)
    _check_for_update(main_root)
    main_root.mainloop()


def _start_login():
    root = tk.Tk()

    def on_login(_email):
        root.destroy()
        _launch_app()

    LoginWindow(root, on_success=on_login)
    root.mainloop()


if __name__ == "__main__":
    if sc.restore_session():
        _launch_app()
    else:
        _start_login()
