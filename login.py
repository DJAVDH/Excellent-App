import tkinter as tk
from tkinter import ttk
import os

BG       = "#FFFFFF"
BG_FIELD = "#F7F9FC"
ACCENT   = "#1A1A2E"
GREEN    = "#4CAF50"
TEXT     = "#1A1A2E"
MUTED    = "#8A8FA8"
BORDER   = "#E8EAF0"
ERROR    = "#E53935"
FONT     = "Segoe UI"

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


class LoginWindow:
    """Loginscherm — roept on_success(email) aan bij geslaagde login."""

    def __init__(self, root: tk.Tk, on_success):
        self.root = root
        self.on_success = on_success

        root.title("Excellent App — Inloggen")
        root.configure(bg=ACCENT)
        root.geometry("420x520")
        root.resizable(False, False)
        root.eval("tk::PlaceWindow . center")

        ico = os.path.join(os.path.abspath(os.path.dirname(__file__)), "logo.ico")
        if os.path.exists(ico):
            root.iconbitmap(ico)

        self._build()

    def _build(self):
        # Buitenste donkere achtergrond
        outer = tk.Frame(self.root, bg=ACCENT)
        outer.pack(fill=tk.BOTH, expand=True)

        # Witte kaart gecentreerd
        card = tk.Frame(outer, bg=BG, highlightbackground=BORDER, highlightthickness=1)
        card.place(relx=0.5, rely=0.5, anchor="center", width=360, height=440)

        # Logo / banner
        banner_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "logobanner.png")
        if PIL_AVAILABLE and os.path.exists(banner_path):
            try:
                img = Image.open(banner_path).convert("RGB")
                ratio = 280 / img.width
                h = int(img.height * ratio)
                img = img.resize((280, h), Image.LANCZOS)
                self._logo = ImageTk.PhotoImage(img)
                tk.Label(card, image=self._logo, bg=BG).pack(pady=(28, 0))
            except Exception:
                self._logo = None
                self._fallback_title(card)
        else:
            self._logo = None
            self._fallback_title(card)

        tk.Label(card, text="Inloggen op uw account",
                 bg=BG, fg=MUTED, font=(FONT, 9)).pack(pady=(6, 20))

        form = tk.Frame(card, bg=BG)
        form.pack(fill=tk.X, padx=32)

        # E-mail
        tk.Label(form, text="E-mailadres", bg=BG, fg=MUTED,
                 font=(FONT, 9), anchor="w").pack(fill=tk.X, pady=(0, 2))
        self.email_var = tk.StringVar()
        email_entry = tk.Entry(form, textvariable=self.email_var,
                               bg=BG_FIELD, fg=TEXT, insertbackground=TEXT,
                               relief=tk.FLAT, highlightbackground=BORDER,
                               highlightthickness=1, font=(FONT, 10))
        email_entry.pack(fill=tk.X, ipady=6)

        # Wachtwoord
        tk.Label(form, text="Wachtwoord", bg=BG, fg=MUTED,
                 font=(FONT, 9), anchor="w").pack(fill=tk.X, pady=(14, 2))
        self.pass_var = tk.StringVar()
        pass_entry = tk.Entry(form, textvariable=self.pass_var, show="•",
                              bg=BG_FIELD, fg=TEXT, insertbackground=TEXT,
                              relief=tk.FLAT, highlightbackground=BORDER,
                              highlightthickness=1, font=(FONT, 10))
        pass_entry.pack(fill=tk.X, ipady=6)

        # Foutmelding
        self.error_var = tk.StringVar()
        self.error_lbl = tk.Label(form, textvariable=self.error_var,
                                  bg=BG, fg=ERROR, font=(FONT, 8),
                                  anchor="w", wraplength=296)
        self.error_lbl.pack(fill=tk.X, pady=(8, 0))

        # Login knop
        self.btn = tk.Button(form, text="Inloggen", command=self._do_login,
                             bg=ACCENT, fg="#FFFFFF",
                             activebackground="#2E2E4E", activeforeground="#FFFFFF",
                             font=(FONT, 10, "bold"), bd=0, relief=tk.FLAT,
                             pady=10, cursor="hand2")
        self.btn.pack(fill=tk.X, pady=(18, 0))

        # Wachtwoord vergeten link
        forgot_lbl = tk.Label(form, text="Wachtwoord vergeten?",
                              bg=BG, fg=MUTED, font=(FONT, 8, "underline"),
                              cursor="hand2")
        forgot_lbl.pack(pady=(10, 0))
        forgot_lbl.bind("<Button-1>", lambda e: self._forgot_password())

        # Enter-toets koppelen
        self.root.bind("<Return>", lambda e: self._do_login())
        email_entry.focus_set()

    def _forgot_password(self):
        import supabase_client as sc
        from tkinter import simpledialog
        email = simpledialog.askstring(
            "Wachtwoord vergeten",
            "Vul uw e-mailadres in:",
            parent=self.root
        )
        if not email:
            return
        try:
            sc.reset_password(email.strip())
            self.error_var.set("")
            from tkinter import messagebox
            messagebox.showinfo(
                "E-mail verstuurd",
                f"Een reset-link is verstuurd naar {email}.\nControleer uw inbox.",
                parent=self.root
            )
        except Exception as e:
            self.error_var.set(f"Kon e-mail niet versturen: {e}")

    def _fallback_title(self, parent):
        tk.Label(parent, text="Excellent App",
                 bg=BG, fg=ACCENT, font=(FONT, 18, "bold")).pack(pady=(32, 0))

    def _do_login(self):
        import supabase_client as sc

        email = self.email_var.get().strip()
        password = self.pass_var.get()
        self.error_var.set("")

        if not email or not password:
            self.error_var.set("Vul uw e-mailadres en wachtwoord in.")
            return

        self.btn.config(text="Bezig...", state=tk.DISABLED)
        self.root.update()

        try:
            sc.login(email, password)
            self.on_success(email)
        except Exception as e:
            msg = str(e)
            if "Invalid login" in msg or "invalid_credentials" in msg:
                self.error_var.set("Verkeerd e-mailadres of wachtwoord.")
            else:
                self.error_var.set(f"Inloggen mislukt: {msg}")
            self.btn.config(text="Inloggen", state=tk.NORMAL)
