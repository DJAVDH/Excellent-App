import tkinter as tk
from tkinter import ttk, messagebox
import os
import subprocess
import threading
import tempfile
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.utils import formatdate
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.lib.utils import simpleSplit
import win32print

BG         = "#FFFFFF"
BG_FIELD   = "#F7F9FC"
BG_ROW_ALT = "#F0F4F8"
ACCENT     = "#1A1A2E"
TEXT_DARK  = "#1A1A2E"
TEXT_MUTED = "#8A8FA8"
BORDER     = "#E8EAF0"
FONT       = "Segoe UI"
ORANGE     = "#FF6B35"
RED        = "#E53935"

BTN_STYLE = dict(
    bg=ACCENT, fg="#FFFFFF",
    activebackground="#2E2E4E", activeforeground="#FFFFFF",
    font=(FONT, 10), bd=0, relief=tk.FLAT,
    padx=14, pady=7, cursor="hand2"
)
BTN_SECONDARY = dict(
    bg=BG_FIELD, fg=TEXT_DARK,
    activebackground="#E8EAF0", activeforeground=TEXT_DARK,
    font=(FONT, 10), bd=0, relief=tk.FLAT,
    padx=14, pady=7, cursor="hand2"
)


def _label(parent, text, **kw):
    return tk.Label(parent, text=text, bg=BG, fg=TEXT_MUTED,
                    font=(FONT, 9), anchor="w", **kw)


def _entry(parent, textvariable=None, width=28):
    return tk.Entry(parent, textvariable=textvariable, width=width,
                    bg=BG_FIELD, fg=TEXT_DARK, insertbackground=TEXT_DARK,
                    relief=tk.FLAT, highlightbackground=BORDER,
                    highlightthickness=1, font=(FONT, 10))


class KastBonComponent:
    """123kast Leveringsbon + Tevredenheidsonderzoek"""

    SOORT_OPTIONS = ["Full service", "Deels service", "Montage only", "Levering", "Ophaal"]

    def __init__(self, parent, app=None):
        self.parent = parent
        self.app = app
        self._products = []   # lijst van (product, afmetingen, aantallen)

        self.frame = tk.Frame(parent, bg=BG)
        self.frame.pack(fill=tk.BOTH, expand=True)

        card = tk.Frame(self.frame, bg=BG, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True)

        form = tk.Frame(card, bg=BG)
        form.pack(fill=tk.BOTH, expand=True, padx=20, pady=16)
        for col in range(4):
            form.columnconfigure(col, weight=1)

        # ── Variables ────────────────────────────────────────────────────────
        self.datum_var       = tk.StringVar()
        self.tijdsvak_var    = tk.StringVar()
        self.soort_var       = tk.StringVar()
        self.klantnaam_var   = tk.StringVar()
        self.ordernummer_var = tk.StringVar()
        self.straatnaam_var  = tk.StringVar()
        self.postcode_var    = tk.StringVar()
        self.plaatsnaam_var  = tk.StringVar()
        self.telefoon_var    = tk.StringVar()
        self.email_var       = tk.StringVar()

        r = 0

        # ── Rij 1: Datum | Tijdsvak | Soort levering ─────────────────────────
        _label(form, "Datum levering *").grid(row=r, column=0, sticky='w', pady=(0, 2), padx=(0, 8))
        _label(form, "Tijdsvak").grid(row=r, column=1, sticky='w', pady=(0, 2), padx=(0, 8))
        _label(form, "Soort levering").grid(row=r, column=2, sticky='w', pady=(0, 2), padx=(0, 8))
        r += 1
        _entry(form, self.datum_var).grid(row=r, column=0, sticky='ew', pady=(0, 12), padx=(0, 12))
        _entry(form, self.tijdsvak_var).grid(row=r, column=1, sticky='ew', pady=(0, 12), padx=(0, 12))
        ttk.Combobox(form, textvariable=self.soort_var, values=self.SOORT_OPTIONS,
                     width=16, font=(FONT, 10)).grid(row=r, column=2, columnspan=2,
                                                      sticky='ew', pady=(0, 12))
        r += 1

        # ── Rij 2: Klantnaam | Ordernummer | Telefoon ────────────────────────
        _label(form, "Klantnaam *").grid(row=r, column=0, sticky='w', pady=(0, 2), padx=(0, 8))
        _label(form, "Ordernummer *").grid(row=r, column=1, sticky='w', pady=(0, 2), padx=(0, 8))
        _label(form, "Telefoonnummer").grid(row=r, column=2, sticky='w', pady=(0, 2), padx=(0, 8))
        r += 1
        _entry(form, self.klantnaam_var).grid(row=r, column=0, sticky='ew', pady=(0, 12), padx=(0, 12))
        _entry(form, self.ordernummer_var).grid(row=r, column=1, sticky='ew', pady=(0, 12), padx=(0, 12))
        _entry(form, self.telefoon_var).grid(row=r, column=2, columnspan=2, sticky='ew', pady=(0, 12))
        r += 1

        # ── Rij 3: Straatnaam | Postcode | Plaatsnaam ────────────────────────
        _label(form, "Straatnaam + huisnummer *").grid(row=r, column=0, sticky='w', pady=(0, 2), padx=(0, 8))
        _label(form, "Postcode *").grid(row=r, column=1, sticky='w', pady=(0, 2), padx=(0, 8))
        _label(form, "Plaatsnaam *").grid(row=r, column=2, sticky='w', pady=(0, 2), padx=(0, 8))
        r += 1
        _entry(form, self.straatnaam_var).grid(row=r, column=0, sticky='ew', pady=(0, 12), padx=(0, 12))
        _entry(form, self.postcode_var, width=12).grid(row=r, column=1, sticky='ew', pady=(0, 12), padx=(0, 12))
        _entry(form, self.plaatsnaam_var).grid(row=r, column=2, columnspan=2, sticky='ew', pady=(0, 12))
        r += 1

        # ── Rij 4: E-mailadres klant ─────────────────────────────────────────
        _label(form, "E-mailadres klant").grid(row=r, column=0, sticky='w', pady=(0, 2), padx=(0, 8))
        r += 1
        _entry(form, self.email_var, width=40).grid(row=r, column=0, columnspan=4, sticky='ew', pady=(0, 12))
        r += 1

        # ── Divider ──────────────────────────────────────────────────────────
        tk.Frame(form, bg=BORDER, height=1).grid(row=r, column=0, columnspan=4,
                                                  sticky='ew', pady=(0, 10))
        r += 1

        # ── Producten sectie ─────────────────────────────────────────────────
        prod_hdr = tk.Frame(form, bg=BG)
        prod_hdr.grid(row=r, column=0, columnspan=4, sticky='ew', pady=(0, 6))
        tk.Label(prod_hdr, text="Producten", bg=BG, fg=TEXT_DARK,
                 font=(FONT, 10, "bold")).pack(side=tk.LEFT)
        r += 1

        # Productenlijst container
        self._prod_list_frame = tk.Frame(form, bg=BG)
        self._prod_list_frame.grid(row=r, column=0, columnspan=4, sticky='ew', pady=(0, 6))
        self._prod_list_frame.columnconfigure(0, weight=3)
        self._prod_list_frame.columnconfigure(1, weight=2)
        self._prod_list_frame.columnconfigure(2, weight=1)
        self._prod_list_frame.columnconfigure(3, weight=0)

        # Kolomkoppen in de lijst
        for ci, txt in enumerate(["Product type", "Afmetingen", "Aantal", ""]):
            tk.Label(self._prod_list_frame, text=txt, bg=BG, fg=TEXT_MUTED,
                     font=(FONT, 8), anchor="w").grid(
                row=0, column=ci, sticky='w', padx=(0, 6), pady=(0, 2))
        r += 1

        # Invoerrij voor nieuw product
        inp_row = tk.Frame(form, bg=BG)
        inp_row.grid(row=r, column=0, columnspan=4, sticky='ew', pady=(0, 8))
        inp_row.columnconfigure(0, weight=3)
        inp_row.columnconfigure(1, weight=2)
        inp_row.columnconfigure(2, weight=1)

        self._inp_product    = _entry(inp_row)
        self._inp_afmetingen = _entry(inp_row)
        self._inp_aantallen  = _entry(inp_row, width=8)
        self._inp_product.grid(row=0, column=0, sticky='ew', padx=(0, 6))
        self._inp_afmetingen.grid(row=0, column=1, sticky='ew', padx=(0, 6))
        self._inp_aantallen.grid(row=0, column=2, sticky='ew', padx=(0, 6))

        tk.Button(inp_row, text="+ Toevoegen", command=self._add_product,
                  bg=ORANGE, fg="#FFFFFF", activebackground="#E55A25",
                  activeforeground="#FFFFFF", font=(FONT, 9, "bold"),
                  bd=0, relief=tk.FLAT, padx=10, pady=5,
                  cursor="hand2").grid(row=0, column=3, sticky='ew')
        r += 1

        # ── Divider ──────────────────────────────────────────────────────────
        tk.Frame(form, bg=BORDER, height=1).grid(row=r, column=0, columnspan=4,
                                                  sticky='ew', pady=(0, 10))
        r += 1

        # ── Bijzonderheden ───────────────────────────────────────────────────
        _label(form, "Bijzonderheden").grid(row=r, column=0, sticky='w', pady=(0, 2))
        r += 1
        self.bijzonderheden_text = tk.Text(
            form, height=3, width=60, bg=BG_FIELD, fg=TEXT_DARK,
            insertbackground=TEXT_DARK, relief=tk.FLAT,
            highlightbackground=BORDER, highlightthickness=1,
            font=(FONT, 10), wrap=tk.WORD)
        self.bijzonderheden_text.grid(row=r, column=0, columnspan=4, sticky='ew', pady=(0, 12))
        r += 1

        # ── Printer + Papierlade ─────────────────────────────────────────────
        _label(form, "Printer").grid(row=r, column=0, sticky='w', pady=(0, 2), padx=(0, 8))
        _label(form, "Papierlade").grid(row=r, column=3, sticky='w', pady=(0, 2))
        r += 1
        self.printer_var = tk.StringVar()
        self.printer_combo = ttk.Combobox(form, textvariable=self.printer_var,
                                          state="readonly", width=28, font=(FONT, 10))
        self.printer_combo.grid(row=r, column=0, columnspan=3, sticky='ew', pady=(0, 4), padx=(0, 12))
        self._tray_map = {"Lade 1": 1, "Lade 2": 3, "Lade 3 (Stickers)": 2}
        self.tray_var = tk.StringVar()
        self.tray_combo = ttk.Combobox(form, textvariable=self.tray_var,
                                       values=list(self._tray_map.keys()),
                                       state="readonly", width=14, font=(FONT, 10))
        self.tray_combo.grid(row=r, column=3, sticky='ew', pady=(0, 4))
        self.tray_combo.set("Lade 1")

        self.set_defaults()
        threading.Thread(target=self._detect_printers_bg, daemon=True).start()

        # ── Buttons ──────────────────────────────────────────────────────────
        tk.Frame(card, bg=BORDER, height=1).pack(fill=tk.X)
        btn_bar = tk.Frame(card, bg=BG)
        btn_bar.pack(fill=tk.X, padx=20, pady=14)
        tk.Button(btn_bar, text="Genereer PDF",    command=self.gen_pdf,      **BTN_STYLE).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_bar, text="Direct Printen",  command=self.print_direct, **BTN_STYLE).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_bar, text="Mail: Datum",    command=self._send_datum_mail,
                  bg=ORANGE, fg="#FFFFFF", activebackground="#E55A25",
                  activeforeground="#FFFFFF", font=(FONT, 10), bd=0,
                  relief=tk.FLAT, padx=14, pady=7, cursor="hand2").pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_bar, text="Mail: Tijdvak",  command=self._send_tijdvak_mail,
                  bg=ORANGE, fg="#FFFFFF", activebackground="#E55A25",
                  activeforeground="#FFFFFF", font=(FONT, 10), bd=0,
                  relief=tk.FLAT, padx=14, pady=7, cursor="hand2").pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_bar, text="Velden Leegmaken",command=self.clear_fields, **BTN_SECONDARY).pack(side=tk.LEFT)

    # ── Productenlijst ────────────────────────────────────────────────────────

    def _add_product(self):
        p = self._inp_product.get().strip()
        a = self._inp_afmetingen.get().strip()
        n = self._inp_aantallen.get().strip()
        if not p:
            messagebox.showwarning("Leeg veld", "Vul minimaal het product type in.")
            return
        self._products.append((p, a, n))
        self._inp_product.delete(0, tk.END)
        self._inp_afmetingen.delete(0, tk.END)
        self._inp_aantallen.delete(0, tk.END)
        self._refresh_product_list()
        self._inp_product.focus()

    def _remove_product(self, idx):
        self._products.pop(idx)
        self._refresh_product_list()

    def _refresh_product_list(self):
        for w in self._prod_list_frame.winfo_children():
            if int(w.grid_info().get("row", 0)) > 0:
                w.destroy()

        for i, (p, a, n) in enumerate(self._products):
            bg = BG_ROW_ALT if i % 2 else BG
            row_idx = i + 1
            for ci, txt in enumerate([p, a, n]):
                tk.Label(self._prod_list_frame, text=txt, bg=bg, fg=TEXT_DARK,
                         font=(FONT, 9), anchor="w",
                         highlightbackground=BORDER, highlightthickness=1
                         ).grid(row=row_idx, column=ci, sticky='ew', padx=(0, 4), pady=1)
            tk.Button(self._prod_list_frame, text="✕",
                      command=lambda idx=i: self._remove_product(idx),
                      bg=bg, fg=RED, activebackground="#FFEBEE",
                      font=(FONT, 9, "bold"), bd=0, relief=tk.FLAT,
                      cursor="hand2", padx=4).grid(
                row=row_idx, column=3, sticky='ew', pady=1)

    # ── Helpers ───────────────────────────────────────────────────────────────

    def set_defaults(self):
        tomorrow = (datetime.now() + timedelta(days=1)).strftime("%d-%m-%Y")
        self.datum_var.set(tomorrow)
        self.tijdsvak_var.set("08:00-12:00")
        self.soort_var.set("Full service")

    def clear_fields(self):
        self.klantnaam_var.set("")
        self.ordernummer_var.set("")
        self.straatnaam_var.set("")
        self.postcode_var.set("")
        self.plaatsnaam_var.set("")
        self.telefoon_var.set("")
        self.email_var.set("")
        self._products.clear()
        self._refresh_product_list()
        self._inp_product.delete(0, tk.END)
        self._inp_afmetingen.delete(0, tk.END)
        self._inp_aantallen.delete(0, tk.END)
        self.bijzonderheden_text.delete("1.0", tk.END)

    def _detect_printers_bg(self):
        try:
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
            names = [p[2] for p in printers]
        except Exception:
            names = []

        def apply():
            try:
                self.printer_combo['values'] = names
                if names and not self.printer_combo.get():
                    canon = next((p for p in names if 'canon' in p.lower()), None)
                    self.printer_combo.set(canon or names[0])
            except Exception:
                pass

        try:
            self.app.root.after(0, apply)
        except Exception:
            apply()

    # ── PDF ───────────────────────────────────────────────────────────────────

    def _build_pdf(self, path):
        datum       = self.datum_var.get().strip()
        tijdsvak    = self.tijdsvak_var.get().strip()
        soort       = self.soort_var.get().strip()
        klantnaam   = self.klantnaam_var.get().strip()
        ordernummer = self.ordernummer_var.get().strip()
        straatnaam  = self.straatnaam_var.get().strip()
        postcode    = self.postcode_var.get().strip()
        plaatsnaam  = self.plaatsnaam_var.get().strip()
        telefoon    = self.telefoon_var.get().strip()
        bijz        = self.bijzonderheden_text.get("1.0", tk.END).strip()

        if not datum or not klantnaam or not ordernummer or not straatnaam or not postcode or not plaatsnaam:
            messagebox.showerror("Fout", "Vul alle verplichte velden in (*).")
            return False

        products = self._products if self._products else [("—", "—", "—")]

        c = rl_canvas.Canvas(path, pagesize=A4)
        W, H = A4
        M = 15 * mm

        self._page1(c, W, H, M, datum, tijdsvak, soort, klantnaam, ordernummer,
                    straatnaam, postcode, plaatsnaam, telefoon, products, bijz)
        c.showPage()
        self._page2(c, W, H, M, klantnaam, ordernummer, datum)
        c.showPage()
        self._page3(c, W, H, M)
        c.showPage()
        self._page4(c, W, H, M)
        c.save()
        return True

    # ── Pagina 1: Leveringsbon ────────────────────────────────────────────────

    def _page1(self, c, W, H, M, datum, tijdsvak, soort, klantnaam, ordernummer,
               straatnaam, postcode, plaatsnaam, telefoon, products, bijz):
        y = H - M
        base = os.path.abspath(os.path.dirname(__file__))

        logo_path = os.path.join(base, "logobanner.png")
        logo_h = 18 * mm
        if os.path.exists(logo_path):
            try:
                from PIL import Image as PILImage
                img = PILImage.open(logo_path)
                ow, oh = img.size
                logo_w = (ow / oh) * logo_h
                c.drawImage(logo_path, M, y - logo_h, width=logo_w, height=logo_h,
                            preserveAspectRatio=True, mask='auto')
            except Exception:
                pass

        c.setFont("Helvetica-Bold", 20)
        c.setFillColor(colors.HexColor("#1A1A2E"))
        c.drawRightString(W - M, y - logo_h / 2 - 3 * mm, "Leveringsbon")
        y -= logo_h + 5 * mm

        c.setStrokeColor(colors.HexColor("#1A1A2E"))
        c.setLineWidth(2)
        c.line(M, y, W - M, y)
        y -= 8 * mm

        lw = 47 * mm
        hw = (W - 2 * M - lw * 2) / 2

        def info_row(l1, v1, l2="", v2=""):
            nonlocal y
            c.setFont("Helvetica-Bold", 9)
            c.setFillColor(colors.HexColor("#8A8FA8"))
            c.drawString(M, y, l1)
            c.setFont("Helvetica", 10)
            c.setFillColor(colors.HexColor("#1A1A2E"))
            c.drawString(M + lw, y, v1)
            if l2:
                c.setFont("Helvetica-Bold", 9)
                c.setFillColor(colors.HexColor("#8A8FA8"))
                c.drawString(M + lw + hw + 5 * mm, y, l2)
                c.setFont("Helvetica", 10)
                c.setFillColor(colors.HexColor("#1A1A2E"))
                c.drawString(M + lw + hw + 5 * mm + lw, y, v2)
            y -= 7 * mm

        info_row("Datum levering:", datum,  "Tijdsvak levering:", tijdsvak)
        info_row("Soort levering:", soort,  "Ordernummer:", ordernummer)
        info_row("Klantnaam:", klantnaam,   "Telefoonnummer:", telefoon)
        info_row("Straatnaam:", straatnaam, "Postcode:", postcode)
        info_row("Plaatsnaam:", plaatsnaam)
        y -= 3 * mm

        c.setStrokeColor(colors.HexColor("#E8EAF0"))
        c.setLineWidth(1)
        c.line(M, y, W - M, y)
        y -= 6 * mm

        # Product tabel (meerdere rijen)
        col_w = [75 * mm, 65 * mm, 30 * mm]
        hdr = [["Product type", "Afmetingen", "Aantallen"]]
        rows = [[p, a, n] for p, a, n in products]
        tbl_data = hdr + rows
        row_h = [9 * mm] + [8 * mm] * len(rows)
        tbl = Table(tbl_data, colWidths=col_w, rowHeights=row_h)
        tbl.setStyle(TableStyle([
            ('BACKGROUND',     (0, 0), (-1, 0),  colors.HexColor("#1A1A2E")),
            ('TEXTCOLOR',      (0, 0), (-1, 0),  colors.white),
            ('FONTNAME',       (0, 0), (-1, 0),  'Helvetica-Bold'),
            ('FONTSIZE',       (0, 0), (-1, 0),  9),
            ('FONTNAME',       (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE',       (0, 1), (-1, -1), 9),
            ('INNERGRID',      (0, 0), (-1, -1), 0.5, colors.HexColor("#B0B5C8")),
            ('BOX',            (0, 0), (-1, -1), 1,   colors.HexColor("#1A1A2E")),
            ('VALIGN',         (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING',    (0, 0), (-1, -1), 5),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F9FC")]),
        ]))
        tw, th = tbl.wrapOn(c, W - 2 * M, H)
        tbl.drawOn(c, M, y - th)
        y -= th + 6 * mm

        # Bijzonderheden
        bijz_h = 18 * mm
        c.setFont("Helvetica-Bold", 9)
        c.setFillColor(colors.HexColor("#8A8FA8"))
        c.drawString(M, y, "Bijzonderheden:")
        y -= 5 * mm
        c.setStrokeColor(colors.HexColor("#1A1A2E"))
        c.setLineWidth(1)
        c.rect(M, y - bijz_h, W - 2 * M, bijz_h)
        if bijz:
            to = c.beginText(M + 3 * mm, y - 6 * mm)
            to.setFont("Helvetica", 9)
            to.setLeading(11)
            to.setFillColor(colors.HexColor("#1A1A2E"))
            for line in simpleSplit(bijz, "Helvetica", 9, W - 2 * M - 6 * mm):
                to.textLine(line)
            c.drawText(to)
        y -= bijz_h + 7 * mm

        # Checklist
        c.setStrokeColor(colors.HexColor("#E8EAF0"))
        c.line(M, y, W - M, y)
        y -= 7 * mm

        checklist = [
            "Item(s) gecontroleerd op schades voor aanvang laden",
            "Item(s) gecontroleerd op juiste aantal colli",
            "Klant gebeld van te voren met exacte aankomsttijd",
            "Schade rondje aanlooproute met foto's gedaan",
            "Transportlatten verwijderd",
            "Spotjes vastgeklikt",
            "Item(s) gecontroleerd op schades na leveren",
            "Aflever foto gemaakt",
            "Klant tevredenheids onderzoek laten invullen",
        ]

        box_sz = 3.5 * mm
        line_h = 7 * mm
        half   = (len(checklist) + 1) // 2
        col2_x = W / 2 + 5 * mm

        for i, item in enumerate(checklist):
            if i < half:
                xi = M
                yi = y - i * line_h
            else:
                xi = col2_x
                yi = y - (i - half) * line_h
            c.setStrokeColor(colors.HexColor("#1A1A2E"))
            c.setLineWidth(0.8)
            c.rect(xi, yi - box_sz + 1 * mm, box_sz, box_sz, fill=0)
            c.setFont("Helvetica", 8.5)
            c.setFillColor(colors.HexColor("#1A1A2E"))
            c.drawString(xi + box_sz + 2 * mm, yi - box_sz + 1.5 * mm, item)

        y -= half * line_h + 10 * mm

        decl = ("Ondergetekende verklaart de geleverde kasten volledig en zonder zichtbare schade "
                "in goede staat te hebben ontvangen.")
        c.setFont("Helvetica-Bold", 9)
        c.setFillColor(colors.HexColor("#1A1A2E"))
        for line in simpleSplit(decl, "Helvetica-Bold", 9, W - 2 * M):
            c.drawString(M, y, line)
            y -= 5 * mm
        y -= 4 * mm

        y -= 12 * mm

        sig_w = (W - 2 * M - 20 * mm) / 2
        c.setStrokeColor(colors.HexColor("#1A1A2E"))
        c.setLineWidth(0.8)
        c.line(M, y, M + sig_w, y)
        c.line(M + sig_w + 20 * mm, y, W - M, y)
        y -= 5 * mm
        c.setFont("Helvetica-Bold", 9)
        c.setFillColor(colors.HexColor("#1A1A2E"))
        c.drawString(M, y, "Handtekening klant:")
        c.drawString(M + sig_w + 20 * mm, y, "Handtekening voorman:")

        c.setFont("Helvetica", 8)
        c.setFillColor(colors.HexColor("#8A8FA8"))
        c.drawCentredString(W / 2, 10 * mm, "Excellent Packing and Moving B.V.  •  123kast Leveringsbon")

    # ── Pagina 2: Tevredenheidsonderzoek deel 1 ───────────────────────────────

    def _page2(self, c, W, H, M, klantnaam, ordernummer, datum):
        y = H - M
        base = os.path.abspath(os.path.dirname(__file__))

        logo_path = os.path.join(base, "123kast-banner.jpg")
        logo_h = 20 * mm
        if os.path.exists(logo_path):
            try:
                from PIL import Image as PILImage
                img = PILImage.open(logo_path)
                ow, oh = img.size
                logo_w = min((ow / oh) * logo_h, W - 2 * M)
                c.drawImage(logo_path, M, y - logo_h, width=logo_w, height=logo_h,
                            preserveAspectRatio=True, mask='auto')
            except Exception:
                c.setFont("Helvetica-Bold", 18)
                c.setFillColor(colors.HexColor("#1A1A2E"))
                c.drawString(M, y - logo_h / 2, "123kast.nl")
        y -= logo_h + 6 * mm

        c.setFont("Helvetica-Bold", 14)
        c.setFillColor(colors.HexColor("#1A1A2E"))
        c.drawString(M, y, "TEVREDENHEIDSONDERZOEK")
        y -= 6 * mm
        c.setFont("Helvetica", 9)
        c.setFillColor(colors.HexColor("#8A8FA8"))
        c.drawString(M, y, "123kast.nl & Excellent Packing and Moving")
        y -= 7 * mm

        c.setStrokeColor(colors.HexColor("#E8EAF0"))
        c.setLineWidth(1)
        c.line(M, y, W - M, y)
        y -= 8 * mm

        # Klantinformatie
        c.setFont("Helvetica-Bold", 10)
        c.setFillColor(colors.HexColor("#1A1A2E"))
        c.drawString(M, y, "Klantinformatie")
        y -= 7 * mm

        for label, value in [("Naam:", klantnaam), ("Ordernummer:", ordernummer), ("Datum levering:", datum)]:
            c.setFont("Helvetica", 9)
            c.setFillColor(colors.HexColor("#8A8FA8"))
            c.drawString(M, y, label)
            c.setFillColor(colors.HexColor("#1A1A2E"))
            c.drawString(M + 40 * mm, y, value)
            c.setStrokeColor(colors.HexColor("#CCCCCC"))
            c.setLineWidth(0.5)
            c.line(M + 40 * mm, y - 1.5 * mm, W - M, y - 1.5 * mm)
            y -= 7 * mm
        y -= 4 * mm

        def section_hdr(text):
            nonlocal y
            c.setFillColor(colors.HexColor("#1A1A2E"))
            c.rect(M, y - 8 * mm, W - 2 * M, 8 * mm, fill=1, stroke=0)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 9)
            c.drawString(M + 3 * mm, y - 5.5 * mm, text)
            y -= 12 * mm

        def sub(num, text):
            nonlocal y
            c.setFont("Helvetica-Bold", 9)
            c.setFillColor(colors.HexColor("#1A1A2E"))
            c.drawString(M, y, f"{num}  {text}")
            y -= 6 * mm

        def rating():
            nonlocal y
            bs = 4 * mm
            c.setFont("Helvetica", 8)
            c.setFillColor(colors.HexColor("#8A8FA8"))
            c.drawString(M, y, "Beoordeling:")
            y -= 6 * mm
            for i in range(1, 6):
                bx = M + (i - 1) * 16 * mm
                c.setStrokeColor(colors.HexColor("#1A1A2E"))
                c.setLineWidth(0.8)
                c.rect(bx, y - bs + 1 * mm, bs, bs, fill=0)
                c.setFont("Helvetica", 9)
                c.setFillColor(colors.HexColor("#1A1A2E"))
                c.drawString(bx + bs + 2 * mm, y - bs + 1.5 * mm, str(i))
            y -= 9 * mm

        def checkboxes(options):
            nonlocal y
            bs = 4 * mm
            for opt in options:
                c.setStrokeColor(colors.HexColor("#1A1A2E"))
                c.setLineWidth(0.8)
                c.rect(M, y - bs + 1 * mm, bs, bs, fill=0)
                c.setFont("Helvetica", 9)
                c.setFillColor(colors.HexColor("#1A1A2E"))
                c.drawString(M + bs + 3 * mm, y - bs + 1.5 * mm, opt)
                y -= 8 * mm
            y -= 2 * mm

        def toelichting():
            nonlocal y
            c.setFont("Helvetica", 8)
            c.setFillColor(colors.HexColor("#8A8FA8"))
            c.drawString(M, y, "Toelichting:")
            y -= 5 * mm
            c.setStrokeColor(colors.HexColor("#AAAAAA"))
            c.setLineWidth(0.5)
            c.line(M, y, W - M, y)
            y -= 10 * mm

        # ── Sectie 1 ─────────────────────────────────────────────────────────
        section_hdr("1. PRODUCTBEOORDELING – 123kast.nl")

        sub("1.1", "Algemene tevredenheid over de kast")
        rating()
        sub("1.2", "Voldoet de kast aan uw verwachtingen?")
        checkboxes(["Ja, volledig", "Gedeeltelijk", "Nee, niet helemaal"])
        sub("1.3", "Materiaalkwaliteit")
        rating()
        sub("1.4", "Afwerking van het product")
        rating()
        sub("1.5", "Prijs–kwaliteitverhouding")
        rating()
        sub("1.6", "Opmerkingen over het product")
        toelichting()

        c.setFont("Helvetica", 8)
        c.setFillColor(colors.HexColor("#8A8FA8"))
        c.drawCentredString(W / 2, 10 * mm,
                            "123kast.nl & Excellent Packing and Moving B.V.  •  Tevredenheidsonderzoek  (1/3)")

    # ── Pagina 3: Tevredenheidsonderzoek deel 2 (Sectie 2 + 3) ──────────────

    def _page3(self, c, W, H, M):
        y = H - M
        base = os.path.abspath(os.path.dirname(__file__))

        logo_path = os.path.join(base, "123kast-banner.jpg")
        logo_h = 14 * mm
        if os.path.exists(logo_path):
            try:
                from PIL import Image as PILImage
                img = PILImage.open(logo_path)
                ow, oh = img.size
                logo_w = min((ow / oh) * logo_h, 80 * mm)
                c.drawImage(logo_path, M, y - logo_h, width=logo_w, height=logo_h,
                            preserveAspectRatio=True, mask='auto')
            except Exception:
                pass
        y -= logo_h + 8 * mm

        def section_hdr(text):
            nonlocal y
            c.setFillColor(colors.HexColor("#1A1A2E"))
            c.rect(M, y - 8 * mm, W - 2 * M, 8 * mm, fill=1, stroke=0)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 9)
            c.drawString(M + 3 * mm, y - 5.5 * mm, text)
            y -= 12 * mm

        def sub(num, text):
            nonlocal y
            c.setFont("Helvetica-Bold", 9)
            c.setFillColor(colors.HexColor("#1A1A2E"))
            c.drawString(M, y, f"{num}  {text}")
            y -= 6 * mm

        def rating():
            nonlocal y
            bs = 4 * mm
            c.setFont("Helvetica", 8)
            c.setFillColor(colors.HexColor("#8A8FA8"))
            c.drawString(M, y, "Beoordeling:")
            y -= 6 * mm
            for i in range(1, 6):
                bx = M + (i - 1) * 16 * mm
                c.setStrokeColor(colors.HexColor("#1A1A2E"))
                c.setLineWidth(0.8)
                c.rect(bx, y - bs + 1 * mm, bs, bs, fill=0)
                c.setFont("Helvetica", 9)
                c.setFillColor(colors.HexColor("#1A1A2E"))
                c.drawString(bx + bs + 2 * mm, y - bs + 1.5 * mm, str(i))
            y -= 9 * mm

        def checkboxes(options):
            nonlocal y
            bs = 4 * mm
            for opt in options:
                c.setStrokeColor(colors.HexColor("#1A1A2E"))
                c.setLineWidth(0.8)
                c.rect(M, y - bs + 1 * mm, bs, bs, fill=0)
                c.setFont("Helvetica", 9)
                c.setFillColor(colors.HexColor("#1A1A2E"))
                c.drawString(M + bs + 3 * mm, y - bs + 1.5 * mm, opt)
                y -= 8 * mm
            y -= 2 * mm

        def toelichting():
            nonlocal y
            c.setFont("Helvetica", 8)
            c.setFillColor(colors.HexColor("#8A8FA8"))
            c.drawString(M, y, "Toelichting:")
            y -= 5 * mm
            c.setStrokeColor(colors.HexColor("#AAAAAA"))
            c.setLineWidth(0.5)
            c.line(M, y, W - M, y)
            y -= 10 * mm

        # ── Sectie 2 ─────────────────────────────────────────────────────────
        section_hdr("2. LEVERING & SERVICE – EXCELLENT PACKING AND MOVING")

        sub("2.1", "Communicatie voorafgaand aan de levering")
        rating()
        sub("2.2", "Professionaliteit & vriendelijkheid van de bezorgers")
        rating()
        sub("2.3", "Zorgvuldigheid bij levering en plaatsing")
        rating()
        sub("2.4", "Levering op afgesproken tijd")
        checkboxes(["Ja", "Nee, te vroeg", "Nee, te laat"])
        sub("2.5", "Opmerkingen over de levering")
        toelichting()

        y -= 4 * mm

        # ── Sectie 3 ─────────────────────────────────────────────────────────
        section_hdr("3. GEBRUIK VAN AFLEVERFOTO'S")

        sub("3.1", "Toestemming voor promotioneel gebruik")
        c.setFont("Helvetica", 9)
        c.setFillColor(colors.HexColor("#8A8FA8"))
        c.drawString(M, y, "Mogen wij de afleverfoto's van uw kast gebruiken voor marketingdoeleinden")
        y -= 5 * mm
        c.drawString(M, y, "(website, social media, brochures)?")
        y -= 8 * mm
        checkboxes(["Ja, dat is prima",
                    "Ja, maar alleen anoniem (geen herkenbare personen)",
                    "Nee, liever niet"])

        c.setFont("Helvetica", 8)
        c.setFillColor(colors.HexColor("#8A8FA8"))
        c.drawCentredString(W / 2, 10 * mm,
                            "123kast.nl & Excellent Packing and Moving B.V.  •  Tevredenheidsonderzoek  (2/3)")

    # ── Pagina 4: Tevredenheidsonderzoek deel 3 (Sectie 4 + handtekening) ────

    def _page4(self, c, W, H, M):
        y = H - M
        base = os.path.abspath(os.path.dirname(__file__))

        logo_path = os.path.join(base, "123kast-banner.jpg")
        logo_h = 14 * mm
        if os.path.exists(logo_path):
            try:
                from PIL import Image as PILImage
                img = PILImage.open(logo_path)
                ow, oh = img.size
                logo_w = min((ow / oh) * logo_h, 80 * mm)
                c.drawImage(logo_path, M, y - logo_h, width=logo_w, height=logo_h,
                            preserveAspectRatio=True, mask='auto')
            except Exception:
                pass
        y -= logo_h + 8 * mm

        def section_hdr(text):
            nonlocal y
            c.setFillColor(colors.HexColor("#1A1A2E"))
            c.rect(M, y - 8 * mm, W - 2 * M, 8 * mm, fill=1, stroke=0)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 9)
            c.drawString(M + 3 * mm, y - 5.5 * mm, text)
            y -= 12 * mm

        def sub(num, text):
            nonlocal y
            c.setFont("Helvetica-Bold", 9)
            c.setFillColor(colors.HexColor("#1A1A2E"))
            c.drawString(M, y, f"{num}  {text}")
            y -= 6 * mm

        def checkboxes(options):
            nonlocal y
            bs = 4 * mm
            for opt in options:
                c.setStrokeColor(colors.HexColor("#1A1A2E"))
                c.setLineWidth(0.8)
                c.rect(M, y - bs + 1 * mm, bs, bs, fill=0)
                c.setFont("Helvetica", 9)
                c.setFillColor(colors.HexColor("#1A1A2E"))
                c.drawString(M + bs + 3 * mm, y - bs + 1.5 * mm, opt)
                y -= 8 * mm
            y -= 4 * mm

        def toelichting(lines=2):
            nonlocal y
            c.setStrokeColor(colors.HexColor("#AAAAAA"))
            c.setLineWidth(0.5)
            for _ in range(lines):
                c.line(M, y, W - M, y)
                y -= 9 * mm
            y -= 4 * mm

        # ── Sectie 4 ─────────────────────────────────────────────────────────
        section_hdr("4. AANBEVELING & SLOT")

        sub("4.1", "Zou u ons aanbevelen aan anderen?")
        checkboxes(["Ja", "Misschien", "Nee"])

        sub("4.2", "Overige opmerkingen of suggesties:")
        toelichting(lines=3)

        # Handtekening
        y -= 4 * mm
        sig_w = 80 * mm
        c.setStrokeColor(colors.HexColor("#1A1A2E"))
        c.setLineWidth(0.8)
        c.line(M, y, M + sig_w, y)
        y -= 6 * mm
        c.setFont("Helvetica-Bold", 9)
        c.setFillColor(colors.HexColor("#1A1A2E"))
        c.drawString(M, y, "Handtekening klant:")

        c.setFont("Helvetica", 8)
        c.setFillColor(colors.HexColor("#8A8FA8"))
        c.drawCentredString(W / 2, 10 * mm,
                            "123kast.nl & Excellent Packing and Moving B.V.  •  Tevredenheidsonderzoek  (3/3)")

    # ── Mail ─────────────────────────────────────────────────────────────────

    def _open_outlook_mail(self, email, subject, html_body):
        try:
            base = os.path.abspath(os.path.dirname(__file__))

            msg = MIMEMultipart("related")
            msg["To"]       = email
            msg["Subject"]  = subject
            msg["Date"]     = formatdate(localtime=True)
            msg["X-Unsent"] = "1"

            msg.attach(MIMEText(html_body, "html", "utf-8"))

            for cid, fname, subtype in [
                ("logo_123kast",   "123kast-banner.jpg", "jpeg"),
                ("logo_excellent", "logobanner.png",     "png"),
            ]:
                path = os.path.join(base, fname)
                if os.path.exists(path):
                    with open(path, "rb") as f:
                        img = MIMEImage(f.read(), _subtype=subtype)
                    img.add_header("Content-ID", f"<{cid}>")
                    img.add_header("Content-Disposition", "inline", filename=fname)
                    msg.attach(img)

            tmp = tempfile.NamedTemporaryFile(
                suffix=".eml", delete=False, mode="w", encoding="utf-8"
            )
            tmp.write(msg.as_string())
            tmp.close()
            os.startfile(tmp.name)
        except Exception as e:
            messagebox.showerror("Fout", f"Kon mail-app niet openen: {e}")

    def _send_datum_mail(self):
        email = self.email_var.get().strip()
        datum = self.datum_var.get().strip()
        if not email:
            messagebox.showerror("Fout", "Vul het e-mailadres van de klant in.")
            return
        subject = "Bezorgafspraak 123Kast.nl"
        html = f"""<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#1a1a1a;">
<p>Geachte klant van 123kast.nl,</p>
<p>Graag willen wij u bij deze de datum doorgeven dat wij bij u langskomen.<br>
U heeft met 123kast afspraken gemaakt over leveren en/of uitpakken of andere werkzaamheden.<br>
Wij van Excellent Packing and Moving zullen dit uitvoeren.</p>
<p>Uw afspraak is ingepland voor: <strong>{datum}</strong></p>
<p>U ontvangt uiterlijk één werkdag van tevoren een e-mail met de tijdsindicatie.<br>
Hopende u hiermee voldoende te hebben geïnformeerd.</p>
<p style="color:#1155CC;">Met vriendelijke groet namens 123Kast.nl<br>
<strong>Excellent Packing and Moving</strong></p>
<br>
<img src="cid:logo_123kast" height="55" style="display:block;margin-bottom:6px;">
<img src="cid:logo_excellent" height="55" style="display:block;">
</body></html>"""
        self._open_outlook_mail(email, subject, html)

    def _send_tijdvak_mail(self):
        email   = self.email_var.get().strip()
        datum   = self.datum_var.get().strip()
        tijdvak = self.tijdsvak_var.get().strip()
        if not email:
            messagebox.showerror("Fout", "Vul het e-mailadres van de klant in.")
            return
        subject = "Tijdvak van levering 123Kast.nl"
        html = f"""<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#1a1a1a;">
<p>Geachte klant van 123kast.nl,</p>
<p>Graag willen wij u bij deze de tijdsindicatie doorgeven dat wij bij u langskomen.<br>
U kunt ons verwachten op: <strong>{datum}</strong> tussen <strong>{tijdvak}</strong>.<br>
<em>*De aangegeven tijd is een indicatie en kan afhankelijk van omstandigheden afwijken.</em></p>
<p>Onze bezorger(s) zullen circa 30 minuten van te voren bellen voor aankomst.</p>
<p>Om de levering zo soepel als mogelijk te laten verlopen, vragen wij u het volgende voor te bereiden:</p>
<ul>
<li>Zorg voor een vrije toegang tot de plek waar de kast moet komen</li>
<li>Heeft u gekozen voor Full Service levering? Zorg er dan voor dat de plek waar de kast moet komen helemaal vrij en schoon is</li>
</ul>
<p><strong>Let op:</strong><br>
Excellent Packing and Moving en/of 123kast.nl is niet verantwoordelijk voor eventuele schade
die kan ontstaan aan en rond de woning tijdens leveren, tenzij er sprake en bewijs is van grove nalatigheid.</p>
<p>Hopende u hiermee voldoende te hebben geïnformeerd.</p>
<p style="color:#1155CC;">Met vriendelijke groet namens 123Kast.nl<br>
<strong>Excellent Packing and Moving</strong></p>
<br>
<img src="cid:logo_123kast" height="55" style="display:block;margin-bottom:6px;">
<img src="cid:logo_excellent" height="55" style="display:block;">
</body></html>"""
        self._open_outlook_mail(email, subject, html)

    # ── Printen ───────────────────────────────────────────────────────────────

    def print_via_gdi(self, printer_name, images, tray_code=1):
        import win32ui
        import ctypes
        import ctypes.wintypes as wintypes

        DM_DEFAULTSOURCE = 0x00000200
        DM_OUT_BUFFER    = 2
        DM_IN_BUFFER     = 8
        IDOK             = 1
        CCHDEVICENAME    = 32

        class DEVMODEW(ctypes.Structure):
            _fields_ = [
                ("dmDeviceName",    ctypes.c_wchar * CCHDEVICENAME),
                ("dmSpecVersion",   ctypes.c_uint16),
                ("dmDriverVersion", ctypes.c_uint16),
                ("dmSize",          ctypes.c_uint16),
                ("dmDriverExtra",   ctypes.c_uint16),
                ("dmFields",        ctypes.c_uint32),
                ("dmOrientation",   ctypes.c_int16),
                ("dmPaperSize",     ctypes.c_int16),
                ("dmPaperLength",   ctypes.c_int16),
                ("dmPaperWidth",    ctypes.c_int16),
                ("dmScale",         ctypes.c_int16),
                ("dmCopies",        ctypes.c_int16),
                ("dmDefaultSource", ctypes.c_int16),
                ("dmPrintQuality",  ctypes.c_int16),
            ]

        winspool = ctypes.WinDLL("winspool.drv")
        gdi32    = ctypes.windll.gdi32
        hDC      = None

        try:
            devmode_buf = None
            try:
                hPrinter = wintypes.HANDLE()
                if winspool.OpenPrinterW(printer_name, ctypes.byref(hPrinter), None):
                    buf_size = winspool.DocumentPropertiesW(0, hPrinter, printer_name, None, None, 0)
                    if buf_size > 0:
                        in_buf = (ctypes.c_byte * buf_size)()
                        if winspool.DocumentPropertiesW(0, hPrinter, printer_name, in_buf, None, DM_OUT_BUFFER) == IDOK:
                            dm = DEVMODEW.from_buffer(in_buf)
                            dm.dmDefaultSource = tray_code
                            dm.dmFields |= DM_DEFAULTSOURCE
                            out_buf = (ctypes.c_byte * buf_size)()
                            if winspool.DocumentPropertiesW(0, hPrinter, printer_name, out_buf, in_buf, DM_IN_BUFFER | DM_OUT_BUFFER) == IDOK:
                                devmode_buf = out_buf
                            else:
                                devmode_buf = in_buf
                    winspool.ClosePrinter(hPrinter)
            except Exception:
                pass

            if devmode_buf is not None:
                hdc_raw = gdi32.CreateDCW("WINSPOOL", printer_name, None, devmode_buf)
                if hdc_raw:
                    hDC = win32ui.CreateDCFromHandle(hdc_raw)
            if hDC is None:
                hDC = win32ui.CreateDC()
                hDC.CreatePrinterDC(printer_name)

            for img in images:
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                hDC.StartDoc("123kast Bon")
                hDC.StartPage()
                try:
                    ox = hDC.GetDeviceCaps(112)
                    oy = hDC.GetDeviceCaps(113)
                except Exception:
                    ox = oy = 0
                from PIL import ImageWin as IW
                IW.Dib(img).draw(hDC.GetHandleOutput(),
                                 (-ox, -oy, img.size[0] - ox, img.size[1] - oy))
                hDC.EndPage()
                hDC.EndDoc()
            return True
        except Exception as e:
            print(f"GDI print error: {e}")
            return False
        finally:
            if hDC is not None:
                try:
                    hDC.DeleteDC()
                except Exception:
                    pass

    def print_direct(self):
        if not self.printer_var.get():
            messagebox.showerror("Fout", "Selecteer eerst een printer.")
            return
        pdf_path = os.path.abspath("123kast_bon_print.pdf")
        if not self._build_pdf(pdf_path):
            return
        try:
            from PIL import Image
            import fitz
            import win32ui
            import win32con
            printer_name = self.printer_var.get()
            tempDC = win32ui.CreateDC()
            tempDC.CreatePrinterDC(printer_name)
            try:
                dpi_x = tempDC.GetDeviceCaps(win32con.LOGPIXELSX)
            except Exception:
                dpi_x = 300
            finally:
                try:
                    tempDC.DeleteDC()
                except Exception:
                    pass
            doc = fitz.open(pdf_path)
            zoom = dpi_x / 72.0
            mat = fitz.Matrix(zoom, zoom)
            images = []
            for p in range(len(doc)):
                pix = doc.load_page(p).get_pixmap(matrix=mat, alpha=False)
                images.append(Image.frombytes("RGB", [pix.width, pix.height], pix.samples))
            doc.close()
        except Exception as e:
            messagebox.showerror("Fout", f"Kon PDF niet renderen: {e}")
            return
        tray_code = self._tray_map.get(self.tray_var.get(), 1)
        if self.print_via_gdi(printer_name, images, tray_code):
            messagebox.showinfo("Succes", f"123kast bon (4 pagina's) verstuurd naar {printer_name}.")
        else:
            messagebox.showerror("Fout", "Kon niet printen.")

    def gen_pdf(self):
        pdf_path = os.path.abspath("123kast_bon.pdf")
        if not self._build_pdf(pdf_path):
            return
        try:
            if os.name == 'nt':
                os.startfile(pdf_path)
            else:
                subprocess.call(['open' if os.name == 'posix' else 'xdg-open', pdf_path])
            messagebox.showinfo("Succes", "123kast bon PDF gegenereerd (4 pagina's)!")
        except Exception as e:
            messagebox.showerror("Fout", f"Kon PDF niet openen: {str(e)}")
