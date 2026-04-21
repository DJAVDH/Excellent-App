import tkinter as tk
from tkinter import ttk, messagebox
import os
import subprocess
import threading
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
import win32print

# ── Shared theme constants (mirrors main.py) ────────────────────────────────
BG          = "#FFFFFF"
BG_FIELD    = "#F7F9FC"
ACCENT      = "#1A1A2E"
GREEN       = "#4CAF50"
TEXT_DARK   = "#1A1A2E"
TEXT_MUTED  = "#8A8FA8"
BORDER      = "#E8EAF0"
FONT        = "Segoe UI"

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
    e = tk.Entry(parent, textvariable=textvariable, width=width,
                 bg=BG_FIELD, fg=TEXT_DARK, insertbackground=TEXT_DARK,
                 relief=tk.FLAT, highlightbackground=BORDER,
                 highlightthickness=1, font=(FONT, 10))
    return e


class AssistentieBonComponent:
    """Simpele Assistentie Bon component"""

    def __init__(self, parent, app=None):
        self.parent = parent
        self.app = app

        # Main frame
        self.frame = tk.Frame(parent, bg=BG)
        self.frame.pack(fill=tk.BOTH, expand=True)

        # ── Form card ────────────────────────────────────────────────────────
        card = tk.Frame(self.frame, bg=BG,
                        highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True)

        form_frame = tk.Frame(card, bg=BG)
        form_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=16)

        # Configure columns
        for col in range(4):
            form_frame.columnconfigure(col, weight=1 if col % 2 == 1 else 0)

        # ── Variables ────────────────────────────────────────────────────────
        self.datum_var        = tk.StringVar()
        self.betreft_var      = tk.StringVar()
        self.wagen_var        = tk.StringVar()
        self.werkzaamheden_var = tk.StringVar()
        self.adres_var        = tk.StringVar()
        self.tijd_var         = tk.StringVar()
        self.contactpersoon_var = tk.StringVar()
        self.namen_var        = tk.StringVar()
        self.bedrijf_var      = tk.StringVar()

        # ── Row 0: Datum / Betreft ───────────────────────────────────────────
        _label(form_frame, "Datum").grid(row=0, column=0, sticky='w', pady=(0, 2), padx=(0, 8))
        self.datum_entry = _entry(form_frame, self.datum_var)
        self.datum_entry.grid(row=1, column=0, sticky='ew', pady=(0, 12), padx=(0, 16))

        _label(form_frame, "Betreft").grid(row=0, column=1, sticky='w', pady=(0, 2), padx=(0, 8))
        self.betreft_entry = _entry(form_frame, self.betreft_var)
        self.betreft_entry.grid(row=1, column=1, sticky='ew', pady=(0, 12), padx=(0, 16))

        _label(form_frame, "Wagen").grid(row=0, column=2, sticky='w', pady=(0, 2), padx=(0, 8))
        wagen_options = ["EV", "WG1", "WG2", "WG3", "WG4", "WG5", "WG6",
                         "WG7", "WG8", "WG9", "WG10", "WG11", "WG12",
                         "WG13", "WG14", "WG15", "Huur Bus"]
        self.wagen_var.set("EV")
        self.wagen_combo = ttk.Combobox(form_frame, textvariable=self.wagen_var,
                                        values=wagen_options, width=14,
                                        font=(FONT, 10))
        self.wagen_combo.grid(row=1, column=2, sticky='ew', pady=(0, 12), padx=(0, 16))

        _label(form_frame, "Bedrijf").grid(row=0, column=3, sticky='w', pady=(0, 2))
        self.bedrijf_entry = _entry(form_frame, self.bedrijf_var)
        self.bedrijf_entry.grid(row=1, column=3, sticky='ew', pady=(0, 12))

        # ── Row 2: Werkzaamheden ─────────────────────────────────────────────
        _label(form_frame, "Werkzaamheden").grid(row=2, column=0, sticky='w', pady=(0, 2))
        self.werkzaamheden_entry = _entry(form_frame, self.werkzaamheden_var, width=60)
        self.werkzaamheden_entry.grid(row=3, column=0, columnspan=4, sticky='ew', pady=(0, 12))

        # ── Row 4: Adres ─────────────────────────────────────────────────────
        _label(form_frame, "Adres").grid(row=4, column=0, sticky='w', pady=(0, 2))
        self.adres_entry = _entry(form_frame, self.adres_var, width=60)
        self.adres_entry.grid(row=5, column=0, columnspan=4, sticky='ew', pady=(0, 12))

        # ── Row 6: Tijd / Contactpersoon ─────────────────────────────────────
        _label(form_frame, "Tijd").grid(row=6, column=0, sticky='w', pady=(0, 2), padx=(0, 8))
        self.tijd_entry = _entry(form_frame, self.tijd_var, width=14)
        self.tijd_entry.grid(row=7, column=0, sticky='ew', pady=(0, 12), padx=(0, 16))

        _label(form_frame, "Contactpersoon").grid(row=6, column=1, sticky='w', pady=(0, 2), padx=(0, 8))
        self.contactpersoon_entry = _entry(form_frame, self.contactpersoon_var)
        self.contactpersoon_entry.grid(row=7, column=1, columnspan=3, sticky='ew', pady=(0, 12))

        # ── Row 8: Namen ─────────────────────────────────────────────────────
        _label(form_frame, "Namen").grid(row=8, column=0, sticky='w', pady=(0, 2))
        self.namen_entry = _entry(form_frame, self.namen_var, width=60)
        self.namen_entry.grid(row=9, column=0, columnspan=4, sticky='ew', pady=(0, 12))

        # ── Row 10: Notities ─────────────────────────────────────────────────
        _label(form_frame, "Notities").grid(row=10, column=0, sticky='w', pady=(0, 2))
        self.notities_text = tk.Text(form_frame, height=4, width=60,
                                     bg=BG_FIELD, fg=TEXT_DARK, insertbackground=TEXT_DARK,
                                     relief=tk.FLAT, highlightbackground=BORDER,
                                     highlightthickness=1, font=(FONT, 10), wrap=tk.WORD)
        self.notities_text.grid(row=11, column=0, columnspan=4, sticky='ew', pady=(0, 12))

        # ── Row 12: Printer / Papierlade ─────────────────────────────────────
        _label(form_frame, "Printer").grid(row=12, column=0, sticky='w', pady=(0, 2), padx=(0, 8))
        self.printer_var = tk.StringVar()
        self.printer_combo = ttk.Combobox(form_frame, textvariable=self.printer_var,
                                          state="readonly", width=28, font=(FONT, 10))
        self.printer_combo.grid(row=13, column=0, columnspan=3, sticky='ew', pady=(0, 12), padx=(0, 16))

        _label(form_frame, "Papierlade").grid(row=12, column=3, sticky='w', pady=(0, 2))
        self._tray_map = {"Lade 1": 1, "Lade 2": 3, "Lade 3 (Stickers)": 2}
        self.tray_var = tk.StringVar()
        self.tray_combo = ttk.Combobox(form_frame, textvariable=self.tray_var,
                                       values=list(self._tray_map.keys()), state="readonly",
                                       width=14, font=(FONT, 10))
        self.tray_combo.grid(row=13, column=3, sticky='ew', pady=(0, 12))
        self.tray_combo.set("Lade 1")

        # Set defaults
        self.set_defaults()

        # Start printer detection in background
        threading.Thread(target=self._detect_printers_bg, daemon=True).start()

        # ── Divider ──────────────────────────────────────────────────────────
        tk.Frame(card, bg=BORDER, height=1).pack(fill=tk.X)

        # ── Buttons bar ──────────────────────────────────────────────────────
        btn_bar = tk.Frame(card, bg=BG)
        btn_bar.pack(fill=tk.X, padx=20, pady=14)

        tk.Button(btn_bar, text="Genereer PDF", command=self.gen_pdf,
                  **BTN_STYLE).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_bar, text="Direct Printen", command=self.print_direct,
                  **BTN_STYLE).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_bar, text="Velden Leegmaken", command=self.clear_all_fields,
                  **BTN_SECONDARY).pack(side=tk.LEFT, padx=(0, 8))

    # ── Logic (unchanged) ────────────────────────────────────────────────────

    def set_defaults(self):
        tomorrow = (datetime.now() + timedelta(days=1)).strftime("%d-%m-%Y")
        self.datum_var.set(tomorrow)
        self.betreft_var.set("ASS")
        self.wagen_var.set("EV")
        self.tijd_var.set("08:00")
        self.bedrijf_var.set("")
        self.werkzaamheden_var.set("")
        self.adres_var.set("")
        self.contactpersoon_var.set("")
        self.namen_var.set("")

    def clear_all_fields(self):
        """Wist alleen de inhoudsvelden; datum, betreft, wagen en tijd blijven behouden."""
        self.werkzaamheden_var.set("")
        self.adres_var.set("")
        self.bedrijf_var.set("")
        self.contactpersoon_var.set("")
        self.namen_var.set("")
        self.notities_text.delete("1.0", tk.END)

    def clear_fields(self):
        self.werkzaamheden_var.set("")
        self.adres_var.set("")
        self.contactpersoon_var.set("")
        self.namen_var.set("")
        self.bedrijf_var.set("")
        self.notities_text.delete("1.0", tk.END)

    def go_back(self):
        if self.app:
            self.app.show_main_menu()

    def _detect_printers_bg(self):
        try:
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
            printer_names = [p[2] for p in printers]
        except Exception:
            printer_names = []

        def apply():
            try:
                self.printer_combo['values'] = printer_names
                if printer_names and not self.printer_combo.get():
                    canon = next((p for p in printer_names if 'canon' in p.lower()), None)
                    self.printer_combo.set(canon or printer_names[0])
            except Exception:
                pass

        try:
            self.app.root.after(0, apply)
        except Exception:
            apply()

    def _build_pdf(self, path):
        datum          = self.datum_var.get().strip()
        betreft        = self.betreft_var.get().strip()
        wagen          = self.wagen_var.get().strip()
        werkzaamheden  = self.werkzaamheden_var.get().strip()
        adres          = self.adres_var.get().strip()
        tijd           = self.tijd_var.get().strip()
        contactpersoon = self.contactpersoon_var.get().strip()
        namen          = self.namen_var.get().strip()
        bedrijf        = self.bedrijf_var.get().strip()
        notities       = self.notities_text.get("1.0", tk.END).strip()

        if not datum or not betreft or not wagen or not werkzaamheden or not adres or not tijd or not contactpersoon:
            messagebox.showerror("Fout", "Vul alle verplichte velden in.")
            return False

        c = canvas.Canvas(path, pagesize=A4)
        W, H = A4
        M = 15 * mm
        y = H - 15 * mm

        logo_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "logobanner.png")
        logo_h = 20 * mm
        logo_w = 0
        if os.path.exists(logo_path):
            try:
                from PIL import Image as PILImage
                img = PILImage.open(logo_path)
                orig_w, orig_h = img.size
                logo_w = (orig_w / orig_h) * logo_h
                c.drawImage(logo_path, M, y - logo_h, width=logo_w, height=logo_h,
                            preserveAspectRatio=True, mask='auto')
            except Exception:
                logo_w = 0

        c.setFont("Helvetica", 9)
        c.setFillColor(colors.HexColor("#8A8FA8"))
        c.drawString(M, y - logo_h - 4 * mm, "t.a.v. planning")
        c.setFont("Helvetica-Bold", 18)
        c.setFillColor(colors.HexColor("#1A1A2E"))
        c.drawRightString(W - M, y - logo_h / 2 - 3 * mm, "Assistentie Bon")
        y -= logo_h + 10 * mm

        c.setStrokeColor(colors.HexColor("#E8EAF0"))
        c.setLineWidth(1)
        c.line(M, y, W - M, y)
        y -= 8 * mm

        label_w = 42 * mm
        value_w = (W - 2 * M - label_w * 2) / 2
        row_h   = 7 * mm

        def info_row(label1, val1, label2="", val2=""):
            nonlocal y
            c.setFont("Helvetica-Bold", 9)
            c.setFillColor(colors.HexColor("#8A8FA8"))
            c.drawString(M, y, label1)
            c.setFont("Helvetica", 10)
            c.setFillColor(colors.HexColor("#1A1A2E"))
            c.drawString(M + label_w, y, val1)
            if label2:
                c.setFont("Helvetica-Bold", 9)
                c.setFillColor(colors.HexColor("#8A8FA8"))
                c.drawString(M + label_w + value_w + 4 * mm, y, label2)
                c.setFont("Helvetica", 10)
                c.setFillColor(colors.HexColor("#1A1A2E"))
                c.drawString(M + label_w + value_w + 4 * mm + label_w, y, val2)
            y -= row_h

        info_row("Datum:", datum, "Betreft:", betreft)
        info_row("Wagen:", wagen, "Tijd:", tijd)
        info_row("Contactpersoon:", contactpersoon)
        info_row("Werkzaamheden:", werkzaamheden)
        info_row("Adres:", adres)
        y -= 4 * mm

        c.setStrokeColor(colors.HexColor("#E8EAF0"))
        c.line(M, y, W - M, y)
        y -= 8 * mm

        col_headers = ["Namen", "Aanwezig\nVan:", "Aanwezig\nTot:", "Pauzetijd",
                       "Handtekening", "Beoordeling"]
        col_w = [45*mm, 22*mm, 22*mm, 20*mm, 38*mm, 28*mm]

        if namen:
            naam_lijst = [n.strip() for n in (namen.split(",") if "," in namen else namen.split()) if n.strip()]
        else:
            naam_lijst = []

        data_rows = [[naam, "", "", "", "", ""] for naam in naam_lijst]
        while len(data_rows) < 5:
            data_rows.append(["", "", "", "", "", ""])

        tbl = Table([col_headers] + data_rows, colWidths=col_w,
                    rowHeights=[10*mm] + [9*mm] * len(data_rows))
        tbl.setStyle(TableStyle([
            ('BACKGROUND',    (0, 0), (-1, 0),  colors.HexColor("#1A1A2E")),
            ('TEXTCOLOR',     (0, 0), (-1, 0),  colors.white),
            ('FONTNAME',      (0, 0), (-1, 0),  'Helvetica-Bold'),
            ('FONTSIZE',      (0, 0), (-1, 0),  8),
            ('ALIGN',         (0, 0), (-1, 0),  'CENTER'),
            ('VALIGN',        (0, 0), (-1, 0),  'MIDDLE'),
            ('FONTNAME',      (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE',      (0, 1), (-1, -1), 9),
            ('ALIGN',         (0, 1), (-1, -1), 'LEFT'),
            ('VALIGN',        (0, 1), (-1, -1), 'MIDDLE'),
            ('ROWBACKGROUNDS',(0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F9FC")]),
            ('GRID',          (0, 0), (-1, -1), 0.5, colors.HexColor("#E8EAF0")),
            ('BOX',           (0, 0), (-1, -1), 1,   colors.HexColor("#1A1A2E")),
            ('LEFTPADDING',   (0, 0), (-1, -1), 4),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 4),
            ('TOPPADDING',    (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        tbl_w, tbl_h = tbl.wrapOn(c, W - 2*M, H)
        tbl.drawOn(c, M, y - tbl_h)
        y -= tbl_h + 10 * mm

        dots = "." * 42

        def footer_line(label, value=""):
            nonlocal y
            c.setFont("Helvetica-Bold", 9)
            c.setFillColor(colors.HexColor("#1A1A2E"))
            c.drawString(M, y, label)
            c.setFont("Helvetica", 10)
            c.drawString(M + 52 * mm, y, value if value else dots)
            y -= 8 * mm

        footer_line("Parkeergeld:")
        footer_line("Aantal gereden km:")
        footer_line("Bedrijf:", bedrijf)
        y -= 4 * mm
        c.setFont("Helvetica-Bold", 9)
        c.setFillColor(colors.HexColor("#8A8FA8"))
        c.drawString(M, y, "Af laten tekenen door voorman")
        y -= 8 * mm
        footer_line("Naam voorman:")
        footer_line("Paraaf:")

        y -= 4 * mm
        notities_h = 28 * mm
        c.setStrokeColor(colors.HexColor("#1A1A2E"))
        c.setLineWidth(1)
        c.rect(M, y - notities_h, W - 2 * M, notities_h)
        c.setFont("Helvetica-Bold", 9)
        c.setFillColor(colors.HexColor("#8A8FA8"))
        c.drawString(M + 3 * mm, y - 5 * mm, "Notities:")
        if notities:
            text_obj = c.beginText(M + 3 * mm, y - 13 * mm)
            text_obj.setFont("Helvetica", 9)
            text_obj.setLeading(12)
            for line in notities.splitlines():
                text_obj.textLine(line)
            c.drawText(text_obj)

        c.setFont("Helvetica", 8)
        c.setFillColor(colors.HexColor("#8A8FA8"))
        c.drawCentredString(W / 2, 10 * mm,
                            "Excellent Packing and Moving B.V.  •  Assistentie Bon")
        c.save()
        return True

    def print_via_gdi(self, printer_name, images, tray_code=1):
        import win32ui
        import ctypes
        import ctypes.wintypes as wintypes
        from PIL import ImageWin

        DM_DEFAULTSOURCE = 0x00000200
        DM_OUT_BUFFER    = 2
        DM_IN_BUFFER     = 8
        IDOK             = 1
        CCHDEVICENAME    = 32
        TRAY_CODE        = tray_code

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
                    buf_size = winspool.DocumentPropertiesW(
                        0, hPrinter, printer_name, None, None, 0)
                    if buf_size > 0:
                        in_buf = (ctypes.c_byte * buf_size)()
                        if winspool.DocumentPropertiesW(
                                0, hPrinter, printer_name, in_buf, None,
                                DM_OUT_BUFFER) == IDOK:
                            dm = DEVMODEW.from_buffer(in_buf)
                            dm.dmDefaultSource = TRAY_CODE
                            dm.dmFields |= DM_DEFAULTSOURCE
                            out_buf = (ctypes.c_byte * buf_size)()
                            if winspool.DocumentPropertiesW(
                                    0, hPrinter, printer_name, out_buf, in_buf,
                                    DM_IN_BUFFER | DM_OUT_BUFFER) == IDOK:
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
                hDC.StartDoc("AssistentieBon")
                hDC.StartPage()
                try:
                    phys_off_x = hDC.GetDeviceCaps(112)
                    phys_off_y = hDC.GetDeviceCaps(113)
                except Exception:
                    phys_off_x = phys_off_y = 0
                from PIL import ImageWin as IW
                IW.Dib(img).draw(hDC.GetHandleOutput(),
                                 (-phys_off_x, -phys_off_y,
                                  img.size[0] - phys_off_x,
                                  img.size[1] - phys_off_y))
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

        pdf_path = os.path.abspath("assistentie_bon_print.pdf")
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

        sel = self.tray_var.get()
        tray_code = self._tray_map.get(sel, 1)
        ok = self.print_via_gdi(printer_name, images, tray_code)
        if ok:
            messagebox.showinfo("Succes", f"Assistentie bon verstuurd naar {printer_name} ({sel}).")
            try:
                if hasattr(self.app, 'increment_bons'):
                    self.app.increment_bons()
            except Exception:
                pass
        else:
            messagebox.showerror("Fout", "Kon niet printen.")

    def gen_pdf(self):
        pdf_path = os.path.abspath("assistentie_bon.pdf")
        if not self._build_pdf(pdf_path):
            return
        try:
            if os.name == 'nt':
                os.startfile(pdf_path)
            else:
                subprocess.call(['open', pdf_path] if os.name == 'posix' else ['xdg-open', pdf_path])
            messagebox.showinfo("Succes", "Assistentie bon PDF gegenereerd!")
            try:
                if hasattr(self.app, 'increment_bons'):
                    self.app.increment_bons()
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("Fout", f"Kon PDF niet openen: {str(e)}")


# Legacy class for backward compatibility
class AssistentieBonApp:
    """Legacy standalone Assistentie Bon application"""
    def __init__(self, root):
        self.root = root
        root.title("Assistentie Bon - Excellent Packing and Moving B.V.")
        ico_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "logo.ico")
        if os.path.exists(ico_path):
            root.iconbitmap(ico_path)

        main_frame = tk.Frame(root, bg=BG)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tk.Label(main_frame, text="Assistentie Bon", bg=BG, fg=TEXT_DARK,
                 font=(FONT, 14, "bold")).pack(pady=(0, 10))

        class SimpleApp:
            def __init__(self, root):
                self.root = root
                self.colors = {}
            def show_main_menu(self):
                pass

        simple_app = SimpleApp(root)
        self.component = AssistentieBonComponent(main_frame, simple_app)
        root.geometry("700x560")


# This module can be imported and used as a component
