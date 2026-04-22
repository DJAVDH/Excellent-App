import tkinter as tk
from tkinter import ttk, messagebox
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import os
import platform
import subprocess
import win32print
import threading

# ── Shared theme constants (mirrors main.py) ────────────────────────────────
BG          = "#FFFFFF"
BG_FIELD    = "#F7F9FC"
ACCENT      = "#1A1A2E"
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


def _lbl(parent, text, **kw):
    return tk.Label(parent, text=text, bg=BG, fg=TEXT_MUTED,
                    font=(FONT, 9), anchor="w", **kw)


def _entry(parent, width=28, **kw):
    return tk.Entry(parent, width=width,
                    bg=BG_FIELD, fg=TEXT_DARK, insertbackground=TEXT_DARK,
                    relief=tk.FLAT, highlightbackground=BORDER,
                    highlightthickness=1, font=(FONT, 10), **kw)


# Constanten
LABEL_WIDTH = 70 * mm
LABEL_HEIGHT = 37 * mm
LABELS_PER_ROW = 3
LABELS_PER_COL = 8
LABELS_PER_PAGE = LABELS_PER_ROW * LABELS_PER_COL
PAGE_WIDTH, PAGE_HEIGHT = A4


class LabelMakerComponent:
    """Simpele Label Maker component"""

    def __init__(self, parent, app):
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

        for col in range(4):
            form_frame.columnconfigure(col, weight=1 if col % 2 == 1 else 0)

        # ── Row 0: Achternaam / Transport ────────────────────────────────────
        _lbl(form_frame, "Achternaam").grid(row=0, column=0, sticky='w', pady=(0, 2), padx=(0, 8))
        self.achternaam_entry = _entry(form_frame)
        self.achternaam_entry.grid(row=1, column=0, sticky='ew', pady=(0, 12), padx=(0, 16))

        _lbl(form_frame, "Transport").grid(row=0, column=1, sticky='w', pady=(0, 2), padx=(0, 8))
        self.sea_air_var = tk.StringVar(value="Sea")
        self.sea_air_combo = ttk.Combobox(form_frame, textvariable=self.sea_air_var,
                                          values=["Sea", "Air", "Road"],
                                          state="readonly", width=14, font=(FONT, 10))
        self.sea_air_combo.grid(row=1, column=1, sticky='ew', pady=(0, 12), padx=(0, 16))

        # ── Row 2: TO / Referentie ───────────────────────────────────────────
        _lbl(form_frame, "TO").grid(row=2, column=0, sticky='w', pady=(0, 2), padx=(0, 8))
        self.to_entry = _entry(form_frame)
        self.to_entry.grid(row=3, column=0, sticky='ew', pady=(0, 12), padx=(0, 16))

        _lbl(form_frame, "Referentie").grid(row=2, column=1, sticky='w', pady=(0, 2), padx=(0, 8))
        self.ref_entry = _entry(form_frame)
        self.ref_entry.grid(row=3, column=1, sticky='ew', pady=(0, 12), padx=(0, 16))

        # ── Row 4: FM / Startnummer ──────────────────────────────────────────
        _lbl(form_frame, "FM").grid(row=4, column=0, sticky='w', pady=(0, 2), padx=(0, 8))
        self.fm_entry = _entry(form_frame)
        self.fm_entry.insert(0, "NETHERLANDS")
        self.fm_entry.grid(row=5, column=0, sticky='ew', pady=(0, 12), padx=(0, 16))

        _lbl(form_frame, "Startnummer").grid(row=4, column=1, sticky='w', pady=(0, 2), padx=(0, 8))
        self.startnum_entry = _entry(form_frame, width=14)
        self.startnum_entry.insert(0, "1")
        self.startnum_entry.grid(row=5, column=1, sticky='ew', pady=(0, 12), padx=(0, 16))

        # ── Row 6: Aantal pagina's / sticker count ───────────────────────────
        _lbl(form_frame, "Aantal pagina's").grid(row=6, column=0, sticky='w', pady=(0, 2), padx=(0, 8))
        self.pages_var = tk.StringVar(value="1")
        self.pages_entry = _entry(form_frame, width=14, textvariable=self.pages_var)
        self.pages_entry.grid(row=7, column=0, sticky='ew', pady=(0, 12), padx=(0, 16))

        self.sticker_label = tk.Label(form_frame, text="Aantal stickers: 24",
                                      bg=BG, fg=TEXT_MUTED, font=(FONT, 9), anchor="w")
        self.sticker_label.grid(row=7, column=1, sticky='w', pady=(0, 12), padx=(0, 16))

        # ── Row 8: Printer / Papierlade ──────────────────────────────────────
        _lbl(form_frame, "Printer").grid(row=8, column=0, sticky='w', pady=(0, 2), padx=(0, 8))
        self.printer_var = tk.StringVar()
        self.printer_combo = ttk.Combobox(form_frame, textvariable=self.printer_var,
                                          state="readonly", width=28, font=(FONT, 10))
        self.printer_combo.grid(row=9, column=0, sticky='ew', pady=(0, 12), padx=(0, 16))

        _lbl(form_frame, "Papierlade").grid(row=8, column=1, sticky='w', pady=(0, 2), padx=(0, 8))
        self.tray_var = tk.StringVar()
        self._tray_display = ["Lade 1", "Lade 2", "Lade 3 (Stickers)"]
        self._tray_map = {"Lade 1": 1, "Lade 2": 3, "Lade 3 (Stickers)": 2}
        self.tray_combo = ttk.Combobox(form_frame, textvariable=self.tray_var,
                                       values=self._tray_display, state="readonly",
                                       width=14, font=(FONT, 10))
        self.tray_combo.grid(row=9, column=1, sticky='ew', pady=(0, 12), padx=(0, 16))
        self.tray_combo.set("Lade 3 (Stickers)")

        # Update sticker count when pages change
        self.pages_var.trace_add('write', lambda *args: self.update_sticker_count())

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
        tk.Button(btn_bar, text="Leegmaken", command=self.clear_fields,
                  **BTN_SECONDARY).pack(side=tk.LEFT, padx=(0, 8))
    
    def toggle_top(self):
        self.app.root.attributes("-topmost", self.top_var.get())
    
    def update_sticker_count(self):
        try:
            paginas = int(self.pages_var.get())
            totaal = paginas * LABELS_PER_PAGE
            self.sticker_label.config(text=f"Aantal stickers: {totaal}")
        except ValueError:
            self.sticker_label.config(text="Aantal stickers: ?")
    
    def _detect_printers_bg(self):
        try:
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
            printer_names = [printer[2] for printer in printers]
        except Exception:
            printer_names = []
        
        def apply_names():
            try:
                self.printer_combo['values'] = printer_names
                if printer_names and not self.printer_combo.get():
                    canon = next((p for p in printer_names if 'Canon' in p or 'canon' in p), None)
                    if canon:
                        self.printer_combo.set(canon)
                    else:
                        self.printer_combo.set(printer_names[0])
            except Exception:
                pass
        
        try:
            self.app.root.after(0, apply_names)
        except Exception:
            apply_names()
    
    def clear_fields(self):
        self.achternaam_entry.delete(0, tk.END)
        self.sea_air_var.set("Sea")
        self.to_entry.delete(0, tk.END)
        self.ref_entry.delete(0, tk.END)
        self.fm_entry.delete(0, tk.END)
        self.fm_entry.insert(0, "NETHERLANDS")
        self.startnum_entry.delete(0, tk.END)
        self.startnum_entry.insert(0, "1")
        self.pages_entry.delete(0, tk.END)
        self.pages_entry.insert(0, "1")
        self.update_sticker_count()
    
    def gen_pdf(self):
        achternaam = self.achternaam_entry.get().strip()
        sea_air = self.sea_air_var.get()
        to = self.to_entry.get().strip().upper()
        referentie = self.ref_entry.get().strip()
        fm = self.fm_entry.get().strip().upper()

        try:
            startnummer = int(self.startnum_entry.get())
            paginas = int(self.pages_entry.get())
        except ValueError:
            messagebox.showerror("Fout", "Startnummer en aantal pagina's moeten een getal zijn.")
            return

        if paginas < 1 or paginas > 500:
            messagebox.showerror("Fout", "Aantal pagina's moet tussen 1 en 500 zijn.")
            return

        if not achternaam or not to or not referentie or not fm:
            messagebox.showerror("Fout", "Vul alle velden in.")
            return

        total_labels = paginas * LABELS_PER_PAGE
        c = canvas.Canvas("labels.pdf", pagesize=A4)
        
        for i in range(total_labels):
            pos_in_page = i % LABELS_PER_PAGE
            row = pos_in_page // LABELS_PER_ROW
            col = pos_in_page % LABELS_PER_ROW
            
            x = col * LABEL_WIDTH
            y = PAGE_HEIGHT - (0.5 * mm) - ((row + 1) * LABEL_HEIGHT)
            
            nummer = startnummer + i
            nummer_str = f"#{nummer:03d}"
            
            pad_x = 7 * mm
            pad_y_bottom = 8 * mm
            pad_y_top = 9 * mm
            
            text_x = x + pad_x
            text_y = y + LABEL_HEIGHT - pad_y_top
            
            c.setFont("Helvetica", 10)
            c.drawString(text_x, text_y, achternaam)
            
            c.setFont("Helvetica-Bold", 19)
            sea_air_x = x + LABEL_WIDTH - pad_x
            c.drawRightString(sea_air_x, text_y - 1.75, sea_air)
            
            c.setFont("Helvetica", 10)
            c.drawString(text_x, text_y - 13, f"FM: {fm}")
            c.drawString(text_x, text_y - 25, f"TO: {to}")
            
            c.drawString(text_x, y + pad_y_bottom, referentie)
            
            c.setFont("Helvetica-Bold", 19)
            c.drawRightString(x + LABEL_WIDTH - pad_x, y + pad_y_bottom, nummer_str)
            
            if pos_in_page == LABELS_PER_PAGE - 1 and i != total_labels - 1:
                c.showPage()
        
        c.save()
        
        pdf_path = os.path.abspath("labels.pdf")
        if platform.system() == "Windows":
            os.startfile(pdf_path)
        elif platform.system() == "Darwin":
            subprocess.call(["open", pdf_path])
        else:
            subprocess.call(["xdg-open", pdf_path])
        
        messagebox.showinfo("Klaar", f"labels.pdf gegenereerd met {total_labels} labels.")
        
        # ── Statistieken bijhouden ────────────────────────────────────────────
        try:
            if hasattr(self.app, 'increment_labels'):
                self.app.increment_labels(total_labels)
        except Exception:
            pass
    
    def render_labels_pages(self, total_labels, startnummer, achternaam, sea_air, to, referentie, fm, dpi=300):
        from PIL import Image, ImageDraw, ImageFont
        
        A4_MM = (210, 297)
        px_w = int(A4_MM[0] / 25.4 * dpi)
        px_h = int(A4_MM[1] / 25.4 * dpi)
        
        label_w = int((70) / 25.4 * dpi)
        label_h = int((37) / 25.4 * dpi)
        left_margin = int(5 / 25.4 * dpi)
        top_offset = int(2 / 25.4 * dpi)
        
        pages = []
        labels_per_page = LABELS_PER_PAGE
        
        try:
            font_small = ImageFont.truetype("arial.ttf", int(10 / 25.4 * dpi))
            font_large = ImageFont.truetype("arialbd.ttf", int(19 / 25.4 * dpi))
        except Exception:
            font_small = ImageFont.load_default()
            font_large = ImageFont.load_default()
        
        for page_start in range(0, total_labels, labels_per_page):
            img = Image.new("RGB", (px_w, px_h), "white")
            draw = ImageDraw.Draw(img)
            
            for i in range(labels_per_page):
                idx = page_start + i
                if idx >= total_labels:
                    break
                row = i // LABELS_PER_ROW
                col = i % LABELS_PER_ROW
                x = col * label_w + left_margin
                y = row * label_h + top_offset
                nummer = startnummer + idx
                nummer_str = f"#{nummer:03d}"
                draw.text((x + int(2 / 25.4 * dpi), y + int(2 / 25.4 * dpi)), achternaam, font=font_small, fill="black")
                draw.text((x + label_w - int(18 / 25.4 * dpi), y + int(2 / 25.4 * dpi)), sea_air, font=font_large, fill="black")
                draw.text((x + int(2 / 25.4 * dpi), y + int(6 / 25.4 * dpi)), f"FM: {fm}", font=font_small, fill="black")
                draw.text((x + int(2 / 25.4 * dpi), y + int(11 / 25.4 * dpi)), f"TO: {to}", font=font_small, fill="black")
                draw.text((x + int(2 / 25.4 * dpi), y + int(27 / 25.4 * dpi)), referentie, font=font_small, fill="black")
                draw.text((x + label_w - int(12 / 25.4 * dpi), y + int(27 / 25.4 * dpi)), nummer_str, font=font_large, fill="black")
            
            pages.append(img)
        return pages
    
    def print_via_gdi(self, printer_name, images, source_code):
        import win32ui
        import ctypes
        import ctypes.wintypes as wintypes
        from PIL import ImageWin

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
            # Build a DEVMODE buffer with the requested paper tray via DocumentProperties
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
                            dm.dmDefaultSource = int(source_code)
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

            # Create DC with modified DEVMODE so the tray is actually honoured
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

                hDC.StartDoc("Labels")
                hDC.StartPage()

                try:
                    phys_off_x = hDC.GetDeviceCaps(112)
                    phys_off_y = hDC.GetDeviceCaps(113)
                except Exception:
                    phys_off_x = phys_off_y = 0

                bmp_dib = ImageWin.Dib(img)
                bmp_dib.draw(hDC.GetHandleOutput(),
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
        
        achternaam = self.achternaam_entry.get().strip()
        sea_air = self.sea_air_var.get()
        to = self.to_entry.get().strip().upper()
        referentie = self.ref_entry.get().strip()
        fm = self.fm_entry.get().strip().upper()
        
        try:
            startnummer = int(self.startnum_entry.get())
            paginas = int(self.pages_entry.get())
        except ValueError:
            messagebox.showerror("Fout", "Startnummer en aantal pagina's moeten een getal zijn.")
            return
        
        if not achternaam or not to or not referentie or not fm:
            messagebox.showerror("Fout", "Vul alle velden in.")
            return
        
        total_labels = paginas * LABELS_PER_PAGE
        
        printer_name = self.printer_var.get()
        sel = self.tray_var.get()
        tray_number = self._tray_map.get(sel, 1)
        raw_code = tray_number
        
        try:
            import win32ui
            import win32con
            tempDC = win32ui.CreateDC()
            tempDC.CreatePrinterDC(printer_name)
            try:
                dpi_x = tempDC.GetDeviceCaps(win32con.LOGPIXELSX)
                dpi_y = tempDC.GetDeviceCaps(win32con.LOGPIXELSY)
            finally:
                try:
                    tempDC.DeleteDC()
                except Exception:
                    pass
        except Exception:
            dpi_x = dpi_y = 300
        
        pdf_filename = os.path.abspath("labels_print.pdf")
        c = canvas.Canvas(pdf_filename, pagesize=A4)
        for i in range(total_labels):
            pos_in_page = i % LABELS_PER_PAGE
            row = pos_in_page // LABELS_PER_ROW
            col = pos_in_page % LABELS_PER_ROW
            
            x = col * LABEL_WIDTH
            y = PAGE_HEIGHT - (0.5 * mm) - ((row + 1) * LABEL_HEIGHT)
            
            nummer = startnummer + i
            nummer_str = f"#{nummer:03d}"
            
            pad_x = 7 * mm
            pad_y_bottom = 8 * mm
            pad_y_top = 9 * mm
            
            text_x = x + pad_x
            text_y = y + LABEL_HEIGHT - pad_y_top
            
            c.setFont("Helvetica", 10)
            c.drawString(text_x, text_y, achternaam)
            
            c.setFont("Helvetica-Bold", 19)
            sea_air_x = x + LABEL_WIDTH - pad_x
            c.drawRightString(sea_air_x, text_y - 1.75, sea_air)
            
            c.setFont("Helvetica", 10)
            c.drawString(text_x, text_y - 13, f"FM: {fm}")
            c.drawString(text_x, text_y - 25, f"TO: {to}")
            
            c.drawString(text_x, y + pad_y_bottom, referentie)
            
            c.setFont("Helvetica-Bold", 19)
            c.drawRightString(x + LABEL_WIDTH - pad_x, y + pad_y_bottom, nummer_str)
            
            if pos_in_page == LABELS_PER_PAGE - 1 and i != total_labels - 1:
                c.showPage()
        
        c.save()
        
        images = []
        try:
            from PIL import Image
            import fitz
            doc = fitz.open(pdf_filename)
            zoom = dpi_x / 72.0
            mat = fitz.Matrix(zoom, zoom)
            for p in range(len(doc)):
                page = doc.load_page(p)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                images.append(img)
            doc.close()
        except Exception as e:
            messagebox.showerror("Fout", f"Kon PDF niet renderen: {e}")
            return
        
        ok = self.print_via_gdi(printer_name, images, raw_code)
        if ok:
            messagebox.showinfo("Succes", f"Print opdracht naar {printer_name} ({sel}) verstuurd.\n{total_labels} labels")
            try:
                if hasattr(self.app, 'increment_labels'):
                    self.app.increment_labels(total_labels)
            except Exception:
                pass
        else:
            messagebox.showerror("Fout", "Kon niet printen via GDI.")


# Legacy class for backward compatibility
class LabelMakerApp:
    """Legacy standalone Label Maker application"""
    def __init__(self, root):
        self.root = root
        root.title("Excellent Packing and Moving B.V. – Label Generator")
        ico_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "logo.ico")
        if os.path.exists(ico_path):
            root.iconbitmap(ico_path)
        
        # Main container
        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Title
        title_label = tk.Label(main_frame, text="Label Maker", font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Create a simple app-like object for the component
        class SimpleApp:
            def __init__(self, root):
                self.root = root
                self.colors = {}
        
        simple_app = SimpleApp(root)
        self.component = LabelMakerComponent(main_frame, simple_app)
        
        root.geometry("700x500")


# This module can be imported and used as a component