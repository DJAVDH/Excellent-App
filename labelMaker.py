import tkinter as tk
from tkinter import ttk, messagebox
from ttkbootstrap import Style
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import os
import platform
import subprocess
import win32print
import json
import shutil
import time
import win32ui
import win32con
import pywintypes
from PIL import Image, ImageDraw, ImageFont, ImageWin
import threading

# ------ CONSTANTEN VOOR GENERATIE ------
LABEL_WIDTH = 70 * mm
LABEL_HEIGHT = 37 * mm
LABELS_PER_ROW = 3
LABELS_PER_COL = 8
LABELS_PER_PAGE = LABELS_PER_ROW * LABELS_PER_COL
PAGE_WIDTH, PAGE_HEIGHT = A4

# ------ GUI ------
class LabelMakerApp:
    def __init__(self, root):
        self.root = root
        root.title("Excellent Packing and Moving B.V. – Label Generator")
        ico_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), "logo.ico")
        if os.path.exists(ico_path):
            root.iconbitmap(ico_path)
        style = Style("flatly")
        style.configure(".", font=("Helvetica", 11))

        frm = ttk.Frame(root, padding=10)
        frm.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frm, text="Achternaam:").grid(row=0, column=0, sticky='w')
        self.achternaam_entry = ttk.Entry(frm)
        self.achternaam_entry.grid(row=0, column=1, sticky='ew')

        ttk.Label(frm, text="Transport:").grid(row=1, column=0, sticky='w')
        self.sea_air_var = tk.StringVar(value="Sea")
        self.sea_air_menu = ttk.Combobox(frm, textvariable=self.sea_air_var, values=["Sea", "Air", "Road"], state="readonly")
        self.sea_air_menu.grid(row=1, column=1, sticky='w')

        ttk.Label(frm, text="TO:").grid(row=2, column=0, sticky='w')
        self.to_entry = ttk.Entry(frm)
        self.to_entry.grid(row=2, column=1, sticky='ew')

        ttk.Label(frm, text="Referentie:").grid(row=3, column=0, sticky='w')
        self.ref_entry = ttk.Entry(frm)
        self.ref_entry.grid(row=3, column=1, sticky='ew')

        ttk.Label(frm, text="FM:").grid(row=4, column=0, sticky='w')
        self.fm_entry = ttk.Entry(frm)
        self.fm_entry.insert(0, "NETHERLANDS")
        self.fm_entry.grid(row=4, column=1, sticky='ew')

        ttk.Label(frm, text="Startnummer (bv. 1 = #001):").grid(row=5, column=0, sticky='w')
        self.startnum_entry = ttk.Entry(frm)
        self.startnum_entry.insert(0, "1")
        self.startnum_entry.grid(row=5, column=1, sticky='ew')

        ttk.Label(frm, text="Aantal pagina's:").grid(row=6, column=0, sticky='w')
        self.pages_var = tk.StringVar(value="1")
        self.pages_entry = ttk.Entry(frm, textvariable=self.pages_var)
        self.pages_entry.grid(row=6, column=1, sticky='ew')

        self.sticker_label = ttk.Label(frm, text="Aantal stickers: 24")
        self.sticker_label.grid(row=7, column=0, columnspan=2, sticky='w')

        self.pages_var.trace_add('write', lambda *args: self.update_sticker_count())

        ttk.Label(frm, text="Printer:").grid(row=8, column=0, sticky='w')
        self.printer_var = tk.StringVar(value="")
        self.printer_menu = ttk.Combobox(frm, textvariable=self.printer_var, state="readonly")
        self.printer_menu.grid(row=8, column=1, sticky='ew')
        self.printer_var.trace_add('write', lambda *args: self.load_config())
        threading.Thread(target=self._detect_printers_bg, daemon=True).start()

        ttk.Label(frm, text="Papierlade:").grid(row=9, column=0, sticky='w')
        self.tray_var = tk.StringVar()
        self._tray_display = ["1", "Stickers (3)", "3", "4"]
        self._tray_map = {"1":1, "Stickers (3)":2, "3":3, "4":4}
        self.tray_menu = ttk.Combobox(frm, textvariable=self.tray_var, values=self._tray_display, state="readonly")
        self.tray_menu.grid(row=9, column=1, sticky='w')
        self.tray_menu.set("Stickers (3)") 

        self.top_var = tk.BooleanVar(value=False)
        self.top_check = ttk.Checkbutton(frm, text="Venster altijd vooraan", variable=self.top_var, command=self.toggle_top)
        self.top_check.grid(row=10, column=0, columnspan=2, sticky='w', pady=(5, 5))

        self.gen_btn = ttk.Button(frm, text="Genereer PDF", command=self.gen_pdf)
        self.gen_btn.grid(row=11, column=0, pady=(10, 0), sticky='ew')

        self.print_btn = ttk.Button(frm, text="Direct Printen", command=self.print_direct)
        self.print_btn.grid(row=11, column=1, pady=(10, 0), sticky='ew')

        self.clear_btn = ttk.Button(frm, text="Leegmaken", command=self.clear_fields, style="Danger.TButton")
        self.clear_btn.grid(row=12, column=1, pady=(10, 0), sticky='ew')

        frm.columnconfigure(1, weight=1)

        root.resizable(False, False)
        root.update_idletasks()
        root.geometry(f"{root.winfo_reqwidth()}x{root.winfo_reqheight()}")

    def toggle_top(self):
        self.root.attributes("-topmost", self.top_var.get())

    def refresh_printers(self):
        threading.Thread(target=self._detect_printers_bg, daemon=True).start()

    def _detect_printers_bg(self):
        try:
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
            printer_names = [printer[2] for printer in printers]
        except Exception:
            printer_names = []

        def apply_names():
            try:
                self.printer_menu['values'] = printer_names
                if printer_names and not self.printer_menu.get():
                    canon = next((p for p in printer_names if 'Canon' in p or 'canon' in p), None)
                    if canon:
                        self.printer_menu.set(canon)
                    else:
                        self.printer_menu.set(printer_names[0])
            except Exception:
                pass

        try:
            self.root.after(0, apply_names)
        except Exception:
            apply_names()

    def show_printer_info(self):
        messagebox.showinfo("Info", "Printer-informatie functie verwijderd.")

    def show_printer_props(self):
        messagebox.showinfo("Info", "Printer eigenschappen functie verwijderd.")

    def load_config(self):
        return

    def save_config(self):
        return

    def clear_fields(self):
        self.achternaam_entry.delete(0, tk.END)
        self.sea_air_var.set("Sea")
        self.to_entry.delete(0, tk.END)
        self.ref_entry.delete(0, tk.END)
        self.startnum_entry.delete(0, tk.END)
        self.startnum_entry.insert(0, "1")
        self.pages_entry.delete(0, tk.END)
        self.pages_entry.insert(0, "1")
        self.update_sticker_count()

    def update_sticker_count(self):
        try:
            paginas = int(self.pages_var.get())
            totaal = paginas * LABELS_PER_PAGE
            self.sticker_label.config(text=f"Aantal stickers: {totaal}")
        except ValueError:
            self.sticker_label.config(text="Aantal stickers: ?")

    # ------ GENERATIE ------
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

            # NIEUW: Veilige marges van 7mm aan zijkanten, 8mm onder, en nu 9mm vanaf de bovenkant!
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

    def render_labels_pages(self, total_labels, startnummer, achternaam, sea_air, to, referentie, fm, dpi=300):
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
        try:
            try:
                hPrinter = win32print.OpenPrinter(printer_name)
                pinfo = win32print.GetPrinter(hPrinter, 2)
                devmode = pinfo.get('pDevMode')
                if devmode is not None:
                    try:
                        devmode.dmDefaultSource = int(source_code)
                    except Exception:
                        pass
                    try:
                        devmode.DefaultSource = int(source_code)
                    except Exception:
                        pass
                    try:
                        if hasattr(devmode, 'dmFields'):
                            dm_defsrc = getattr(win32con, 'DM_DEFAULTSOURCE', 0)
                            devmode.dmFields = int(getattr(devmode, 'dmFields', 0)) | dm_defsrc
                    except Exception:
                        pass
                    pinfo['pDevMode'] = devmode
                    try:
                        win32print.SetPrinter(hPrinter, 2, pinfo)
                    except Exception:
                        pass
                win32print.ClosePrinter(hPrinter)
            except Exception:
                pass

            hDC = win32ui.CreateDC()
            hDC.CreatePrinterDC(printer_name)

            for idx, img in enumerate(images):
                if img.mode != 'RGB':
                    bmp = img.convert('RGB')
                else:
                    bmp = img

                hDC.StartDoc(f"Labels_{idx}")
                hDC.StartPage()

                try:
                    dpi_x = hDC.GetDeviceCaps(win32con.LOGPIXELSX)
                    dpi_y = hDC.GetDeviceCaps(win32con.LOGPIXELSY)
                except Exception:
                    dpi_x = dpi_y = 300

                try:
                    phys_off_x = hDC.GetDeviceCaps(112)
                    phys_off_y = hDC.GetDeviceCaps(113)
                except Exception:
                    phys_off_x = 0
                    phys_off_y = 0

                bmp_dib = ImageWin.Dib(bmp)
                
                dest_x1 = -phys_off_x
                dest_y1 = -phys_off_y
                dest_x2 = bmp.size[0] - phys_off_x
                dest_y2 = bmp.size[1] - phys_off_y
                
                bmp_dib.draw(hDC.GetHandleOutput(), (dest_x1, dest_y1, dest_x2, dest_y2))

                hDC.EndPage()
                hDC.EndDoc()

            try:
                hDC.DeleteDC()
            except Exception:
                pass

            return True
        except Exception as e:
            print(f"GDI print error: {e}")
            return False

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

            # NIEUW: Veilige marges van 7mm aan zijkanten, 8mm onder, en nu 9mm vanaf de bovenkant!
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
            messagebox.showinfo("Succes", f"Print opdracht naar {printer_name} (lade {tray_number}) verstuurd.\n{total_labels} labels")
        else:
            messagebox.showerror("Fout", "Kon niet printen via GDI.")

if __name__ == "__main__":
    root = tk.Tk()
    app = LabelMakerApp(root)
    root.mainloop()