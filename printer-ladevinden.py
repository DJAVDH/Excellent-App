import json
import time
import win32print
import win32ui
import win32con
import pywintypes


def list_bins(printer_name):
    """Return (bins, bin_names) for given printer.

    Falls back to a numeric range when DeviceCapabilities constants are missing
    or the driver doesn't expose bin names.
    """
    try:
        cap_bins = getattr(win32print, 'DC_BINS', None) or getattr(win32con, 'DC_BINS', None)
        cap_binnames = getattr(win32print, 'DC_BINNAMES', None) or getattr(win32con, 'DC_BINNAMES', None)
        if cap_bins is None or cap_binnames is None:
            raise AttributeError('DC_BINS/DC_BINNAMES not available')

        bins = win32print.DeviceCapabilities(printer_name, "", cap_bins)
        bin_names = win32print.DeviceCapabilities(printer_name, "", cap_binnames)
        # bin_names may be a sequence of bytes; convert and strip
        names = []
        for n in bin_names:
            if isinstance(n, bytes):
                names.append(n.decode('utf-8', errors='ignore').strip('\x00').strip())
            else:
                names.append(str(n).strip())
        return bins, names
    except Exception:
        # Fallback: many drivers don't expose bin names via DeviceCapabilities in pywin32.
        # Offer a sensible numeric range for manual testing so the user can try common IDs.
        fallback_bins = list(range(1, 21))
        fallback_names = ["(naam onbekend)" for _ in fallback_bins]
        return fallback_bins, fallback_names


def set_printer_default_source(printer_name, bin_id):
    """Set DefaultSource on the printer's DEVMODE. Returns original pDevMode (to restore).

    Note: This sets the printer info (PRINTER_INFO_2) temporarily. You should restore
    the original settings after the test.
    """
    # Try default open first (may fail if not enough privileges)
    hPrinter = None
    try:
        hPrinter = win32print.OpenPrinter(printer_name)
        info = win32print.GetPrinter(hPrinter, 2)
        original_pdm = info['pDevMode']
        pdm = info['pDevMode']
        try:
            pdm.DefaultSource = int(bin_id)
        except Exception:
            pass
        try:
            pdm.dmDefaultSource = int(bin_id)
        except Exception:
            pass
        # write back
        info['pDevMode'] = pdm
        try:
            win32print.SetPrinter(hPrinter, 2, info, 0)
        except pywintypes.error as e:
            # If access denied, try reopening with higher desired access and retry
            if getattr(e, 'winerror', None) == 5:
                try:
                    if hPrinter:
                        win32print.ClosePrinter(hPrinter)
                        hPrinter = None
                    hPrinter = win32print.OpenPrinter(printer_name, {"DesiredAccess": win32print.PRINTER_ALL_ACCESS})
                    win32print.SetPrinter(hPrinter, 2, info, 0)
                except Exception as e2:
                    raise
            else:
                raise
        # small delay to allow driver to pick up new setting
        time.sleep(0.5)
        return original_pdm
    finally:
        if hPrinter:
            try:
                win32print.ClosePrinter(hPrinter)
            except Exception:
                pass


def restore_printer_pdevmode(printer_name, original_pdm):
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        info = win32print.GetPrinter(hPrinter, 2)
        info['pDevMode'] = original_pdm
        win32print.SetPrinter(hPrinter, 2, info, 0)
        time.sleep(0.5)
    finally:
        win32print.ClosePrinter(hPrinter)


def print_test_page(printer_name, message="Test pagina - kies lade" ):
    """Print een eenvoudige tekstpagina via GDI naar `printer_name`."""
    dc = win32ui.CreateDC()
    try:
        dc.CreatePrinterDC(printer_name)
        dc.StartDoc('Tray test')
        dc.StartPage()
        # Maak en selecteer een font
        lf = win32ui.LOGFONT()
        lf.lfFaceName = 'Arial'
        lf.lfHeight = 200  # points-ish; pas aan bij nood
        font = win32ui.CreateFont(lf)
        dc.SelectObject(font)
        dc.TextOut(100, 100, message)
        dc.EndPage()
        dc.EndDoc()
    finally:
        try:
            dc.DeleteDC()
        except Exception:
            pass


def save_mapping(printer_name, bin_id, filename='tray_mappings.json'):
    try:
        data = {}
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception:
            data = {}
        data[printer_name] = int(bin_id)
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)
        print(f"Opgeslagen {bin_id} voor printer '{printer_name}' in {filename}")
    except Exception as e:
        print('Kon mapping niet opslaan:', e)


def interactive():
    print('Enumerating printers...')
    printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
    for i, p in enumerate(printers):
        print(f"{i}: {p}")
    idx = input('Kies printer index of typ naam (enter = eerste): ').strip()
    if idx == '':
        printer_name = printers[0]
    else:
        if idx.isdigit():
            printer_name = printers[int(idx)]
        else:
            printer_name = idx

    try:
        bins, names = list_bins(printer_name)
    except Exception as e:
        print('Kon lades niet ophalen:', e)
        return

    print('\nBeschikbare lades:')
    print('-'*40)
    for b, n in zip(bins, names):
        print(f"ID: {b}   |  Naam: {n}")

    choice = input('\nTyp het ID dat je wilt testen (of "save <id>" om op te slaan, of "quit"): ').strip()
    if choice.lower().startswith('save '):
        parts = choice.split()
        if len(parts) >= 2:
            save_mapping(printer_name, parts[1])
        return
    if choice.lower() in ('q', 'quit', 'exit'):
        return

    try:
        bin_id = int(choice)
    except Exception:
        print('Ongeldige invoer.')
        return

    confirm = input(f"We gaan printer '{printer_name}' tijdelijk op lade {bin_id} zetten en een testpagina printen. Doorgaan? (y/n) ").strip().lower()
    if confirm != 'y':
        print('Geannuleerd')
        return

    print('Lezen en wijzigen van printerinstellingen...')
    try:
        original = set_printer_default_source(printer_name, bin_id)
        print('Instelling toegepast. Printen testpagina...')
        print_test_page(printer_name, f'Testpagina, probeerlade {bin_id}')
        print('Testpagina verzonden. Controleer van welke lade het papier komt.')
        saveq = input('Is dit de juiste lade? Wil je deze code bewaren voor deze printer? (y/n) ').strip().lower()
        if saveq == 'y':
            save_mapping(printer_name, bin_id)
    except Exception as e:
        print('Fout tijdens test:', e)
    finally:
        try:
            if 'original' in locals() and original is not None:
                restore_printer_pdevmode(printer_name, original)
                print('Oude printerinstellingen hersteld.')
        except Exception as e:
            print('Kon printerinstellingen niet herstellen:', e)


if __name__ == '__main__':
    print('Printer lade-vinder - interactief')
    interactive()