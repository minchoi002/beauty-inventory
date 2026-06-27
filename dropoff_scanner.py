"""
ACI Express - Drop-Off Scanner
Zebra 4x6 label printer (ZPL via win32print)
USB barcode scanner (keyboard input mode)

Modes:
  * Carrier Drop-Off (default) - auto-detects UPS / USPS / FedEx and prints
    a drop-off confirmation slip.
  * Amazon Return (toggle with F2) - records a label-free Amazon return,
    prints a return receipt with a scannable QR code of the return code.
"""

import tkinter as tk
from tkinter import messagebox, ttk
import re
import csv
import os
import sys
from datetime import datetime

try:
    import win32print
    WIN32 = True
except ImportError:
    WIN32 = False

# ── Settings ──────────────────────────────────────────────
COMPANY_NAME  = "EASYSHIP EXPRESS"
COMPANY_PHONE = "(470) 524-0882"
COMPANY_ADDR  = "525 PEACHTREE INDUSTRIAL BLVD STE G SUWANEE GA 30024"
PRINTER_NAME  = None               # None = use Windows default printer
CSV_PATH      = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dropoff_log.csv")
# ──────────────────────────────────────────────────────────


def save_to_csv(tracking: str, carrier: str) -> None:
    now = datetime.now()
    file_exists = os.path.isfile(CSV_PATH)
    with open(CSV_PATH, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["Date", "Time", "Carrier", "Tracking Number"])
        writer.writerow([
            now.strftime("%m/%d/%Y"),
            now.strftime("%I:%M %p"),
            carrier,
            tracking,
        ])


def detect_carrier(raw: str) -> str:
    t = raw.strip().replace(" ", "").upper()
    if t.startswith("1Z"):
        return "UPS"
    if re.match(r"^(9[40][01]\d{18}|420\d{27}|9[2345]\d{20})", t):
        return "USPS"
    if re.match(r"^(96|74|75|88|62)\d+", t) or re.match(r"^\d{12}$", t):
        return "FedEx"
    return "FedEx / UPS"


def is_amazon_return(raw: str) -> bool:
    """Auto-detect Amazon label-free / box-free return QR codes.

    Amazon "No Box No Label" returns (dropped at The UPS Store, Whole Foods,
    Kohl's, etc.) encode a pipe-delimited string that begins with the marker
    ``AFNLFBF``. We also catch the generic ``AMAZON RETURN`` / ``AMZN`` text.
    Anything else is left to manual Amazon Return mode (F2).
    """
    u = raw.strip().upper()
    if u.startswith("AFNLFBF"):
        return True
    if "AMAZON RETURN" in u or "AMZN" in u:
        return True
    return False


def parse_amazon_return(raw: str):
    """Pull a clean (return_id, item_description) out of an Amazon QR string.

    The QR payload looks like::

        AFNLFBF|<return_id>|01|<return center>|...|<qty>|<item name>|

    For anything that isn't the recognized pipe format we fall back to using
    the raw scan as the return id with no item description.
    """
    parts = [p.strip() for p in raw.split("|")]
    if len(parts) >= 2 and parts[0].upper() == "AFNLFBF":
        return_id = parts[1]
        nonempty  = [p for p in parts if p]
        item      = nonempty[-1] if nonempty else ""
        if item == return_id:           # no separate item field present
            item = ""
        return return_id, item
    return raw, ""


# ── Korean IME recovery ───────────────────────────────────
# A USB scanner types its data like a keyboard. If the Windows Korean IME is
# active, the Latin keystrokes get composed into Hangul (e.g. "AFNLFBF" -> a
# string of jamo). disable_ime() stops that at the source; hangul_to_qwerty()
# is a fallback that maps any Hangul that still slips through back to the keys
# the scanner pressed, using the standard 2-set (두벌식) layout.

_CHO  = list("ㄱㄲㄴㄷㄸㄹㅁㅂㅃㅅㅆㅇㅈㅉㅊㅋㅌㅍㅎ")
_JUNG = list("ㅏㅐㅑㅒㅓㅔㅕㅖㅗㅘㅙㅚㅛㅜㅝㅞㅟㅠㅡㅢㅣ")
_JONG = ["", "ㄱ", "ㄲ", "ㄳ", "ㄴ", "ㄵ", "ㄶ", "ㄷ", "ㄹ", "ㄺ", "ㄻ", "ㄼ",
         "ㄽ", "ㄾ", "ㄿ", "ㅀ", "ㅁ", "ㅂ", "ㅄ", "ㅅ", "ㅆ", "ㅇ", "ㅈ", "ㅊ",
         "ㅋ", "ㅌ", "ㅍ", "ㅎ"]

# jamo -> QWERTY key(s). Shifted doubles preserve upper-case (ㄲ->R, ㅃ->Q ...);
# all other letters share a key with no case info, so they default to lower.
_JAMO_TO_KEY = {
    "ㄱ": "r", "ㄲ": "R", "ㄴ": "s", "ㄷ": "e", "ㄸ": "E", "ㄹ": "f", "ㅁ": "a",
    "ㅂ": "q", "ㅃ": "Q", "ㅅ": "t", "ㅆ": "T", "ㅇ": "d", "ㅈ": "w", "ㅉ": "W",
    "ㅊ": "c", "ㅋ": "z", "ㅌ": "x", "ㅍ": "v", "ㅎ": "g",
    "ㅏ": "k", "ㅐ": "o", "ㅑ": "i", "ㅒ": "O", "ㅓ": "j", "ㅔ": "p", "ㅕ": "u",
    "ㅖ": "P", "ㅗ": "h", "ㅛ": "y", "ㅜ": "n", "ㅠ": "b", "ㅡ": "m", "ㅣ": "l",
    # compound vowels / finals decompose into two keys
    "ㅘ": "hk", "ㅙ": "ho", "ㅚ": "hl", "ㅝ": "nj", "ㅞ": "np", "ㅟ": "nl",
    "ㅢ": "ml", "ㄳ": "rt", "ㄵ": "sw", "ㄶ": "sg", "ㄺ": "fr", "ㄻ": "fa",
    "ㄼ": "fq", "ㄽ": "ft", "ㄾ": "fx", "ㄿ": "fv", "ㅀ": "fg", "ㅄ": "qt",
}


def contains_hangul(text: str) -> bool:
    return any(0xAC00 <= ord(c) <= 0xD7A3 or 0x3131 <= ord(c) <= 0x3163
               for c in text)


def hangul_to_qwerty(text: str) -> str:
    """Best-effort: undo a Korean-IME mangle of scanner input back to Latin."""
    out = []
    for ch in text:
        code = ord(ch)
        if 0xAC00 <= code <= 0xD7A3:               # composed syllable
            idx  = code - 0xAC00
            out.append(_JAMO_TO_KEY.get(_CHO[idx // 588], ""))
            out.append(_JAMO_TO_KEY.get(_JUNG[(idx % 588) // 28], ""))
            jong = idx % 28
            if jong:
                out.append(_JAMO_TO_KEY.get(_JONG[jong], ""))
        elif 0x3131 <= code <= 0x3163:             # standalone jamo
            out.append(_JAMO_TO_KEY.get(ch, ch))
        else:
            out.append(ch)
    return "".join(out)


def disable_ime(*widgets) -> None:
    """Detach the Windows IME from our windows so scans arrive as raw ASCII.

    No-op on non-Windows or when the IME APIs aren't available.
    """
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        imm32 = ctypes.windll.imm32
        for w in widgets:
            try:
                imm32.ImmAssociateContext(w.winfo_id(), 0)
            except Exception:
                pass
    except Exception:
        pass


def build_zpl(tracking: str, carrier: str) -> str:
    now      = datetime.now()
    date_str = now.strftime("%m/%d/%Y")
    time_str = now.strftime("%I:%M %p")

    # 4 × 6 inch @ 203 DPI  →  812 × 1218 dots
    # Fonts: A0 = scalable. ^BCN = Code128 barcode (auto-detects subset)
    return f"""^XA
^PW812
^LL1218
^CI28

^FO40,40^A0N,55,55^FD{COMPANY_NAME}^FS
^FO40,105^A0N,32,32^FDDROP-OFF CONFIRMATION^FS

^FO40,155^GB732,3,3^FS

^FO40,175^A0N,28,28^FDDate :^FS^FO220,175^A0N,28,28^FD{date_str}^FS
^FO40,215^A0N,28,28^FDTime :^FS^FO220,215^A0N,28,28^FD{time_str}^FS
^FO40,255^A0N,28,28^FDCarrier :^FS^FO220,255^A0N,28,28^FD{carrier}^FS

^FO40,305^A0N,28,28^FDTracking # :^FS
^FO40,340^A0N,34,34^FD{tracking}^FS

^FO40,395^BCN,80,Y,N,N
^FD{tracking}^FS

^FO40,495^GB732,3,3^FS

^FO40,515^A0N,26,26^FDYour package has been received.^FS
^FO40,550^A0N,26,26^FDPlease keep this slip for your records.^FS

^FO40,605^GB732,3,3^FS

^FO40,625^A0N,24,24^FDQuestions? {COMPANY_PHONE}^FS
^FO40,660^A0N,24,24^FD{COMPANY_ADDR}^FS

^XZ"""


def build_amazon_zpl(return_id: str, item: str = "", qr_payload: str = "") -> str:
    """Amazon return drop-off slip.

    The printed QR encodes the FULL original Amazon payload (``qr_payload``)
    so it stays identical to the customer's QR and can be scanned at the UPS
    counter to accept the return. The return id / item are printed as plain
    text for the store's own reference. Falls back to the return id if no raw
    payload was supplied.
    """
    now      = datetime.now()
    date_str = now.strftime("%m/%d/%Y")
    time_str = now.strftime("%I:%M %p")

    qr_data = qr_payload or return_id

    # Optional item line (wraps up to 2 lines via ^FB field block).
    item_block = (
        f"^FO40,395^A0N,26,26^FDItem :^FS\n"
        f"^FO40,430^A0N,24,24^FB732,2,0,L^FD{item}^FS\n"
        if item else ""
    )

    # 4 × 6 inch @ 203 DPI  →  812 × 1218 dots
    # ^BQ = QR code (model 2), magnification 5. ^FDMA, = error-correction M,
    # automatic input. The whole Amazon payload is re-encoded here.
    return f"""^XA
^PW812
^LL1218
^CI28

^FO40,40^A0N,55,55^FD{COMPANY_NAME}^FS

^FO40,110^GB732,52,52^FS
^FO40,118^A0N,38,38^FR^FDAMAZON RETURN^FS

^FO40,185^GB732,3,3^FS

^FO40,205^A0N,28,28^FDDate :^FS^FO220,205^A0N,28,28^FD{date_str}^FS
^FO40,245^A0N,28,28^FDTime :^FS^FO220,245^A0N,28,28^FD{time_str}^FS

^FO40,305^A0N,28,28^FDReturn ID :^FS
^FO40,340^A0N,32,32^FD{return_id}^FS
{item_block}
^FO40,510^A0N,26,26^FDScan this QR at UPS to drop off:^FS
^FO180,555^BQN,2,5^FDMA,{qr_data}^FS

^FO40,1010^GB732,3,3^FS

^FO40,1030^A0N,24,24^FDQuestions? {COMPANY_PHONE}^FS
^FO40,1065^A0N,24,24^FD{COMPANY_ADDR}^FS

^XZ"""


def send_to_printer(zpl: str) -> None:
    if not WIN32:
        raise RuntimeError("pywin32 not installed. Run: pip install pywin32")

    printer = PRINTER_NAME or win32print.GetDefaultPrinter()
    h = win32print.OpenPrinter(printer)
    try:
        job = win32print.StartDocPrinter(h, 1, ("Drop-Off Label", None, "RAW"))
        try:
            win32print.StartPagePrinter(h)
            win32print.WritePrinter(h, zpl.encode("utf-8"))
            win32print.EndPagePrinter(h)
        finally:
            win32print.EndDocPrinter(h)
    finally:
        win32print.ClosePrinter(h)


# ── GUI ───────────────────────────────────────────────────

class App:
    BG      = "#0f1117"
    CARD    = "#1a1d27"
    GREEN   = "#00d4aa"
    ORANGE  = "#ff9900"          # Amazon accent
    RED     = "#ff4757"
    MUTED   = "#5c6278"
    WHITE   = "#f0f0f0"
    FONT    = "Segoe UI"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.amazon_mode = False
        root.title("ACI Express – Drop-Off Scanner")
        root.geometry("640x540")
        root.configure(bg=self.BG)
        root.resizable(False, False)
        root.attributes("-topmost", True)
        self._build()
        # Keep the Korean IME from rewriting scanner keystrokes into Hangul.
        root.update_idletasks()
        disable_ime(root, self.entry)
        root.bind("<F2>", lambda _: self._toggle_amazon())

    def _build(self):
        # ── header ──
        hdr = tk.Frame(self.root, bg=self.CARD, pady=18)
        hdr.pack(fill="x")
        tk.Label(hdr, text=COMPANY_NAME, font=(self.FONT, 26, "bold"),
                 bg=self.CARD, fg=self.GREEN).pack()
        tk.Label(hdr, text="Drop-Off Scanner", font=(self.FONT, 12),
                 bg=self.CARD, fg=self.MUTED).pack()

        # ── mode toggle ──
        mode_frame = tk.Frame(self.root, bg=self.BG)
        mode_frame.pack(pady=(16, 0))
        self.mode_btn = tk.Button(
            mode_frame, text="↩  Amazon Return Mode  (F2)",
            font=(self.FONT, 11, "bold"), relief="flat",
            bg=self.CARD, fg=self.MUTED, activebackground=self.ORANGE,
            activeforeground=self.BG, cursor="hand2", bd=0, padx=16, pady=8,
            command=self._toggle_amazon)
        self.mode_btn.pack()

        # ── status ──
        self.status_var = tk.StringVar(value="Ready to Scan")
        self.status_lbl = tk.Label(self.root, textvariable=self.status_var,
                                    font=(self.FONT, 18, "bold"),
                                    bg=self.BG, fg=self.GREEN)
        self.status_lbl.pack(pady=(22, 8))

        # ── scan entry ──
        entry_frame = tk.Frame(self.root, bg=self.BG)
        entry_frame.pack()
        tk.Label(entry_frame, text="Scan barcode:", font=(self.FONT, 11),
                 bg=self.BG, fg=self.MUTED).pack(anchor="w")
        self.entry_var = tk.StringVar()
        self.entry = tk.Entry(entry_frame, textvariable=self.entry_var,
                               font=(self.FONT, 16), bg=self.CARD, fg=self.WHITE,
                               insertbackground=self.WHITE, relief="flat",
                               bd=8, width=30, justify="center")
        self.entry.pack(ipady=8)
        self.entry.bind("<Return>", self._on_scan)
        self.entry.focus_set()

        # ── last scan log ──
        self.log_var = tk.StringVar()
        tk.Label(self.root, textvariable=self.log_var,
                 font=(self.FONT, 11), bg=self.BG, fg=self.MUTED).pack(pady=14)

        # ── history listbox ──
        hist_frame = tk.Frame(self.root, bg=self.BG)
        hist_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        tk.Label(hist_frame, text="Today's scans", font=(self.FONT, 9),
                 bg=self.BG, fg=self.MUTED).pack(anchor="w")
        self.listbox = tk.Listbox(hist_frame, bg=self.CARD, fg=self.WHITE,
                                   font=(self.FONT, 10), relief="flat",
                                   selectbackground=self.GREEN,
                                   selectforeground=self.BG, height=6)
        self.listbox.pack(fill="both", expand=True)

        # keep focus
        self.root.bind("<FocusIn>", lambda _: self.entry.focus_set())

    # ── mode toggle ────────────────────────────────────────
    def _toggle_amazon(self):
        self.amazon_mode = not self.amazon_mode
        if self.amazon_mode:
            self.mode_btn.config(text="↩  Amazon Return Mode: ON  (F2)",
                                 bg=self.ORANGE, fg=self.BG)
            self._set_status("Amazon Return — Scan QR", self.ORANGE)
        else:
            self.mode_btn.config(text="↩  Amazon Return Mode  (F2)",
                                 bg=self.CARD, fg=self.MUTED)
            self._set_status("Ready to Scan", self.GREEN)
        self.entry.focus_set()

    # ── scan handler ──────────────────────────────────────
    def _on_scan(self, _=None):
        raw = self.entry_var.get().strip()
        self.entry_var.set("")
        self.entry.focus_set()
        if not raw:
            return

        # Safety net: if the Korean IME slipped through and turned the scan
        # into Hangul, convert it back to the Latin keys the scanner sent.
        ime_fixed = False
        if contains_hangul(raw):
            raw = hangul_to_qwerty(raw)
            ime_fixed = True

        amazon = self.amazon_mode or is_amazon_return(raw)

        if amazon:
            return_id, item = parse_amazon_return(raw)
            label    = "Amazon Return"
            logged   = return_id
            zpl      = build_amazon_zpl(return_id, item, raw)
            ok_text  = "✓  Amazon Return Logged!"
            ok_color = self.ORANGE
        else:
            carrier = detect_carrier(raw)
            label   = carrier
            logged  = raw
            zpl     = build_zpl(raw, carrier)
            ok_text = "✓  Label Printed!"
            ok_color = self.GREEN

        try:
            send_to_printer(zpl)
            save_to_csv(logged, label)
            self._set_status(ok_text, ok_color)
            ts   = datetime.now().strftime("%I:%M %p")
            note = "  ⌨ 한글→영문 자동변환" if ime_fixed else ""
            self.log_var.set(f"[{label}]  {logged}   {ts}{note}")
            self.listbox.insert(0, f"{ts}  {label:18}  {logged}")
        except Exception as exc:
            self._set_status("✗  Print Error", self.RED)
            messagebox.showerror("Print Error", str(exc))

        # return to the resting state for the current mode
        rest_text  = "Amazon Return — Scan QR" if self.amazon_mode else "Ready to Scan"
        rest_color = self.ORANGE if self.amazon_mode else self.GREEN
        self.root.after(2500, lambda: self._set_status(rest_text, rest_color))

    def _set_status(self, text: str, color: str):
        self.status_var.set(text)
        self.status_lbl.config(fg=color)


# ── entry point ───────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
