"""
DataMatrix / QR Code Extractor для NiceLabel
=============================================
Поддерживает: DataMatrix (Честный Знак), QR Code
При первом запуске автоматически установит все необходимые библиотеки.
Требования: Python 3.8+ (https://python.org)
"""

import sys
import subprocess
import importlib

# ── Автоустановка зависимостей ──────────────────────────────────────────────
REQUIRED = {
    "fitz":       "pymupdf",
    "cv2":        "opencv-python",
    "pylibdmtx":  "pylibdmtx",
    "openpyxl":   "openpyxl",
    "PIL":        "Pillow",
    "numpy":      "numpy",
}

def auto_install():
    missing = []
    for module, package in REQUIRED.items():
        try:
            importlib.import_module(module)
        except ImportError:
            missing.append(package)
    if missing:
        print(f"Устанавливаю библиотеки: {', '.join(missing)} ...")
        for pkg in missing:
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "--quiet"])
        print("Установка завершена! Перезапустите программу.")
        input("Нажмите Enter для выхода...")
        sys.exit(0)

auto_install()

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import csv

import fitz
import cv2
import numpy as np
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── DataMatrix декодер (pylibdmtx) ──────────────────────────────────────────
try:
    from pylibdmtx.pylibdmtx import decode as dmtx_decode
    DMTX_OK = True
except Exception:
    DMTX_OK = False
    dmtx_decode = None

# ── QR Code декодер (opencv встроенный) ─────────────────────────────────────
def try_qr_opencv(img_cv):
    det = cv2.QRCodeDetector()
    gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
    for variant in [gray,
                    cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1],
                    cv2.bitwise_not(cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1])]:
        data, _, _ = det.detectAndDecode(variant)
        if data:
            return data
    return None


class QRExtractorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("DataMatrix / QR Extractor → NiceLabel")
        self.root.geometry("640x560")
        self.root.resizable(False, False)
        self.root.configure(bg="#f1f5f9")
        self.pdf_path   = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.results    = []
        self._build_ui()

    def _build_ui(self):
        hdr = tk.Frame(self.root, bg="#1d4ed8", height=64)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="🏷  DataMatrix / QR  →  NiceLabel",
                 font=("Segoe UI", 13, "bold"), fg="white", bg="#1d4ed8").pack(expand=True)

        body = tk.Frame(self.root, bg="#f1f5f9", padx=24, pady=18)
        body.pack(fill="both", expand=True)

        self._section(body, "1.  Выберите PDF файл")
        row1 = tk.Frame(body, bg="#f1f5f9"); row1.pack(fill="x", pady=(4, 12))
        tk.Entry(row1, textvariable=self.pdf_path, state="readonly",
                 font=("Segoe UI", 9), bg="white", relief="solid", bd=1
                 ).pack(side="left", fill="x", expand=True, ipady=5)
        self._btn(row1, "Обзор...", self.choose_pdf, "#1d4ed8").pack(side="left", padx=(6, 0))

        self._section(body, "2.  Папка для сохранения результатов")
        row2 = tk.Frame(body, bg="#f1f5f9"); row2.pack(fill="x", pady=(4, 16))
        tk.Entry(row2, textvariable=self.output_dir, state="readonly",
                 font=("Segoe UI", 9), bg="white", relief="solid", bd=1
                 ).pack(side="left", fill="x", expand=True, ipady=5)
        self._btn(row2, "Обзор...", self.choose_output, "#1d4ed8").pack(side="left", padx=(6, 0))

        self.btn_start = self._btn(body, "Начать извлечение кодов",
                                   self.start_extraction, "#15803d",
                                   font=("Segoe UI", 11, "bold"), pady=10)
        self.btn_start.pack(fill="x", pady=(0, 14))

        tk.Label(body, text="Прогресс:", font=("Segoe UI", 9, "bold"),
                 bg="#f1f5f9", fg="#1e293b").pack(anchor="w")
        self.progress = ttk.Progressbar(body, mode="determinate")
        self.progress.pack(fill="x", pady=(4, 10))

        log_wrap = tk.Frame(body, bg="#0f172a")
        log_wrap.pack(fill="both", expand=True)
        self.log = tk.Text(log_wrap, font=("Consolas", 9), bg="#0f172a",
                           fg="#94a3b8", relief="flat", state="disabled", wrap="word", height=10)
        sb = tk.Scrollbar(log_wrap, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True, padx=8, pady=6)
        sb.pack(side="right", fill="y")
        self.log.tag_config("green",  foreground="#4ade80")
        self.log.tag_config("yellow", foreground="#facc15")
        self.log.tag_config("red",    foreground="#f87171")
        self.log.tag_config("blue",   foreground="#60a5fa")

    @staticmethod
    def _section(parent, text):
        tk.Label(parent, text=text, font=("Segoe UI", 10, "bold"),
                 bg="#f1f5f9", fg="#1e293b").pack(anchor="w")

    @staticmethod
    def _btn(parent, text, cmd, color, **kw):
        return tk.Button(parent, text=text, command=cmd, bg=color, fg="white",
                         relief="flat", font=kw.pop("font", ("Segoe UI", 9, "bold")),
                         padx=12, cursor="hand2", **kw)

    def log_msg(self, msg, tag=None):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag or "")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.root.update_idletasks()

    def choose_pdf(self):
        p = filedialog.askopenfilename(filetypes=[("PDF файлы", "*.pdf")])
        if p:
            self.pdf_path.set(p)
            if not self.output_dir.get():
                self.output_dir.set(os.path.dirname(p))

    def choose_output(self):
        p = filedialog.askdirectory()
        if p:
            self.output_dir.set(p)

    def start_extraction(self):
        if not self.pdf_path.get():
            messagebox.showwarning("Внимание", "Выберите PDF файл!"); return
        if not self.output_dir.get():
            messagebox.showwarning("Внимание", "Выберите папку для сохранения!"); return
        self.btn_start.configure(state="disabled", bg="#94a3b8")
        self.results = []
        self.progress["value"] = 0
        threading.Thread(target=self._run_extraction, daemon=True).start()

    def _decode_page(self, pix, pn):
        """Пробует распознать DataMatrix и QR на одном pixmap."""

        # ── 1. DataMatrix через pylibdmtx ──────────────────────────────────
        if DMTX_OK:
            pil_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            results = dmtx_decode(pil_img, timeout=3000)
            if results:
                for r in results:
                    data = r.data.decode("utf-8", errors="replace")
                    self.log_msg(f"  Стр.{pn} [DataMatrix]: {data[:70]}{'...' if len(data)>70 else ''}", "green")
                    self.results.append({"Страница": pn, "Тип": "DataMatrix", "Данные": data})
                return True

            # Попробуем grayscale
            gray_pil = pil_img.convert("L")
            results = dmtx_decode(gray_pil, timeout=3000)
            if results:
                for r in results:
                    data = r.data.decode("utf-8", errors="replace")
                    self.log_msg(f"  Стр.{pn} [DataMatrix]: {data[:70]}{'...' if len(data)>70 else ''}", "green")
                    self.results.append({"Страница": pn, "Тип": "DataMatrix", "Данные": data})
                return True

        # ── 2. QR Code через OpenCV ─────────────────────────────────────────
        img_cv = cv2.imdecode(
            np.frombuffer(pix.tobytes("png"), np.uint8),
            cv2.IMREAD_COLOR
        )
        data = try_qr_opencv(img_cv)
        if data:
            self.log_msg(f"  Стр.{pn} [QR Code]: {data[:70]}{'...' if len(data)>70 else ''}", "green")
            self.results.append({"Страница": pn, "Тип": "QR Code", "Данные": data})
            return True

        return False

    def _run_extraction(self):
        try:
            pdf   = self.pdf_path.get()
            self.log_msg(f"Файл: {os.path.basename(pdf)}")

            if not DMTX_OK:
                self.log_msg("ВНИМАНИЕ: pylibdmtx не загружен - DataMatrix не будет работать!", "red")
            else:
                self.log_msg("pylibdmtx OK - DataMatrix (Честный Знак) поддерживается", "green")

            doc   = fitz.open(pdf)
            total = len(doc)
            self.log_msg(f"Страниц: {total}")
            found = 0

            for i, page in enumerate(doc):
                pn = i + 1

                # Сначала zoom=4 (~288 dpi)
                pix = page.get_pixmap(matrix=fitz.Matrix(4, 4), colorspace=fitz.csRGB)
                ok = self._decode_page(pix, pn)

                # Если не нашли - пробуем zoom=6 (~432 dpi)
                if not ok:
                    pix2 = page.get_pixmap(matrix=fitz.Matrix(6, 6), colorspace=fitz.csRGB)
                    ok = self._decode_page(pix2, pn)

                if ok:
                    found += 1
                else:
                    self.log_msg(f"  Стр.{pn}: не найден", "yellow")

                self.progress["value"] = pn / total * 100
                self.root.update_idletasks()

            doc.close()
            self.log_msg(f"\nНайдено: {found} кодов из {total} страниц")

            if self.results:
                self._save_results()
            else:
                self.log_msg("Коды не обнаружены!", "red")
                messagebox.showwarning("Результат", "Коды не найдены в документе.")

        except Exception as e:
            self.log_msg(f"Ошибка: {e}", "red")
            messagebox.showerror("Ошибка", str(e))
        finally:
            self.btn_start.configure(state="normal", bg="#15803d")

    def _save_results(self):
        out  = self.output_dir.get()
        base = os.path.splitext(os.path.basename(self.pdf_path.get()))[0]

        csv_path = os.path.join(out, f"{base}_codes.csv")
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=["Страница", "Тип", "Данные"])
            w.writeheader(); w.writerows(self.results)
        self.log_msg(f"CSV: {csv_path}", "blue")

        xlsx_path = os.path.join(out, f"{base}_codes.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Коды"

        thin   = Side(style="thin", color="CBD5E1")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        h_font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
        h_fill = PatternFill("solid", fgColor="1D4ED8")
        h_aln  = Alignment(horizontal="center", vertical="center")

        for col, h in enumerate(["Страница", "Тип кода", "Данные  (для NiceLabel)"], 1):
            c = ws.cell(row=1, column=col, value=h)
            c.font = h_font; c.fill = h_fill; c.alignment = h_aln; c.border = border

        even_fill = PatternFill("solid", fgColor="EFF6FF")
        for ri, item in enumerate(self.results, 2):
            fill = even_fill if ri % 2 == 0 else PatternFill()
            for ci, key in enumerate(["Страница", "Тип", "Данные"], 1):
                c = ws.cell(row=ri, column=ci, value=item[key])
                c.border = border; c.fill = fill
                c.alignment = Alignment(vertical="center", wrap_text=(ci == 3))

        ws.column_dimensions["A"].width = 11
        ws.column_dimensions["B"].width = 16
        ws.column_dimensions["C"].width = 65
        ws.row_dimensions[1].height = 22
        wb.save(xlsx_path)
        self.log_msg(f"Excel: {xlsx_path}", "blue")
        self.log_msg(f"\nГотово! Файлы сохранены в:\n   {out}", "green")

        messagebox.showinfo("Готово!",
            f"Извлечено кодов: {len(self.results)}\n\n"
            f"Сохранены файлы:\n"
            f"  {os.path.basename(csv_path)}\n"
            f"  {os.path.basename(xlsx_path)}\n\n"
            f"Папка: {out}")


if __name__ == "__main__":
    root = tk.Tk()
    QRExtractorApp(root)
    root.mainloop()
