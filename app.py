import csv
import ctypes
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from statistics import median
from tkinter import filedialog, messagebox, ttk

from openpyxl import Workbook
from PIL import Image, ImageOps

INVOICE_HEADER = ["#", "Malın adı", "Sayı", "Qiymət", "Məbləğ"]


def rows_from_tsv(tsv):
    if not tsv.lstrip().startswith("level\t"):
        first = tsv.split("\n", 1)[0]
        if first.startswith("1\t"):
            tsv = "level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\twidth\theight\tconf\ttext\n" + tsv
    words = []
    for item in csv.DictReader(io.StringIO(tsv), delimiter="\t", quoting=csv.QUOTE_NONE):
        value = (item.get("text") or "").strip()
        if value:
            x, y, width, height = (int(item[k]) for k in ("left", "top", "width", "height"))
            words.append((y + height / 2, x, width, height, value))
    if not words:
        return []

    height = median(word[3] for word in words)
    lines = []
    for word in sorted(words):
        if not lines or word[0] - median(w[0] for w in lines[-1]) > height * 0.55:
            lines.append([])
        lines[-1].append(word)

    rows = []
    invoice_columns = None
    for line in lines:
        cells = []
        starts = []
        ends = []
        right = None
        for _, x, width, _, value in sorted(line, key=lambda word: word[1]):
            if right is None or x - right > height * 0.85:
                cells.append(value)
                starts.append(x)
                ends.append(x + width)
            else:
                cells[-1] += " " + value
                ends[-1] = x + width
            right = x + width
        if [cell.lower() for cell in cells] == ["#", "malın adı", "sayı", "qiymət", "məbləğ"]:
            invoice_columns = [
                (ends[index] + starts[index + 1]) / 2 for index in range(4)
            ]
        elif invoice_columns and cells[0].isdigit():
            aligned = [""] * 5
            for _, x, width, _, value in sorted(line, key=lambda word: word[1]):
                column = sum(x + width / 2 > boundary for boundary in invoice_columns)
                aligned[column] = " ".join(filter(None, (aligned[column], value)))
            cells = aligned
        rows.append(cells)
    invoice_rows = []
    product_rows = []
    for index, line in enumerate(lines):
        value = " ".join(word[4] for word in sorted(line, key=lambda word: word[1]))
        total = re.search(r"blok\s+say[ıiI1l]\s*[:|]?\s*(\d+)", value, re.IGNORECASE)
        if total and product_rows:
            if len(invoice_rows) == len(product_rows) and sum(int(row[2]) for row in invoice_rows) == int(total[1]):
                return [INVOICE_HEADER, *invoice_rows]
            return [INVOICE_HEADER, *product_rows]
        if re.search(r"\bbank\s+rekvizit|\bbank[ıi]n\s+v[öo]en|\bswift\s*:", value, re.IGNORECASE):
            return rows[:index]
        cells = rows[index]
        if len(cells) >= 5 and cells[0].isascii() and cells[0].isdigit() and (
            int(cells[0]) > int(product_rows[-1][0]) if product_rows else cells[0] == "1"
        ):
            serial, *name, quantity, price, amount = cells
            product_rows.append([serial, " ".join(name), quantity, price, amount])
        match = re.fullmatch(
            r"(\d+)\s+(.+?)\s+(\d+)\s+(\d+[.,]\s*\d{2})\s+(\d+(?:[.]\d{3})*[.,]\s*\d{2})",
            value,
        )
        if match and int(match[1]) == len(invoice_rows) + 1:
            serial, name, quantity, price, amount = match.groups()
            invoice_rows.append([serial, name, quantity, price.replace(" ", ""), amount.replace(" ", "")])
    return rows


def recognize(path, language):
    if language == "auto":
        rows = recognize(path, "eng")
        if rows and rows[0] == INVOICE_HEADER:
            return rows
        return recognize(path, "aze")

    project = Path(__file__).resolve().parent
    runtime = project / ".ocr-runtime"
    data = project / "tessdata"
    if runtime.is_dir() and (data / f"{language}.traineddata").is_file():
        if str(runtime) not in sys.path:
            sys.path.insert(0, str(runtime))
        from tesserocr import PSM, PyTessBaseAPI

        with Image.open(path) as source:
            image = ImageOps.exif_transpose(source).convert("L")
        data_path = str(data)
        if os.name == "nt":
            short_path = ctypes.create_unicode_buffer(32768)
            if ctypes.windll.kernel32.GetShortPathNameW(data_path, short_path, len(short_path)):
                data_path = short_path.value
        with PyTessBaseAPI(path=data_path, lang=language, psm=PSM.SINGLE_BLOCK) as api:
            api.SetImage(image)
            api.Recognize()
            return rows_from_tsv(api.GetTSVText(0))

    default = Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Tesseract-OCR" / "tesseract.exe"
    local = Path(__file__).parent / "Tesseract-OCR" / "tesseract.exe"
    executable = next((str(p) for p in (local, default) if p.is_file()), None)
    executable = executable or shutil.which("tesseract")
    if not executable:
        raise RuntimeError("Tesseract tapılmadı. Quraşdırma addımlarına README.md faylında baxın.")
    with Image.open(path) as source:
        image = ImageOps.exif_transpose(source).convert("L")
        png = io.BytesIO()
        image.save(png, format="PNG")
    result = subprocess.run(
        [executable, "stdin", "stdout", "-l", language, "--psm", "6", "tsv"],
        input=png.getvalue(),
        capture_output=True,
        timeout=90,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if result.returncode:
        raise RuntimeError(result.stderr.decode("utf-8", errors="replace").strip())
    return rows_from_tsv(result.stdout.decode("utf-8"))


def save_rows(path, rows):
    extension = Path(path).suffix.lower()
    if extension == ".xlsx":
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Qaimə"
        for row_index, row in enumerate(rows, 1):
            for column_index, value in enumerate(row, 1):
                cell = sheet.cell(row_index, column_index, value)
                cell.data_type = "s"
        workbook.save(path)
    elif extension == ".json":
        Path(path).write_text(json.dumps({"rows": rows}, ensure_ascii=False, indent=2), encoding="utf-8")
    else:
        Path(path).write_text("\n".join("\t".join(row) for row in rows), encoding="utf-8")


class App:
    def __init__(self, root):
        self.root = root
        self.pool = ThreadPoolExecutor(max_workers=1)
        root.title("Qaiməni cədvələ çevir")
        root.geometry("850x500")
        bar = ttk.Frame(root, padding=10)
        bar.pack(fill="x")
        self.pick = ttk.Button(bar, text="Şəkil seç", command=self.choose)
        self.pick.pack(side="left")
        self.languages = {"Avtomatik": "auto", "Azərbaycan dili": "aze", "İngilis dili": "eng"}
        self.language = ttk.Combobox(bar, values=list(self.languages), state="readonly")
        self.language.current(0)
        self.language.pack(side="left", padx=8)
        self.exports = []
        for label, extension in (("Excel (.xlsx)", ".xlsx"), ("JSON", ".json"), ("Mətn (.txt)", ".txt")):
            button = ttk.Button(bar, text=label, command=lambda ext=extension: self.save(ext), state="disabled")
            button.pack(side="left", padx=3)
            self.exports.append(button)
        frame = ttk.Frame(root, padding=(10, 0))
        frame.pack(fill="both", expand=True)
        self.preview = tk.Text(frame, wrap="none", undo=True)
        scroll = ttk.Scrollbar(frame, orient="vertical", command=self.preview.yview)
        horizontal = ttk.Scrollbar(frame, orient="horizontal", command=self.preview.xview)
        self.preview.configure(yscrollcommand=scroll.set, xscrollcommand=horizontal.set)
        scroll.pack(side="right", fill="y")
        horizontal.pack(side="bottom", fill="x")
        self.preview.pack(fill="both", expand=True)
        self.status = tk.StringVar(value="Qaimənin cədvəl hissəsini göstərən aydın şəkil seçin.")
        ttk.Label(root, textvariable=self.status, padding=10).pack(anchor="w")
        root.protocol("WM_DELETE_WINDOW", self.close)

    def choose(self):
        path = filedialog.askopenfilename(filetypes=[("Şəkillər", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp")])
        if not path:
            return
        self.pick.configure(state="disabled")
        for button in self.exports:
            button.configure(state="disabled")
        self.status.set("Şəkil oxunur…")
        future = self.pool.submit(recognize, path, self.languages[self.language.get()])
        self.root.after(150, self.finish, future)

    def finish(self, future):
        if not future.done():
            self.root.after(150, self.finish, future)
            return
        self.pick.configure(state="normal")
        try:
            rows = future.result()
        except Exception as error:
            self.status.set("Şəkli oxumaq mümkün olmadı.")
            messagebox.showerror("Oxuma xətası", str(error))
            rows = []
        else:
            self.status.set(
                f"{len(rows)} sətir. Saxlamazdan əvvəl mətni və sütunları yoxlaya bilərsiniz."
                if rows else "Şəkildə oxuna bilən mətn tapılmadı."
            )
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", "\n".join("\t".join(row) for row in rows))
        for button in self.exports:
            button.configure(state="normal" if rows else "disabled")

    def save(self, extension):
        rows = [line.split("\t") for line in self.preview.get("1.0", "end-1c").splitlines()]
        if not rows:
            return
        path = filedialog.asksaveasfilename(defaultextension=extension, filetypes=[(extension, "*" + extension)])
        if path:
            try:
                save_rows(path, rows)
            except (OSError, ValueError) as error:
                messagebox.showerror("Saxlama xətası", str(error))
            else:
                self.status.set(f"Saxlandı: {Path(path).name}")

    def close(self):
        self.pool.shutdown(wait=False, cancel_futures=True)
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
