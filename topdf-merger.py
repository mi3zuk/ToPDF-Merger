#!/usr/bin/env python3
#encoding: utf-8

import os
import re
import io
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
from PyPDF2 import PdfReader, PdfWriter

def natural_key(s: str):
    parts = re.split(r'(\d+)', s)
    key = []
    for p in parts:
        key.append(int(p) if p.isdigit() else p.lower())
    return key

class ToPDFMerger(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("To PDF Merger")
        self.geometry("900x600")
        self.files = []
        self._drag_start = None
        self._current_img = None
        self._build_ui()

    def _build_ui(self):
        paned = tk.PanedWindow(self, orient=tk.HORIZONTAL,
                               sashwidth=8, sashrelief=tk.RAISED)
        paned.pack(fill=tk.BOTH, expand=True)

        # 左側: リスト＋スクロール＋操作ボタン
        left = tk.Frame(paned, bd=2, relief=tk.GROOVE)
        paned.add(left, stretch="always")

        vsb = tk.Scrollbar(left, orient=tk.VERTICAL)
        vsb.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        hsb = tk.Scrollbar(left, orient=tk.HORIZONTAL)
        hsb.pack(side=tk.BOTTOM, fill=tk.X, padx=5)

        self.listbox = tk.Listbox(
            left, selectmode=tk.SINGLE, font=("Arial", 12),
            yscrollcommand=vsb.set, xscrollcommand=hsb.set
        )
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=(5,0), pady=(5,0))
        self.listbox.bind("<<ListboxSelect>>", self._on_select)
        self.listbox.bind("<ButtonPress-1>", self._on_drag_start)
        self.listbox.bind("<B1-Motion>",    self._on_drag_motion)
        vsb.config(command=self.listbox.yview)
        hsb.config(command=self.listbox.xview)

        btns = [
            ("追加",   self.add_files),
            ("削除",   self.remove_selected),
            ("全削除", self.remove_all),
            ("↑",      lambda: self._move(-1)),
            ("↓",      lambda: self._move(1)),
            ("名前昇順", self.sort_by_name),
        ]
        bf = tk.Frame(left)
        bf.pack(fill=tk.X, pady=(0,5))
        for txt, cmd in btns:
            tk.Button(bf, text=txt, command=cmd, width=8).pack(side=tk.LEFT, padx=2)

        # 右側: プレビュー＋PDF化ボタン
        right = tk.Frame(paned, bd=2, relief=tk.GROOVE)
        paned.add(right, stretch="always")

        self.preview = tk.Canvas(right, bg="#f0f0f0")
        self.preview.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.preview.bind("<Configure>", self._on_canvas_resize)

        tk.Button(
            right, text="PDF化して保存",
            command=self.merge_to_pdf
        ).pack(fill=tk.X, padx=50, pady=10)

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="画像・PDFを選択",
            filetypes=[("画像／PDF", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.pdf")]
        )
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert(tk.END, os.path.basename(p))

    def remove_selected(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self.listbox.delete(idx)
        self.files.pop(idx)
        size = self.listbox.size()
        if size > 0:
            new_idx = idx if idx < size else size - 1
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(new_idx)
            self._on_select(None)
        else:
            self._clear_preview()

    def remove_all(self):
        if not self.files:
            return
        if not messagebox.askyesno("確認", "すべて削除しますか？"):
            return
        self.files.clear()
        self.listbox.delete(0, tk.END)
        self._clear_preview()

    def _move(self, offset):
        sel = self.listbox.curselection()
        if not sel:
            return
        i, j = sel[0], sel[0] + offset
        if 0 <= j < self.listbox.size():
            self.files[i], self.files[j] = self.files[j], self.files[i]
            txt = self.listbox.get(i)
            self.listbox.delete(i)
            self.listbox.insert(j, txt)
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(j)

    def sort_by_name(self):
        combined = list(zip(self.files, self.listbox.get(0, tk.END)))
        combined.sort(key=lambda x: natural_key(os.path.basename(x[0])))
        self.files = [p for p, _ in combined]
        self.listbox.delete(0, tk.END)
        for _, name in combined:
            self.listbox.insert(tk.END, name)
        self._clear_preview()

    def _on_drag_start(self, event):
        self._drag_start = self.listbox.nearest(event.y)

    def _on_drag_motion(self, event):
        i = self.listbox.nearest(event.y)
        j = self._drag_start
        if j is None or i == j:
            return
        self.files[i], self.files[j] = self.files[j], self.files[i]
        a, b = self.listbox.get(i), self.listbox.get(j)
        self.listbox.delete(j)
        self.listbox.insert(j, a)
        self.listbox.delete(i)
        self.listbox.insert(i, b)
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(i)
        self._drag_start = i

    def _on_select(self, event):
        sel = self.listbox.curselection()
        if not sel:
            return
        path = self.files[sel[0]]
        if path.lower().endswith('.pdf'):
            self._clear_preview()
            return
        try:
            img = Image.open(path)
            self._current_img = img.copy()
            self._draw_preview()
        except Exception as e:
            messagebox.showerror("プレビューエラー", str(e))

    def _on_canvas_resize(self, event):
        if self._current_img:
            self._draw_preview()

    def _draw_preview(self):
        cw, ch = self.preview.winfo_width(), self.preview.winfo_height()
        ow, oh = self._current_img.size
        ratio = min(cw/ow, ch/oh)
        nw, nh = max(1, int(ow * ratio)), max(1, int(oh * ratio))
        img = self._current_img.resize((nw, nh), Image.LANCZOS)
        tkimg = ImageTk.PhotoImage(img)
        self.preview.delete("all")
        x, y = (cw - nw) // 2, (ch - nh) // 2
        self.preview.create_image(x, y, anchor="nw", image=tkimg)
        self.preview.image = tkimg

    def _clear_preview(self):
        self._current_img = None
        self.preview.delete("all")
        self.preview.image = None

    def merge_to_pdf(self):
        if not self.files:
            messagebox.showwarning("警告", "結合するファイルがありません")
            return
        out = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            title="保存先"
        )
        if not out:
            return

        writer = PdfWriter()
        try:
            for p in self.files:
                if p.lower().endswith('.pdf'):
                    reader = PdfReader(p)
                    for page in reader.pages:
                        writer.add_page(page)
                else:
                    im = Image.open(p)
                    if im.mode in ("RGBA", "P"):
                        im = im.convert("RGB")
                    buf = io.BytesIO()
                    im.save(buf, format='PDF')
                    buf.seek(0)
                    rdr = PdfReader(buf)
                    for page in rdr.pages:
                        writer.add_page(page)

            with open(out, 'wb') as f:
                writer.write(f)
            messagebox.showinfo("完了", f"{os.path.basename(out)} を生成しました")
        except Exception as e:
            messagebox.showerror("エラー", f"PDF結合に失敗しました:\n{e}")

if __name__ == "__main__":
    ToPDFMerger().mainloop()
