#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import io
import base64
import tempfile
import atexit
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
from PyPDF2 import PdfReader, PdfWriter

ICON_BASE64 = """
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
"""

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
        self._tmp_icon_path = None
        self._set_embedded_icon()
        self.geometry("900x600")
        self.files = []
        self._drag_start = None
        self._current_img = None
        self._build_ui()
        atexit.register(self._cleanup_icon)

    def _set_embedded_icon(self):
        # プレフィックスがあれば除去
        b64 = ICON_BASE64.strip().split(",", 1)[-1]
        ico_data = base64.b64decode(b64)
        # 一時ファイルに書き出し
        tmp = tempfile.NamedTemporaryFile(
            suffix=".ico", delete=False
        )
        tmp.write(ico_data)
        tmp.close()
        try:
            self.iconbitmap(tmp.name)
            self._tmp_icon_path = tmp.name
        except Exception as e:
            print("iconbitmap 設定失敗:", e)

    def _cleanup_icon(self):
        if self._tmp_icon_path and os.path.exists(self._tmp_icon_path):
            try:
                os.remove(self._tmp_icon_path)
            except:
                pass

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
