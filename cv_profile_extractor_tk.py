import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import List

from PIL import Image, ImageTk

from cv_profile_core import build_zip, extract_best_from_paths


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("CV Profile Picture Extractor")
        self.geometry("980x700")
        self.minsize(860, 560)

        self.file_paths: List[str] = []
        self.results = []
        self.preview_refs = []

        self._build_ui()

    def _build_ui(self) -> None:
        top = ttk.Frame(self, padding=10)
        top.pack(fill=tk.X)

        ttk.Button(top, text="Add CV Files", command=self.add_files).pack(side=tk.LEFT)
        ttk.Button(top, text="Clear Files", command=self.clear_files).pack(side=tk.LEFT, padx=8)
        ttk.Button(top, text="Extract Profile Pictures", command=self.extract).pack(side=tk.LEFT, padx=8)
        ttk.Button(top, text="Save All PNGs", command=self.save_all_pngs).pack(side=tk.LEFT, padx=8)
        ttk.Button(top, text="Save ZIP", command=self.save_zip).pack(side=tk.LEFT, padx=8)

        middle = ttk.Frame(self, padding=(10, 0, 10, 10))
        middle.pack(fill=tk.BOTH, expand=True)

        left = ttk.LabelFrame(middle, text="Input CV Files", padding=8)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=False)
        self.file_list = tk.Listbox(left, width=48, height=14)
        self.file_list.pack(fill=tk.BOTH, expand=True)

        right = ttk.LabelFrame(middle, text="Extracted Profile Pictures", padding=8)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))

        self.canvas = tk.Canvas(right, borderwidth=0)
        self.scrollbar = ttk.Scrollbar(right, orient="vertical", command=self.canvas.yview)
        self.scroll_frame = ttk.Frame(self.canvas)
        self.scroll_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.status = tk.StringVar(value="Add PDF or DOCX CV files to begin.")
        ttk.Label(self, textvariable=self.status, padding=(10, 0, 10, 10)).pack(fill=tk.X)

    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Select CV files",
            filetypes=[("CV files", "*.pdf *.docx"), ("PDF files", "*.pdf"), ("Word files", "*.docx")],
        )
        if not paths:
            return
        existing = set(self.file_paths)
        for p in paths:
            if p not in existing:
                self.file_paths.append(p)
                self.file_list.insert(tk.END, p)
        self.status.set(f"{len(self.file_paths)} file(s) ready.")

    def clear_files(self) -> None:
        self.file_paths.clear()
        self.file_list.delete(0, tk.END)
        self.results = []
        self._render_results()
        self.status.set("Cleared files and results.")

    def extract(self) -> None:
        if not self.file_paths:
            messagebox.showinfo("No files", "Add at least one PDF or DOCX CV file.")
            return
        self.status.set("Extracting images and detecting faces...")
        self.update_idletasks()
        self.results = extract_best_from_paths(self.file_paths)
        self._render_results()
        self.status.set(f"Done. {len(self.results)} profile image(s) extracted.")

    def _render_results(self) -> None:
        for child in self.scroll_frame.winfo_children():
            child.destroy()
        self.preview_refs = []

        if not self.results:
            ttk.Label(self.scroll_frame, text="No extracted images yet.").pack(anchor="w")
            return

        for idx, r in enumerate(self.results):
            row = ttk.Frame(self.scroll_frame, padding=(0, 0, 0, 12))
            row.pack(fill=tk.X, anchor="w")

            thumb = r.image.copy()
            thumb.thumbnail((180, 180), Image.Resampling.LANCZOS)
            tk_img = ImageTk.PhotoImage(thumb)
            self.preview_refs.append(tk_img)

            ttk.Label(row, image=tk_img).pack(side=tk.LEFT)
            info = ttk.Frame(row)
            info.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)

            ttk.Label(info, text=r.input_name).pack(anchor="w")
            ttk.Label(info, text=r.note).pack(anchor="w")
            ttk.Label(info, text=f"Output: {r.output_name}").pack(anchor="w")
            ttk.Button(info, text="Save This PNG", command=lambda i=idx: self.save_one_png(i)).pack(anchor="w", pady=(6, 0))

    def save_one_png(self, idx: int) -> None:
        r = self.results[idx]
        path = filedialog.asksaveasfilename(
            title="Save image",
            defaultextension=".png",
            initialfile=r.output_name,
            filetypes=[("PNG image", "*.png")],
        )
        if not path:
            return
        with open(path, "wb") as f:
            f.write(r.png_bytes)
        self.status.set(f"Saved: {path}")

    def save_all_pngs(self) -> None:
        if not self.results:
            messagebox.showinfo("No results", "Extract profile pictures first.")
            return
        folder = filedialog.askdirectory(title="Select output folder")
        if not folder:
            return
        for r in self.results:
            out_path = os.path.join(folder, r.output_name)
            with open(out_path, "wb") as f:
                f.write(r.png_bytes)
        self.status.set(f"Saved {len(self.results)} PNG file(s) to {folder}")

    def save_zip(self) -> None:
        if not self.results:
            messagebox.showinfo("No results", "Extract profile pictures first.")
            return
        items = [(r.output_name, r.png_bytes) for r in self.results]
        zip_bytes = build_zip(items)
        path = filedialog.asksaveasfilename(
            title="Save ZIP",
            defaultextension=".zip",
            initialfile="cv_profile_pictures.zip",
            filetypes=[("ZIP archive", "*.zip")],
        )
        if not path:
            return
        with open(path, "wb") as f:
            f.write(zip_bytes)
        self.status.set(f"Saved ZIP: {path}")


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
