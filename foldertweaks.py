import json
import os
import re
import shutil
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

class FolderTweaksApp:
    CONFIG_DIR = Path(os.getenv('APPDATA') or Path.home() / '.config') / 'foldertweaks'
    CONFIG_FILE = CONFIG_DIR / 'templates.json'

    CATEGORY_MAP = {
        'Documents': ['pdf', 'doc', 'docx', 'txt', 'odt', 'rtf', 'xls', 'xlsx', 'ppt', 'pptx', 'csv', 'md', 'epub', 'djvu'],
        'Images': ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff', 'svg', 'webp', 'heic', 'ico', 'raw', 'psd'],
        'Videos': ['mp4', 'avi', 'mov', 'wmv', 'flv', 'mkv', 'webm', 'mpeg', 'mpg', '3gp', 'm4v', 'vob'],
        'Audio': ['mp3', 'wav', 'aac', 'ogg', 'flac', 'm4a', 'wma', 'aiff', 'alac', 'opus'],
        'Archives': ['zip', 'rar', '7z', 'tar', 'gz', 'bz2', 'xz', 'iso', 'cab', 'arj', 'lz', 'lzma'],
        'Code': ['py', 'js', 'ts', 'java', 'c', 'cpp', 'cs', 'html', 'css', 'json', 'xml', 'yml', 'yaml', 'php', 'rb', 'go', 'rs', 'sh', 'bat', 'pl', 'swift', 'kt', 'scala', 'lua', 'asm', 'sql', 'ini', 'cfg', 'toml', 'ipynb'],
        'Websites': ['html', 'htm', 'url', 'webloc', 'desktop', 'asp', 'aspx', 'php', 'jsp'],
        'Programs': ['exe', 'msi', 'bat', 'cmd', 'sh', 'bin', 'app', 'apk', 'jar', 'com', 'gadget', 'wsf', 'vbs', 'ps1'],
        'Fonts': ['ttf', 'otf', 'woff', 'woff2', 'eot', 'fon', 'fnt'],
        'Spreadsheets': ['xls', 'xlsx', 'ods', 'csv', 'tsv'],
        'Presentations': ['ppt', 'pptx', 'odp', 'key'],
        'Shortcuts': ['lnk', 'desktop', 'pif', 'url', 'webloc'],
    }

    def __init__(self, root):
        self.root = root
        self.root.title("Dateien sortieren und verwalten")
        self.root.geometry("800x650")
        self.root.minsize(800, 650)
        self.root.resizable(True, True)
        self.center_window()
        self.root.configure(bg="#f4f4f4")
        self.source_folder = tk.StringVar()
        self.target_folder = tk.StringVar()
        self.move_files = tk.BooleanVar(value=True)
        self.flatten = tk.BooleanVar(value=False)
        self.extensions = tk.StringVar()
        self.preview_data = []
        self.templates = {}
        self.selected_template = tk.StringVar()
        self.sort_folders = tk.BooleanVar(value=False)
        self.exclude_patterns = tk.StringVar()
        self._load_templates()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def center_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        ws = self.root.winfo_screenwidth()
        hs = self.root.winfo_screenheight()
        x = (ws // 2) - (w // 2)
        y = (hs // 2) - (h // 2)
        self.root.geometry(f'+{x}+{y}')

    def _build_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', font=('Segoe UI', 10), padding=6)
        style.configure('TLabel', font=('Segoe UI', 10), background='#f4f4f4')
        style.configure('TCheckbutton', background='#f4f4f4')

        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill='both', expand=True)
        frm.columnconfigure(1, weight=1)
        frm.columnconfigure(4, weight=1)

        # Template selection
        ttk.Label(frm, text="Vorlage:").grid(row=0, column=0, sticky='w')
        self.template_combo = ttk.Combobox(frm, textvariable=self.selected_template, values=list(self.templates.keys()), state='readonly', width=30)
        self.template_combo.grid(row=0, column=1, padx=5, sticky='w')
        ttk.Button(frm, text="Laden", command=self.load_template).grid(row=0, column=2, sticky='w')
        ttk.Button(frm, text="Speichern", command=self.save_template).grid(row=0, column=3, sticky='w')
        ttk.Button(frm, text="Löschen", command=self.delete_template).grid(row=0, column=4, sticky='w')

        # Source folder
        ttk.Label(frm, text="Quellordner:").grid(row=1, column=0, sticky='w')
        ttk.Entry(frm, textvariable=self.source_folder, width=30).grid(row=1, column=1, padx=5, sticky='w')
        ttk.Button(frm, text="Auswählen", command=self.select_source).grid(row=1, column=2)

        # Target folder
        ttk.Label(frm, text="Zielordner:").grid(row=2, column=0, sticky='w')
        ttk.Entry(frm, textvariable=self.target_folder, width=30).grid(row=2, column=1, padx=5, sticky='w')
        ttk.Button(frm, text="Auswählen", command=self.select_target).grid(row=2, column=2)

        # Options
        ttk.Checkbutton(frm, text="Dateien verschieben (statt kopieren)", variable=self.move_files).grid(row=3, column=0, columnspan=2, sticky='w')
        ttk.Checkbutton(frm, text="Unterordner flach machen", variable=self.flatten).grid(row=4, column=0, columnspan=2, sticky='w')
        ttk.Checkbutton(frm, text="Ordner sortieren", variable=self.sort_folders).grid(row=4, column=1, sticky='w')
        ttk.Label(frm, text="Nur folgende Endungen (z.B. jpg,png,leer für alle):").grid(row=5, column=0, sticky='w')
        ttk.Entry(frm, textvariable=self.extensions, width=30).grid(row=5, column=1, padx=5, sticky='w')
        ttk.Label(frm, text="Ausschließen:").grid(row=6, column=0, sticky='w')
        self.exclude_listbox = tk.Listbox(frm, selectmode='multiple', height=4, width=25)
        self.exclude_listbox.grid(row=6, column=1, padx=5, sticky='w')
        ttk.Button(frm, text="Hinzufügen", command=self.add_exclude_item).grid(row=6, column=2, sticky='w')
        ttk.Button(frm, text="Entfernen", command=self.remove_exclude_item).grid(row=6, column=3, sticky='w')

        # Preview and action buttons
        ttk.Button(frm, text="Vorschau", command=self.preview_sort).grid(row=8, column=0, pady=10)
        ttk.Button(frm, text="Sortieren", command=self.sort_files).grid(row=8, column=1, pady=10)

        # Preview area
        preview_frame = ttk.Frame(frm)
        preview_frame.grid(row=9, column=0, columnspan=5, pady=10, sticky='nsew')
        frm.rowconfigure(9, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        self.preview_canvas = tk.Canvas(preview_frame, bg='#fff', highlightthickness=1, highlightbackground='#ccc')
        self.preview_canvas.grid(row=0, column=0, sticky='nsew')
        self.preview_scrollbar = ttk.Scrollbar(preview_frame, orient='vertical', command=self.preview_canvas.yview)
        self.preview_scrollbar.grid(row=0, column=1, sticky='ns')
        self.preview_canvas.configure(yscrollcommand=self.preview_scrollbar.set)
        self.preview_inner = tk.Frame(self.preview_canvas, bg='#fff')
        self.preview_canvas.create_window((0, 0), window=self.preview_inner, anchor='nw')
        self.preview_inner.bind('<Configure>', lambda e: self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox('all')))

        # Status bar
        self.status = tk.Label(self.root, text="Bereit", bd=1, relief=tk.SUNKEN, anchor='w', bg='#e0e0e0', font=('Segoe UI', 9))
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

    def _load_templates(self):
        try:
            if self.CONFIG_FILE.exists():
                with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    self.templates = json.load(f)
            else:
                self.templates = {}
        except Exception:
            self.templates = {}

    def _save_templates(self):
        self.CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.templates, f, indent=2, ensure_ascii=False)

    def select_source(self):
        folder = filedialog.askdirectory()
        if folder:
            self.source_folder.set(folder)

    def select_target(self):
        folder = filedialog.askdirectory()
        if folder:
            self.target_folder.set(folder)

    def add_exclude_item(self):
        src = self.source_folder.get()
        if not src:
            messagebox.showwarning("Warnung", "Bitte zuerst einen Quellordner wählen.")
            return
        items = os.listdir(src)
        sel = tk.Toplevel(self.root)
        sel.title("Dateien/Ordner zum Ausschließen auswählen")
        sel.geometry("350x400")
        sel.transient(self.root)
        sel.grab_set()
        lb = tk.Listbox(sel, selectmode='multiple', width=40, height=20)
        lb.pack(fill='both', expand=True, padx=10, pady=10)
        for item in items:
            lb.insert('end', item)
        def on_ok():
            for idx in lb.curselection():
                val = lb.get(idx)
                if val not in self.exclude_listbox.get(0, 'end'):
                    self.exclude_listbox.insert('end', val)
            sel.destroy()
        ttk.Button(sel, text="OK", command=on_ok).pack(pady=5)

    def remove_exclude_item(self):
        for idx in reversed(self.exclude_listbox.curselection()):
            self.exclude_listbox.delete(idx)

    def get_exclude_list(self):
        return [self.exclude_listbox.get(i) for i in range(self.exclude_listbox.size())]

    def get_files(self, folder, flatten, extensions):
        result = []
        ext_list = [e.strip().lower() for e in extensions.split(',') if e.strip()] if extensions else []
        exclude_list = [e.lower() for e in self.get_exclude_list()]
        for rootdir, dirs, files in os.walk(folder):
            dirs[:] = [d for d in dirs if d.lower() not in exclude_list]
            for f in files:
                if any(ex == f.lower() for ex in exclude_list):
                    continue
                ext = os.path.splitext(f)[1][1:].lower()
                if not ext_list or ext in ext_list or (not ext and 'leer' in ext_list):
                    src = os.path.join(rootdir, f)
                    if flatten:
                        result.append((src, f))
                    else:
                        rel = os.path.relpath(rootdir, folder)
                        result.append((src, os.path.join(rel, f) if rel != '.' else f))
            if not flatten:
                break
        return result

    def is_web_url(self, link):
        return bool(re.match(r'^(https?://|www\.)', link, re.IGNORECASE))

    def is_website_file(self, ext):
        return ext.lower() in ['html', 'htm', 'url', 'webloc', 'desktop', 'asp', 'aspx', 'php', 'jsp']

    def is_program_path(self, link):
        return bool(re.match(r'.*\.(exe|bat|cmd|msi|sh|bin|app|apk|jar|com|gadget|wsf|vbs|ps1)$', link, re.IGNORECASE)) or os.path.isabs(link)

    def get_category(self, ext_or_path):
        ext = ext_or_path.lower()
        if self.is_web_url(ext_or_path) or self.is_website_file(ext):
            return 'Websites'
        if self.is_program_path(ext_or_path):
            return 'Programs'
        for cat, exts in self.CATEGORY_MAP.items():
            if ext in exts:
                return cat
        return 'Other'

    def preview_sort(self):
        for widget in self.preview_inner.winfo_children():
            widget.destroy()
        src = self.source_folder.get()
        tgt = self.target_folder.get()
        flatten = self.flatten.get()
        exts = self.extensions.get()
        if not src or not tgt:
            self.status['text'] = "Bitte Quell- und Zielordner wählen."
            return
        files = self.get_files(src, flatten, exts)
        self.preview_data = []
        row = 0
        icon_map = {
            'Documents': '\U0001F4C4',  # 📄
            'Images': '\U0001F5BC',     # 🖼️
            'Videos': '\U0001F3AC',     # 🎬
            'Audio': '\U0001F3B5',      # 🎵
            'Archives': '\U0001F4E6',   # 📦
            'Code': '\U0001F4BB',       # 💻
            'Websites': '\U0001F310',   # 🌐
            'Programs': '\U0001F5A5',   # 🖥️
            'Fonts': '\U0001F58B',      # 🖋️
            'Spreadsheets': '\U0001F4C8', # 📈
            'Presentations': '\U0001F4FD', # 📽️
            'PDFs': '\U0001F4D6',       # 📖
            'Shortcuts': '\U0001F517',  # 🔗
            'Other': '\U0001F5CE',      # 🗎
        }
        exclude_list = [e.lower() for e in self.get_exclude_list()]
        for src_path, rel_name in files:
            ext = os.path.splitext(rel_name)[1][1:] or "ohne_endung"
            category = self.get_category(ext)
            cat_folder = os.path.join(tgt, category)
            target_path = os.path.join(cat_folder, os.path.basename(rel_name))
            icon = icon_map.get(category, icon_map['Other'])
            label = tk.Label(self.preview_inner, text=f"{icon} {os.path.basename(src_path)} → {target_path}", anchor='w', bg='#fff', font=('Segoe UI Emoji', 10))
            label.grid(row=row, column=0, sticky='w', padx=2, pady=1)
            self.preview_data.append((src_path, target_path))
            row += 1
        if self.sort_folders.get():
            folder_dest = os.path.join(tgt, "folder")
            for entry in os.listdir(src):
                if any(ex == entry.lower() for ex in exclude_list):
                    continue
                full_path = os.path.join(src, entry)
                if os.path.isdir(full_path) and os.path.abspath(full_path) != os.path.abspath(tgt):
                    icon = '\U0001F4C1'  # 📁
                    target_folder = os.path.join(folder_dest, entry)
                    label = tk.Label(self.preview_inner, text=f"{icon} [Ordner] {entry} → {target_folder}", anchor='w', bg='#fff', font=('Segoe UI Emoji', 10, 'bold'))
                    label.grid(row=row, column=0, sticky='w', padx=2, pady=1)
                    self.preview_data.append((full_path, target_folder, 'folder'))
                    row += 1
        self.status['text'] = f"{len(files)} Dateien gefunden." + (f" | {len([x for x in self.preview_data if len(x)==3 and x[2]=='folder'])} Ordner gefunden." if self.sort_folders.get() else "")

    def sort_files(self):
        if not self.preview_data:
            self.preview_sort()
        if not self.preview_data:
            self.status['text'] = "Keine Dateien zum Sortieren gefunden."
            return
        move = self.move_files.get()
        for item in self.preview_data:
            if len(item) == 3 and item[2] == 'folder':
                src, tgt, _ = item
                if os.path.exists(tgt):
                    continue
                try:
                    if move:
                        shutil.move(src, tgt)
                    else:
                        shutil.copytree(src, tgt)
                except Exception as e:
                    self.status['text'] = f"Fehler beim Ordner: {e}"
                    messagebox.showerror("Fehler", f"Fehler bei Ordner {src}: {e}")
                    return
            else:
                src, tgt = item[:2]
                os.makedirs(os.path.dirname(tgt), exist_ok=True)
                try:
                    if move:
                        shutil.move(src, tgt)
                    else:
                        shutil.copy2(src, tgt)
                except Exception as e:
                    self.status['text'] = f"Fehler: {e}"
                    messagebox.showerror("Fehler", f"Fehler bei {src}: {e}")
                    return
        self.status['text'] = "Sortierung abgeschlossen."
        messagebox.showinfo("Fertig", "Dateien und Ordner wurden sortiert.")
        self.preview_data = []
        self.preview_canvas.config(scrollregion=self.preview_canvas.bbox('all'))

    def load_template(self):
        name = self.selected_template.get()
        if name and name in self.templates:
            t = self.templates[name]
            self.source_folder.set(t.get('source_folder', ''))
            self.target_folder.set(t.get('target_folder', ''))
            self.move_files.set(t.get('move_files', True))
            self.flatten.set(t.get('flatten', False))
            self.extensions.set(t.get('extensions', ''))
            self.exclude_listbox.delete(0, 'end')
            for ex in t.get('exclude_patterns', '').split(','):
                if ex:
                    self.exclude_listbox.insert('end', ex)
            self.sort_folders.set(t.get('sort_folders', False))
            self.status['text'] = f"Vorlage '{name}' geladen."

    def save_template(self):
        name = self.selected_template.get()
        if not name:
            name = filedialog.asksaveasfilename(title="Vorlagenname eingeben", defaultextension=".json", filetypes=[("JSON Dateien", "*.json")])
            if not name:
                return
            name = os.path.splitext(os.path.basename(name))[0]
        self.templates[name] = {
            'source_folder': self.source_folder.get(),
            'target_folder': self.target_folder.get(),
            'move_files': self.move_files.get(),
            'flatten': self.flatten.get(),
            'extensions': self.extensions.get(),
            'exclude_patterns': ','.join(self.get_exclude_list()),
            'sort_folders': self.sort_folders.get()
        }
        self._save_templates()
        self.template_combo['values'] = list(self.templates.keys())
        self.selected_template.set(name)
        self.status['text'] = f"Vorlage '{name}' gespeichert."

    def delete_template(self):
        name = self.selected_template.get()
        if name and name in self.templates:
            del self.templates[name]
            self._save_templates()
            self.template_combo['values'] = list(self.templates.keys())
            self.selected_template.set('')
            self.status['text'] = f"Vorlage '{name}' gelöscht."

    def on_close(self):
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = FolderTweaksApp(root)
    root.mainloop()