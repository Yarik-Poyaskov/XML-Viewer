import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont

# Включение высокого разрешения DPI на Windows
if sys.platform == 'win32':
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

from xml_parser import XMLParser
from excel_exporter import ExcelExporter

class XMLViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("XML Viewer & Converter | Created by Yarik Poyaskov")
        self.root.geometry("1200x720")
        self.root.minsize(800, 500)

        # Переменные состояния
        self.current_file_path = None
        self.sheets = []  # List[XMLSheet]
        self.current_sheet_idx = 0
        self.current_headers = []
        self.current_rows = []
        self.filtered_rows = []
        self.sort_state = {}  # {col_idx: reverse_bool}

        # Палитра минималистичного дизайна
        self.colors = {
            "bg": "#F8F9FA",
            "card": "#FFFFFF",
            "text": "#1F2937",
            "text_muted": "#6B7280",
            "border": "#E5E7EB",
            "primary": "#2563EB",
            "primary_hover": "#1D4ED8",
            "success": "#16A34A",
            "success_hover": "#15803D",
            "tree_stripe": "#F9FAFB",
            "tree_select": "#E0E7FF",
            "tree_select_text": "#1E3A8A"
        }

        self.root.configure(bg=self.colors["bg"])

        # Шрифт для точного измерения ширины колонок
        self.header_font = tkfont.Font(font=("Segoe UI", 9, "bold"))
        self.cell_font = tkfont.Font(font=("Segoe UI", 9))

        # Настройка стилей TTK
        self._setup_styles()

        # Создание элементов интерфейса
        self._create_header_banner()
        self._create_toolbar()
        self._create_table_view()
        self._create_statusbar()

        # Проверка передачи файла через аргументы (при двойном клике в Windows)
        if len(sys.argv) > 1:
            candidate = os.path.abspath(sys.argv[1])
            if os.path.exists(candidate) and os.path.isfile(candidate):
                self._load_file(candidate)

    def _setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.style.configure(".", background=self.colors["bg"], foreground=self.colors["text"], font=("Segoe UI", 9))
        self.style.configure("Toolbar.TFrame", background=self.colors["card"], relief="flat")

        self.style.configure("Primary.TButton",
                             font=("Segoe UI", 9, "bold"),
                             background=self.colors["primary"],
                             foreground="#FFFFFF",
                             borderwidth=0,
                             focusthickness=0,
                             padding=(14, 7))
        self.style.map("Primary.TButton",
                       background=[("active", self.colors["primary_hover"])],
                       foreground=[("active", "#FFFFFF")])

        self.style.configure("Success.TButton",
                             font=("Segoe UI", 9, "bold"),
                             background=self.colors["success"],
                             foreground="#FFFFFF",
                             borderwidth=0,
                             focusthickness=0,
                             padding=(14, 7))
        self.style.map("Success.TButton",
                       background=[("active", self.colors["success_hover"])],
                       foreground=[("active", "#FFFFFF")])

        self.style.configure("Secondary.TButton",
                             font=("Segoe UI", 9),
                             background="#E5E7EB",
                             foreground="#374151",
                             borderwidth=0,
                             padding=(10, 6))

        self.style.configure("Treeview.Heading",
                             font=("Segoe UI", 9, "bold"),
                             background="#F3F4F6",
                             foreground="#374151",
                             relief="flat",
                             padding=(6, 6))
        self.style.map("Treeview.Heading",
                       background=[("active", "#E5E7EB")])

        self.style.configure("Custom.Treeview",
                             background=self.colors["card"],
                             foreground=self.colors["text"],
                             fieldbackground=self.colors["card"],
                             rowheight=26,
                             borderwidth=0)
        self.style.map("Custom.Treeview",
                       background=[("selected", self.colors["tree_select"])],
                       foreground=[("selected", self.colors["tree_select_text"])])

        self.style.configure("TCombobox", padding=5)

    def _create_header_banner(self):
        header_container = tk.Frame(self.root, bg=self.colors["card"], bd=1, relief="solid", highlightbackground=self.colors["border"], highlightthickness=1)
        header_container.pack(fill="x", padx=12, pady=(12, 0))

        banner = ttk.Frame(header_container, style="Toolbar.TFrame")
        banner.pack(fill="x", padx=12, pady=8)

        lbl_title = ttk.Label(banner, text="📄 XML Viewer & Converter", font=("Segoe UI", 12, "bold"), background=self.colors["card"], foreground=self.colors["primary"])
        lbl_title.pack(side="left")

        lbl_author = ttk.Label(banner, text="Created by Yarik Poyaskov", font=("Segoe UI", 10, "italic", "bold"), background=self.colors["card"], foreground=self.colors["text_muted"])
        lbl_author.pack(side="right")

    def _create_toolbar(self):
        toolbar_container = tk.Frame(self.root, bg=self.colors["card"], bd=1, relief="solid", highlightbackground=self.colors["border"], highlightthickness=1)
        toolbar_container.pack(fill="x", padx=12, pady=(6, 6))

        toolbar = ttk.Frame(toolbar_container, style="Toolbar.TFrame")
        toolbar.pack(fill="x", padx=10, pady=8)

        left_frame = ttk.Frame(toolbar, style="Toolbar.TFrame")
        left_frame.pack(side="left", fill="y")

        self.btn_open = ttk.Button(left_frame, text="📂 Открыть XML", style="Primary.TButton", command=self._on_open_file)
        self.btn_open.pack(side="left", padx=(0, 8))

        self.btn_export = ttk.Button(left_frame, text="📊 Выгрузить в Excel", style="Success.TButton", command=self._on_export_excel, state="disabled")
        self.btn_export.pack(side="left", padx=(0, 8))

        self.btn_fit = ttk.Button(left_frame, text="↔️ Автоширина", style="Secondary.TButton", command=self._autofit_columns)
        self.btn_fit.pack(side="left", padx=0)

        sep = ttk.Separator(toolbar, orient="vertical")
        sep.pack(side="left", fill="y", padx=14)

        lbl_sheet = ttk.Label(toolbar, text="Лист:", font=("Segoe UI", 9, "bold"), background=self.colors["card"])
        lbl_sheet.pack(side="left", padx=(0, 6))

        self.combo_sheet = ttk.Combobox(toolbar, state="readonly", width=18)
        self.combo_sheet.pack(side="left", padx=(0, 14))
        self.combo_sheet.bind("<<ComboboxSelected>>", self._on_sheet_changed)

        right_frame = ttk.Frame(toolbar, style="Toolbar.TFrame")
        right_frame.pack(side="right", fill="y")

        lbl_search = ttk.Label(right_frame, text="🔍 Поиск:", font=("Segoe UI", 9, "bold"), background=self.colors["card"])
        lbl_search.pack(side="left", padx=(0, 6))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search_query_changed)
        self.entry_search = ttk.Entry(right_frame, textvariable=self.search_var, width=24)
        self.entry_search.pack(side="left", padx=(0, 6))

        btn_clear_search = ttk.Button(right_frame, text="❌", style="Secondary.TButton", command=lambda: self.search_var.set(""))
        btn_clear_search.pack(side="left")

    def _create_table_view(self):
        table_container = tk.Frame(self.root, bg=self.colors["card"], bd=1, relief="solid", highlightbackground=self.colors["border"], highlightthickness=1)
        table_container.pack(fill="both", expand=True, padx=12, pady=6)

        self.v_scrollbar = ttk.Scrollbar(table_container, orient="vertical")
        self.h_scrollbar = ttk.Scrollbar(table_container, orient="horizontal")

        self.tree = ttk.Treeview(table_container,
                                 style="Custom.Treeview",
                                 columns=(),
                                 show="headings",
                                 yscrollcommand=self.v_scrollbar.set,
                                 xscrollcommand=self.h_scrollbar.set)

        self.v_scrollbar.config(command=self.tree.yview)
        self.h_scrollbar.config(command=self.tree.xview)

        self.v_scrollbar.pack(side="right", fill="y")
        self.h_scrollbar.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.tag_configure("evenrow", background=self.colors["card"])
        self.tree.tag_configure("oddrow", background=self.colors["tree_stripe"])

    def _create_statusbar(self):
        status_frame = tk.Frame(self.root, bg=self.colors["card"], bd=1, relief="solid", highlightbackground=self.colors["border"], highlightthickness=1)
        status_frame.pack(fill="x", side="bottom", padx=12, pady=(4, 10))

        self.lbl_status_file = ttk.Label(status_frame, text="Файл не выбран", font=("Segoe UI", 9), background=self.colors["card"], foreground=self.colors["text_muted"])
        self.lbl_status_file.pack(side="left", padx=10, pady=4)

        self.lbl_status_author = ttk.Label(status_frame, text="Created by Yarik Poyaskov", font=("Segoe UI", 9, "bold"), background=self.colors["card"], foreground=self.colors["primary"])
        self.lbl_status_author.pack(side="left", padx=20, pady=4)

        self.lbl_status_count = ttk.Label(status_frame, text="Записей: 0", font=("Segoe UI", 9, "bold"), background=self.colors["card"], foreground=self.colors["text_muted"])
        self.lbl_status_count.pack(side="right", padx=10, pady=4)

    def _on_open_file(self):
        file_path = filedialog.askopenfilename(
            title="Выберите XML-файл",
            filetypes=[("XML Файлы", "*.xml"), ("Все файлы", "*.*")]
        )
        if file_path:
            self._load_file(file_path)

    def _load_file(self, file_path):
        try:
            sheets = XMLParser.parse_file(file_path)
            if not sheets:
                messagebox.showwarning("Предупреждение", "Не удалось извлечь табличные данные из XML файла.")
                return

            self.current_file_path = file_path
            self.sheets = sheets
            self.current_sheet_idx = 0

            sheet_names = [s.name for s in self.sheets]
            self.combo_sheet.config(values=sheet_names)
            self.combo_sheet.current(0)

            self.btn_export.config(state="normal")
            self.search_var.set("")

            self._display_sheet(0)

            file_name = os.path.basename(file_path)
            self.root.title(f"XML Viewer | {file_name} | Created by Yarik Poyaskov")
            self.lbl_status_file.config(text=f"Файл: {file_name}")

        except Exception as e:
            messagebox.showerror("Ошибка парсинга", f"Произошла ошибка при открытии файла:\n{str(e)}")

    def _on_sheet_changed(self, event=None):
        idx = self.combo_sheet.current()
        if idx >= 0 and idx < len(self.sheets):
            self.current_sheet_idx = idx
            self.search_var.set("")
            self._display_sheet(idx)

    def _display_sheet(self, sheet_idx):
        sheet = self.sheets[sheet_idx]
        self.current_headers = sheet.headers
        self.current_rows = sheet.rows
        self.filtered_rows = list(self.current_rows)
        self.sort_state = {}

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = [f"col_{i}" for i in range(len(self.current_headers))]

        for i, header in enumerate(self.current_headers):
            col_id = f"col_{i}"
            self.tree.heading(col_id, text=header, command=lambda c=i: self._sort_column(c))

        self._populate_tree(self.filtered_rows)
        self._autofit_columns()

    def _autofit_columns(self):
        if not self.current_headers:
            return

        for i, header in enumerate(self.current_headers):
            col_id = f"col_{i}"
            max_px = self.header_font.measure(str(header)) + 36

            for row in self.filtered_rows:
                if i < len(row):
                    cell_text = str(row[i])
                    if cell_text:
                        w = self.cell_font.measure(cell_text) + 24
                        if w > max_px:
                            max_px = w

            col_width = min(max(max_px, 80), 1000)
            self.tree.column(col_id, width=col_width, minwidth=60, anchor="w")

    def _populate_tree(self, rows):
        self.tree.delete(*self.tree.get_children())
        for idx, row in enumerate(rows):
            tag = "evenrow" if idx % 2 == 0 else "oddrow"
            self.tree.insert("", "end", values=row, tags=(tag,))

        total = len(self.current_rows)
        shown = len(rows)
        if total == shown:
            self.lbl_status_count.config(text=f"Всего строк: {total}")
        else:
            self.lbl_status_count.config(text=f"Показано строк: {shown} из {total}")

    def _on_search_query_changed(self, *args):
        query = self.search_var.get().strip().lower()
        if not query:
            self.filtered_rows = list(self.current_rows)
        else:
            self.filtered_rows = [
                row for row in self.current_rows
                if any(query in str(cell).lower() for cell in row)
            ]
        self._populate_tree(self.filtered_rows)

    def _sort_column(self, col_idx):
        reverse = self.sort_state.get(col_idx, False)
        
        def _sort_key(row):
            val = row[col_idx] if col_idx < len(row) else ""
            try:
                return (0, float(val))
            except ValueError:
                return (1, str(val).lower())

        self.filtered_rows.sort(key=_sort_key, reverse=reverse)
        self.sort_state[col_idx] = not reverse

        for i, header in enumerate(self.current_headers):
            col_id = f"col_{i}"
            title = header
            if i == col_idx:
                title += " ▲" if reverse else " ▼"
            self.tree.heading(col_id, text=title)

        self._populate_tree(self.filtered_rows)

    def _on_export_excel(self):
        if not self.current_headers or not self.filtered_rows:
            messagebox.showwarning("Предупреждение", "Нет данных для выгрузки в Excel.")
            return

        default_filename = "export.xlsx"
        if self.current_file_path:
            base_name = os.path.splitext(os.path.basename(self.current_file_path))[0]
            default_filename = f"{base_name}.xlsx"

        output_path = filedialog.asksaveasfilename(
            title="Сохранить в Excel",
            initialfile=default_filename,
            defaultextension=".xlsx",
            filetypes=[("Excel Файлы", "*.xlsx"), ("Все файлы", "*.*")]
        )

        if output_path:
            try:
                sheet_name = self.sheets[self.current_sheet_idx].name
                ExcelExporter.export_to_excel(self.current_headers, self.filtered_rows, output_path, sheet_name)
                messagebox.showinfo("Успех", f"Данные успешно экспортированы в Excel:\n{output_path}")
            except Exception as e:
                messagebox.showerror("Ошибка экспорта", f"Не удалось сохранить Excel-файл:\n{str(e)}")

def main():
    root = tk.Tk()
    app = XMLViewerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
