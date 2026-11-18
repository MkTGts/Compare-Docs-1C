# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re
import os
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill
import xlrd

def extract_id_real(text):
    """Извлекает числовой номер документа из строки 1С, игнорируя нули и префиксы."""
    match = re.search(r'\s+([А-Яа-яA-Za-z-]*0*(\d+))\s+от\s+', str(text))
    if match:
        return match.group(2).lstrip('0')
    return None

def read_workbook_sheets(path, description):
    """Читает все листы книги .xlsx или .xls."""
    extension = Path(path).suffix.lower()
    try:
        if extension == ".xlsx":
            workbook = load_workbook(path, data_only=True, read_only=True)
            try:
                return [
                    (sheet.title, [list(row) for row in sheet.iter_rows(values_only=True)])
                    for sheet in workbook.worksheets
                ]
            finally:
                workbook.close()
        if extension == ".xls":
            workbook = xlrd.open_workbook(path)
            return [
                (sheet.name, [sheet.row_values(row_idx) for row_idx in range(sheet.nrows)])
                for sheet in workbook.sheets()
            ]
    except Exception as error:
        raise ValueError(
            f"Не удалось прочитать {description} «{path}» как файл {extension}. "
            "Проверьте соответствие расширения содержимому и целостность книги.\n"
            f"Подробности: {error}"
        ) from error
    raise ValueError(f"Формат {description} не поддерживается. Используйте .xls или .xlsx.")

def normalized_header(value):
    return re.sub(r"[^a-zа-яё0-9]", "", str(value or "").lower())

def column_letter(index):
    result = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result

def guess_columns(kind, rows):
    """Предлагает столбцы по заголовкам, затем по содержимому первых строк."""
    count = max((len(row) for row in rows), default=0)
    if count == 0:
        return (0, 1) if kind == "bp" else (1, 2), False

    first_row = rows[0] if rows else []
    aliases = {
        "bp": {
            "number": {"номер", "номдокумента", "номердокумента", "документ", "№", "docnumber"},
            "amount": {"сумма", "суммадокумента", "итого", "суммабп", "amount", "total"},
        },
        "real": {
            "text": {"продажи", "документ", "наименование", "содержание", "основание", "операция", "text"},
            "amount": {"сумма", "суммадокумента", "итого", "суммареал", "amount", "total"},
        },
    }[kind]
    header_matches = {}
    for field, names in aliases.items():
        header_matches[field] = next(
            (i for i, value in enumerate(first_row) if normalized_header(value) in names), None
        )

    has_header = any(index is not None for index in header_matches.values())
    data_rows = rows[1:] if has_header else rows
    if kind == "bp":
        number_col = header_matches["number"]
        amount_col = header_matches["amount"]
        if number_col is None:
            number_col = max(range(count), key=lambda col: sum(
                1 for row in data_rows[:100] if col < len(row) and re.search(r"\d+", str(row[col] or ""))
            ))
        if amount_col is None:
            amount_col = max(range(count), key=lambda col: sum(
                1 for row in data_rows[:100] if col < len(row) and is_number(row[col])
            ))
        if amount_col == number_col and count > 1:
            amount_col = 1 if number_col != 1 else 0
        return (number_col, amount_col), has_header

    text_col = header_matches["text"]
    if text_col is None:
        text_col = max(range(count), key=lambda col: sum(
            1 for row in data_rows[:100]
            if col < len(row) and "реализац" in str(row[col] or "").lower()
        ))
    amount_col = header_matches["amount"]
    if amount_col is None:
        amount_col = max(range(count), key=lambda col: sum(
            1 for row in data_rows[:100] if col != text_col and col < len(row) and is_number(row[col])
        ))
    if amount_col == text_col and count > 1:
        amount_col = 2 if text_col != 2 and count > 2 else (1 if text_col != 1 else 0)
    return (text_col, amount_col), has_header

def is_number(value):
    if value is None or value == "":
        return False
    try:
        float(value)
        return True
    except (TypeError, ValueError):
        return False

def inspect_sheet(sheet_index, sheet_name, rows, kind):
    """Проверяет, можно ли автоматически использовать лист без ручной настройки."""
    columns, has_header = guess_columns(kind, rows)
    if len(set(columns)) != len(columns):
        return None

    start = 1 if has_header else 0
    valid_records = 0
    issues = []
    for row_num, row in enumerate(rows[start:], start=start + 1):
        first_col, amount_col = columns
        first_value = row[first_col] if first_col < len(row) else None
        amount_value = row[amount_col] if amount_col < len(row) else None

        if kind == "bp":
            if first_value in (None, "") and amount_value in (None, ""):
                continue
            match = re.search(r"(\d+)", str(first_value)) if first_value not in (None, "") else None
            if not match:
                issues.append(f"строка {row_num}: не распознан номер документа")
                continue
        else:
            if not isinstance(first_value, str) or "реализация" not in first_value.lower():
                continue
            if not extract_id_real(first_value):
                issues.append(f"строка {row_num}: не распознан номер реализации")
                continue

        if amount_value not in (None, ""):
            try:
                float(amount_value)
            except (TypeError, ValueError):
                issues.append(f"строка {row_num}: сумма не распознана как число")
                continue
        valid_records += 1

    if valid_records == 0:
        return None
    return {
        "sheet_index": sheet_index,
        "rows": rows,
        "columns": list(columns),
        "skip_header": has_header,
        "valid_records": valid_records,
        "issues": issues,
        "sheet_name": sheet_name,
    }

def find_auto_selection(sheets, kind):
    """Возвращает автонастройку, если подходящий лист найден однозначно и есть валидные строки."""
    candidates = []
    for sheet_index, (sheet_name, rows) in enumerate(sheets):
        candidate = inspect_sheet(sheet_index, sheet_name, rows, kind)
        if candidate:
            candidates.append(candidate)
    if len(candidates) == 1:
        return candidates[0]
    return None

def choose_columns(bp_sheets, real_sheets):
    """Показывает автоопределение, выбор листов/столбцов и предпросмотр."""
    dialog = tk.Toplevel(root)
    dialog.title("Проверьте распознанные столбцы")
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    width = max(520, min(820, screen_width - 40))
    height = max(480, min(650, screen_height - 80))
    x = max(0, (screen_width - width) // 2)
    y = max(0, (screen_height - height) // 2)
    dialog.geometry(f"{width}x{height}+{x}+{y}")
    dialog.minsize(min(width, 620), min(height, 520))
    dialog.transient(root)

    result = {"value": None}
    panels = {}

    def make_panel(parent, key, title, sheets, kind, fields):
        box = ttk.LabelFrame(parent, text=title, padding=10)
        box.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)
        sheet_var = tk.StringVar(value=sheets[0][0] if sheets else "")
        ttk.Label(box, text="Лист:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        sheet_combo = ttk.Combobox(box, textvariable=sheet_var, state="readonly",
                                   values=[name for name, _ in sheets], width=36)
        sheet_combo.grid(row=0, column=1, columnspan=3, sticky="ew", padx=4, pady=4)

        skip_var = tk.BooleanVar(value=False)
        field_vars = {field: tk.StringVar() for field, _ in fields}
        field_combos = {}
        for row_index, (field, label) in enumerate(fields, start=1):
            ttk.Label(box, text=label + ":").grid(row=row_index, column=0, sticky="w", padx=4, pady=4)
            combo = ttk.Combobox(box, textvariable=field_vars[field], state="readonly", width=48)
            combo.grid(row=row_index, column=1, columnspan=3, sticky="ew", padx=4, pady=4)
            field_combos[field] = combo

        checkbox_row = len(fields) + 1
        ttk.Checkbutton(box, text="Пропустить первую строку (это заголовки)", variable=skip_var,
                        command=lambda: update_preview(key)).grid(
                            row=checkbox_row, column=0, columnspan=4, sticky="w", padx=4, pady=4)
        reset_row = checkbox_row + 1
        reset_label = "Автоподбор столбцов БП" if kind == "bp" else "Автоподбор столбцов УТ"
        ttk.Button(box, text=reset_label, command=lambda panel_key=key: refresh_panel(panel_key)).grid(
            row=reset_row, column=0, columnspan=2, sticky="w", padx=4, pady=4)
        preview_row = reset_row + 1
        ttk.Label(box, text="Предпросмотр первых строк:").grid(
            row=preview_row, column=0, columnspan=4, sticky="w", padx=4, pady=(6, 2))
        preview = tk.Text(box, height=4, wrap=tk.WORD, font=("Consolas", 9), state=tk.DISABLED)
        preview.grid(row=preview_row + 1, column=0, columnspan=4, sticky="ew", padx=4, pady=4)
        box.columnconfigure(1, weight=1)
        panels[key] = {
            "sheets": sheets, "kind": kind, "sheet_var": sheet_var,
            "sheet_combo": sheet_combo, "skip_var": skip_var,
            "field_vars": field_vars, "field_combos": field_combos, "preview": preview,
        }
        sheet_combo.bind("<<ComboboxSelected>>", lambda _event, panel_key=key: refresh_panel(panel_key))
        refresh_panel(key)

    def selected_sheet(panel):
        selected_name = panel["sheet_var"].get()
        return next(((i, rows) for i, (name, rows) in enumerate(panel["sheets"]) if name == selected_name), (0, []))

    def update_preview(key):
        panel = panels[key]
        _, rows = selected_sheet(panel)
        start = 1 if panel["skip_var"].get() else 0
        shown = rows[start:start + 5]
        lines = []
        for row_num, values in enumerate(shown, start=start + 1):
            cells = [str(value).replace("\n", " ")[:36] if value not in (None, "") else "" for value in values[:8]]
            lines.append(f"{row_num}: " + " | ".join(cells))
        if not lines:
            lines = ["На выбранном листе нет строк для предпросмотра."]
        preview = panel["preview"]
        preview.config(state=tk.NORMAL)
        preview.delete("1.0", tk.END)
        preview.insert("1.0", "\n".join(lines))
        preview.config(state=tk.DISABLED)

    def refresh_panel(key):
        panel = panels[key]
        _, rows = selected_sheet(panel)
        guesses, has_header = guess_columns(panel["kind"], rows)
        panel["skip_var"].set(has_header)
        count = max((len(row) for row in rows), default=0)
        first_row = rows[0] if rows else []
        choices = [
            f"{column_letter(i)} — {str(first_row[i]).strip()[:36] if i < len(first_row) and first_row[i] not in (None, '') else 'столбец ' + column_letter(i)}"
            for i in range(count)
        ]
        for position, (field, _label) in enumerate(panel["field_combos"].items()):
            combo = panel["field_combos"][field]
            combo["values"] = choices
            if choices:
                combo.current(min(guesses[position], len(choices) - 1))
            else:
                panel["field_vars"][field].set("")
        update_preview(key)

    def save_selection():
        selections = {}
        for key, panel in panels.items():
            sheet_index, rows = selected_sheet(panel)
            chosen = []
            for field, _label in panel["field_combos"].items():
                index = panel["field_combos"][field].current()
                if index < 0:
                    messagebox.showwarning("Не выбран столбец", f"Выберите столбцы для раздела «{key}».", parent=dialog)
                    return
                chosen.append(index)
            if len(set(chosen)) != len(chosen):
                messagebox.showwarning("Проверьте столбцы", "Для разных данных выбраны одинаковые столбцы.", parent=dialog)
                return
            selections[key] = {
                "sheet_index": sheet_index,
                "rows": rows,
                "columns": chosen,
                "skip_header": panel["skip_var"].get(),
            }
        result["value"] = selections
        dialog.destroy()

    def cancel():
        dialog.destroy()

    ttk.Label(dialog, text="Автоподбор не смог однозначно проверить входные данные. Исправьте выбор отдельно для каждой таблицы и сверьте его с предпросмотром.",
              wraplength=width - 40).pack(anchor="w", padx=14, pady=(8, 2))

    buttons = ttk.Frame(dialog)
    buttons.pack(side=tk.BOTTOM, fill=tk.X, padx=14, pady=8)
    ttk.Button(buttons, text="Открыть инструкцию", command=lambda: show_instructions(dialog)).pack(side=tk.LEFT, padx=4)
    ttk.Button(buttons, text="Отмена", command=cancel).pack(side=tk.RIGHT, padx=4)
    ttk.Button(buttons, text="Продолжить сверку", command=save_selection).pack(side=tk.RIGHT, padx=4)

    content = ttk.Frame(dialog)
    content.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=4, pady=2)
    canvas = tk.Canvas(content, highlightthickness=0)
    scrollbar = ttk.Scrollbar(content, orient=tk.VERTICAL, command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    panels_frame = ttk.Frame(canvas)
    canvas_window = canvas.create_window((0, 0), window=panels_frame, anchor="nw")

    def update_scroll_region(_event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def resize_panels(event):
        canvas.itemconfigure(canvas_window, width=event.width)

    panels_frame.bind("<Configure>", update_scroll_region)
    canvas.bind("<Configure>", resize_panels)
    make_panel(panels_frame, "bp", "Документы БП", bp_sheets, "bp", [("number", "Номер документа"), ("amount", "Сумма")])
    make_panel(panels_frame, "real", "Документы УТ / реализации", real_sheets, "real", [("text", "Текст документа"), ("amount", "Сумма")])
    dialog.grab_set()
    root.wait_window(dialog)
    return result["value"]

def run_comparison():
    file_bp = entry_bp.get().strip()
    file_real = entry_real.get().strip()
    output_file = entry_out.get().strip()

    path_errors = []
    if not file_bp:
        path_errors.append("Не выбран файл БП.")
    elif not Path(file_bp).is_file():
        path_errors.append(f"Файл БП не найден или это не файл:\n{file_bp}")
    elif Path(file_bp).suffix.lower() not in (".xls", ".xlsx"):
        path_errors.append("Файл БП должен иметь формат .xls или .xlsx.")

    if not file_real:
        path_errors.append("Не выбран файл реализаций.")
    elif not Path(file_real).is_file():
        path_errors.append(f"Файл реализаций не найден или это не файл:\n{file_real}")
    elif Path(file_real).suffix.lower() not in (".xls", ".xlsx"):
        path_errors.append("Файл реализаций должен иметь формат .xls или .xlsx.")

    if not output_file:
        path_errors.append("Не указан путь для сохранения итогового файла.")
    else:
        out_path = Path(output_file)
        if out_path.suffix.lower() != ".xlsx":
            path_errors.append("Итоговый файл должен иметь расширение .xlsx.")
        elif not out_path.parent.exists():
            path_errors.append(f"Папка для сохранения не существует:\n{out_path.parent}")
        elif (file_bp and os.path.normcase(os.path.abspath(file_bp)) == os.path.normcase(os.path.abspath(output_file))) or (file_real and os.path.normcase(os.path.abspath(file_real)) == os.path.normcase(os.path.abspath(output_file))):
            path_errors.append("Путь итогового файла совпадает с одним из исходных файлов. Укажите другой путь.")

    if path_errors:
        messagebox.showwarning("Проверьте выбранные файлы", "Исправьте следующие пункты:\n\n• " + "\n• ".join(path_errors))
        return

    try:
        bp_sheets = read_workbook_sheets(file_bp, "файл БП")
        real_sheets = read_workbook_sheets(file_real, "файл реализаций")
    except Exception as error:
        messagebox.showerror("Ошибка чтения", str(error))
        return

    if not bp_sheets:
        messagebox.showwarning("Пустая книга", "В файле БП нет листов.")
        return
    if not real_sheets:
        messagebox.showwarning("Пустая книга", "В файле реализаций нет листов.")
        return

    bp_auto = find_auto_selection(bp_sheets, "bp")
    real_auto = find_auto_selection(real_sheets, "real")
    if bp_auto and real_auto:
        selected = {"bp": bp_auto, "real": real_auto}
    else:
        selected = choose_columns(bp_sheets, real_sheets)
        if selected is None:
            return

    bp_selection = selected["bp"]
    real_selection = selected["real"]
    bp_rows = bp_selection["rows"]
    real_rows = real_selection["rows"]
    col_nomer, col_summa = bp_selection["columns"]
    col_real_text, col_real_sum = real_selection["columns"]
    bp_start = 1 if bp_selection["skip_header"] else 0
    real_start = 1 if real_selection["skip_header"] else 0

    btn_run.config(text="Обработка...", state=tk.DISABLED)
    root.update()

    try:
        # 1. Чтение выбранных столбцов таблицы БП

        bp_data = {}
        bp_row_errors = []
        for excel_row, row in enumerate(bp_rows[bp_start:], start=bp_start + 1):
            nomer = row[col_nomer] if col_nomer < len(row) else None
            summa = row[col_summa] if col_summa < len(row) else None
            if nomer in (None, "") and summa in (None, ""):
                continue
            match = re.search(r'(\d+)', str(nomer)) if nomer not in (None, "") else None
            if not match:
                document_label = nomer if nomer not in (None, "") else "не указан"
                bp_row_errors.append(f"БП, строка {excel_row}, документ «{document_label}»: номер не содержит цифр.")
                continue
            id_clean = match.group(1).lstrip('0')
            try:
                s = float(summa) if summa not in (None, "") else 0.0
            except (TypeError, ValueError):
                bp_row_errors.append(f"БП, строка {excel_row}, документ «{nomer}»: сумма «{summa}» не является числом.")
                continue
            bp_data[id_clean] = {'Номер': str(nomer), 'Сумма_БП': s}

        if not bp_data:
            bp_row_errors.append("БП: не найдено ни одной строки с распознаваемым номером документа.")

        # 2. Чтение выбранных столбцов таблицы реализаций
        real_data = {}
        real_row_errors = []
        for row_idx, row in enumerate(real_rows[real_start:], start=real_start + 1):
            text_val = row[col_real_text] if col_real_text < len(row) else None
            sum_val = row[col_real_sum] if col_real_sum < len(row) else None
            if not isinstance(text_val, str) or 'Реализация' not in text_val:
                continue
            id_clean = extract_id_real(text_val)
            if not id_clean:
                real_row_errors.append(f"УТ, строка {row_idx}, документ «{text_val}»: не удалось извлечь номер перед словом «от».")
                continue
            try:
                s = float(sum_val) if sum_val not in ('', None) else 0.0
            except (TypeError, ValueError):
                real_row_errors.append(f"УТ, строка {row_idx}, документ «{text_val}»: сумма «{sum_val}» не является числом.")
                continue
            real_data[id_clean] = {'Text': text_val, 'Сумма_Реал': s}

        if not real_data:
            real_row_errors.append("Реализации: не найдено документов с текстом «Реализация» и номером, который можно распознать.")

        data_errors = bp_row_errors + real_row_errors

        # 3. Сверка данных
        all_ids = set(bp_data.keys()).union(set(real_data.keys()))
        report_data = []
        only_real_data = []
        only_bp_data = []
        
        for doc_id in all_ids:
            b_info = bp_data.get(doc_id, {'Номер': 'Отсутствует в БП', 'Сумма_БП': 0.0})
            r_info = real_data.get(doc_id, {'Text': 'Отсутствует в выгрузке реализаций', 'Сумма_Реал': 0.0})
            
            diff = round(b_info['Сумма_БП'] - r_info['Сумма_Реал'], 2)
            row_data = [
                b_info['Номер'],
                b_info['Сумма_БП'],
                r_info['Text'],
                r_info['Сумма_Реал'],
                diff,
            ]
            if doc_id not in bp_data:
                only_real_data.append(row_data)
            elif doc_id not in real_data:
                only_bp_data.append(row_data)
            if diff != 0:
                report_data.append(row_data)

        # 4. Формирование Excel отчета
        wb_out = Workbook()
        headers = ['Номер документа (из БП)', 'Сумма (БП)', 'Документ (из Окт. Реал)', 'Сумма (Окт. Реал)', 'Разница (БП - Реал)']
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)

        def add_report_sheet(title, rows):
            sheet = wb_out.active if title == "Расхождения" else wb_out.create_sheet(title=title)
            sheet.title = title
            sheet.append(headers)
            for row_data in rows:
                sheet.append(row_data)

            for cell in sheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for col in sheet.columns:
                max_length = max((len(str(cell.value)) for cell in col if cell.value is not None), default=0)
                letter = col[0].column_letter
                for cell in col:
                    if cell.row > 1 and cell.column in [2, 4, 5]:
                        cell.number_format = '#,##0.00'
                sheet.column_dimensions[letter].width = min(max_length + 2, 60)

        add_report_sheet("Расхождения", report_data)
        add_report_sheet("Есть в УТ, нет в БП", only_real_data)
        add_report_sheet("Есть в БП, нет в УТ", only_bp_data)

        wb_out.save(output_file)
        summary = (
            f"Результат сохранен: {output_file}\n"
            f"Расхождений: {len(report_data)}; только в УТ: {len(only_real_data)}; "
            f"только в БП: {len(only_bp_data)}."
        )
        if data_errors:
            shown_errors = data_errors[:20]
            if len(data_errors) > len(shown_errors):
                shown_errors.append(f"…и еще {len(data_errors) - len(shown_errors)} пропущенных строк/проблем.")
            messagebox.showwarning(
                "Сверка завершена с пропусками",
                summary + "\n\nНекоторые строки пропущены или не содержат распознаваемых документов:\n\n• "
                + "\n• ".join(shown_errors),
            )
        else:
            messagebox.showinfo("Готово", summary)

    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка при обработке данных:\n{e}")
    finally:
        btn_run.config(text="Проверить и сверить", state=tk.NORMAL)

def select_file(entry):
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

def select_save_file(entry):
    path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                        initialfile="Расхождения_сверки.xlsx",
                                        filetypes=[("Excel files", "*.xlsx")])
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

def show_instructions(parent=None):
    owner = parent or root
    window = tk.Toplevel(owner)
    window.title("Инструкция по работе")
    window.geometry("640x520")
    window.minsize(500, 380)
    window.transient(owner)

    container = tk.Frame(window, padx=16, pady=14)
    container.pack(fill=tk.BOTH, expand=True)
    scrollbar = tk.Scrollbar(container)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    instructions = tk.Text(container, wrap=tk.WORD, yscrollcommand=scrollbar.set,
                           padx=10, pady=10, font=("Arial", 10), relief=tk.FLAT)
    instructions.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=instructions.yview)

    instructions.insert("1.0", """КАК СВЕРИТЬ ДОКУМЕНТЫ

1. Выберите файл БП и файл реализаций. Поддерживаются Excel .xls и .xlsx.
2. Укажите имя и папку для итогового файла .xlsx.
3. Нажмите «Проверить и сверить».

Какие данные нужны
• В БП должны быть номер документа и сумма. Заголовки «Номер» и «Сумма» помогают программе найти столбцы автоматически.
• В реализации нужен текст документа со словом «Реализация» и номером перед «от» (например, «Реализация ... 000123 от 01.01.2026»), а также сумма.
• Листы и расположение столбцов могут отличаться. Программа попробует распознать их сама.

Как вывести отчет в БП
• Открыть Реализации
• Установить нужный период (Еще -> Установить период...)
• Сделать отбор. (Еще -> Настроить список...-> Отбор) Контрагент: Папки Собств.нужды, Удаленные, Собственные организации, ООО ТД Палитра. Номер: Содержит ЛИУТ
• Еще -> Вывести список. Проставить галки Номер и Сумма и сохранить файл в формате .xlsx 

Если открылось окно проверки
Выберите для БП и реализаций лист, номер/текст и сумму. Сверьтесь с предпросмотром. Если в первой строке заголовки, оставьте включенным флажок «Пропустить первую строку». Автоподбор для БП и УТ настраивается отдельно. Нажмите «Продолжить сверку».

Если окно проверки не открылось, программа распознала обе таблицы и сразу начала сверку.

Что будет в результате
Файл содержит три листа:
• «Расхождения» — документы с разницей в суммах.
• «Есть в УТ, нет в БП» — документы только из УТ.
• «Есть в БП, нет в УТ» — документы только из БП.

Номера сравниваются без ведущих нулей. Отсутствующая сумма считается нулевой, разница округляется до двух знаков. Строки с ошибками пропускаются: сверка продолжится, а после сохранения отчета появится предупреждение с номером строки и документом.""")
    instructions.config(state=tk.DISABLED)
    tk.Button(window, text="Закрыть", command=window.destroy).pack(pady=(0, 12))

root = tk.Tk()
root.title("Сверка документов 1С v1.1")
root.geometry("540x230")
root.resizable(False, False)

frame = tk.Frame(root, padx=15, pady=15)
frame.pack(fill=tk.BOTH, expand=True)

tk.Label(frame, text="Файл из БП:").grid(row=0, column=0, sticky="w", pady=5)
entry_bp = tk.Entry(frame, width=40)
entry_bp.grid(row=0, column=1, padx=5, pady=5)
tk.Button(frame, text="Выбрать", command=lambda: select_file(entry_bp)).grid(row=0, column=2, pady=5)

tk.Label(frame, text="Файл из УТ:").grid(row=1, column=0, sticky="w", pady=5)
entry_real = tk.Entry(frame, width=40)
entry_real.grid(row=1, column=1, padx=5, pady=5)
tk.Button(frame, text="Выбрать", command=lambda: select_file(entry_real)).grid(row=1, column=2, pady=5)

tk.Label(frame, text="Сохранить итог:").grid(row=2, column=0, sticky="w", pady=5)
entry_out = tk.Entry(frame, width=40)
entry_out.grid(row=2, column=1, padx=5, pady=5)
tk.Button(frame, text="Выбрать", command=lambda: select_save_file(entry_out)).grid(row=2, column=2, pady=5)

btn_run = tk.Button(frame, text="Проверить и сверить", command=run_comparison, bg="#4F81BD", fg="white", font=("Arial", 10, "bold"))
btn_run.grid(row=3, column=0, columnspan=2, pady=15, ipadx=10, ipady=3)
tk.Button(frame, text="Инструкция", command=show_instructions).grid(row=3, column=2, pady=15, ipadx=8, ipady=3)

root.mainloop()
