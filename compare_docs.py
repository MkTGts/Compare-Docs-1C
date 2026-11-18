import tkinter as tk
from tkinter import filedialog, messagebox
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

def read_first_sheet(path, description):
    """Читает первый лист .xlsx или .xls и возвращает строки как списки значений."""
    extension = Path(path).suffix.lower()
    try:
        if extension == ".xlsx":
            workbook = load_workbook(path, data_only=True, read_only=True)
            try:
                return [list(row) for row in workbook.active.iter_rows(values_only=True)]
            finally:
                workbook.close()
        if extension == ".xls":
            workbook = xlrd.open_workbook(path)
            sheet = workbook.sheet_by_index(0)
            return [sheet.row_values(row_idx) for row_idx in range(sheet.nrows)]
    except Exception as error:
        raise ValueError(
            f"Не удалось прочитать {description} «{path}» как файл {extension}. "
            "Проверьте соответствие расширения содержимому и целостность книги.\n"
            f"Подробности: {error}"
        ) from error
    raise ValueError(f"Формат {description} не поддерживается. Используйте .xls или .xlsx.")

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

    btn_run.config(text="Обработка...", state=tk.DISABLED)
    root.update()

    try:
        # 1. Чтение файла БП
        bp_rows = read_first_sheet(file_bp, "файл БП")
        if not bp_rows:
            raise ValueError("Файл БП пуст: на первом листе нет строки заголовков.")
        headers = bp_rows[0]
        missing_headers = [name for name in ('Номер', 'Сумма') if name not in headers]
        if missing_headers:
            raise ValueError(
                "В файле БП на первом листе не найдены обязательные заголовки: "
                + ', '.join(f'«{name}»' for name in missing_headers)
                + f". Найденные заголовки: {headers}"
            )
        col_nomer = headers.index('Номер')
        col_summa = headers.index('Сумма')

        bp_data = {}
        bp_row_errors = []
        for excel_row, row in enumerate(bp_rows[1:], start=2):
            if len(row) > max(col_nomer, col_summa):
                nomer = row[col_nomer]
                summa = row[col_summa]
                if nomer is None and summa is None:
                    continue
                match = re.search(r'(\d+)', str(nomer)) if nomer is not None else None
                if not match:
                    bp_row_errors.append(f"БП, строка {excel_row}: в столбце «Номер» нет номера с цифрами.")
                    continue
                id_clean = match.group(1).lstrip('0')
                try:
                    s = float(summa) if summa is not None else 0.0
                except (TypeError, ValueError):
                    bp_row_errors.append(f"БП, строка {excel_row}: сумма «{summa}» не является числом.")
                    continue
                bp_data[id_clean] = {'Номер': str(nomer), 'Сумма_БП': s}

        if not bp_data:
            bp_row_errors.append("БП: не найдено ни одной строки с распознаваемым номером документа.")

        # 2. Чтение файла Реализаций
        real_rows = read_first_sheet(file_real, "файл реализаций")
        real_column_count = max((len(row) for row in real_rows), default=0)
        if real_column_count < 3:
            raise ValueError(
                f"В файле реализаций на первом листе только {real_column_count} столбца(ов). "
                "Нужны как минимум столбцы B (текст документа) и C (сумма)."
            )
        
        real_data = {}
        real_row_errors = []
        for row_idx, row in enumerate(real_rows):
            try:
                text_val = row[1]
                sum_val = row[2]
            except IndexError:
                continue
                
            if isinstance(text_val, str) and 'Реализация' in text_val:
                id_clean = extract_id_real(text_val)
                excel_row = row_idx + 1
                if not id_clean:
                    real_row_errors.append(f"Реализации, строка {excel_row}: не удалось извлечь номер. Ожидается номер перед словом «от».")
                    continue
                try:
                    s = float(sum_val) if sum_val not in ('', None) else 0.0
                except (TypeError, ValueError):
                    real_row_errors.append(f"Реализации, строка {excel_row}: сумма «{sum_val}» в столбце C не является числом.")
                    continue
                real_data[id_clean] = {'Text': text_val, 'Сумма_Реал': s}

        if not real_data:
            real_row_errors.append("Реализации: не найдено документов с текстом «Реализация» и номером, который можно распознать.")

        data_errors = bp_row_errors + real_row_errors
        if data_errors:
            shown_errors = data_errors[:15]
            if len(data_errors) > len(shown_errors):
                shown_errors.append(f"…и еще {len(data_errors) - len(shown_errors)} ошибок.")
            messagebox.showwarning("Проверьте структуру и данные", "Сверка не выполнена. Исправьте найденные проблемы:\n\n• " + "\n• ".join(shown_errors))
            return

        # 3. Сверка данных
        all_ids = set(bp_data.keys()).union(set(real_data.keys()))
        report_data = []
        
        for doc_id in all_ids:
            b_info = bp_data.get(doc_id, {'Номер': 'Отсутствует в БП', 'Сумма_БП': 0.0})
            r_info = real_data.get(doc_id, {'Text': 'Отсутствует в выгрузке реализаций', 'Сумма_Реал': 0.0})
            
            diff = round(b_info['Сумма_БП'] - r_info['Сумма_Реал'], 2)
            if diff != 0:
                report_data.append([
                    b_info['Номер'],
                    b_info['Сумма_БП'],
                    r_info['Text'],
                    r_info['Сумма_Реал'],
                    diff
                ])

        # 4. Формирование Excel отчета
        wb_out = Workbook()
        ws_out = wb_out.active
        ws_out.title = "Расхождения"

        headers = ['Номер документа (из БП)', 'Сумма (БП)', 'Документ (из Окт. Реал)', 'Сумма (Окт. Реал)', 'Разница (БП - Реал)']
        ws_out.append(headers)

        for row_data in report_data:
            ws_out.append(row_data)

        # Стилизация шапки
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in ws_out[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Автоширина колонок и формат чисел
        for col in ws_out.columns:
            max_length = 0
            column_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
                if cell.row > 1 and cell.column in [2, 4, 5]:
                    cell.number_format = '#,##0.00'
            ws_out.column_dimensions[column_letter].width = min(max_length + 2, 60)

        wb_out.save(output_file)
        messagebox.showinfo("Готово", f"Сверка успешно завершена!\nНайдено расхождений: {len(report_data)}")

    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка при обработке данных:\n{e}")
    finally:
        btn_run.config(text="Запустить сверку", state=tk.NORMAL)

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

def show_instructions():
    window = tk.Toplevel(root)
    window.title("Инструкция по работе")
    window.geometry("640x520")
    window.minsize(500, 380)
    window.transient(root)

    container = tk.Frame(window, padx=16, pady=14)
    container.pack(fill=tk.BOTH, expand=True)
    scrollbar = tk.Scrollbar(container)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    instructions = tk.Text(container, wrap=tk.WORD, yscrollcommand=scrollbar.set,
                           padx=10, pady=10, font=("Arial", 10), relief=tk.FLAT)
    instructions.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=instructions.yview)

    instructions.insert("1.0", """ИНСТРУКЦИЯ ПО СВЕРКЕ ДОКУМЕНТОВ 1С

Назначение
Утилита сопоставляет документы из выгрузки БП и выгрузки реализаций по номеру, сравнивает суммы и сохраняет отчет с расхождениями.

Подготовка таблицы БП
• Допускаются файлы .xls и .xlsx; формат выбирается автоматически по расширению.
• Данные должны находиться на первом (активном) листе; строка 1 содержит заголовки.
• Укажите столбцы с заголовками «Номер» и «Сумма» именно в таком порядке

Подготовка таблицы реализаций
• Допускаются файлы .xls и .xlsx; формат выбирается автоматически по расширению.
• Нужен первый лист. В столбце B (вторая колонка) должен быть текст документа, а в столбце C (третья колонка) — сумма.
• Текст документа должен содержать «Реализация» и номер непосредственно перед «от», например: «Реализация ... 000123 от 01.01.2026».
• Заголовки не обязательны: данные читаются по позициям столбцов B и C. Строки без «Реализация» пропускаются.
• При сравнении ведущие нули и буквенный префикс номера реализации не учитываются.

Порядок работы
1. Нажмите «Выбрать» и укажите файл БП (.xls или .xlsx).
2. Выберите файл реализаций (.xls или .xlsx).
3. Введите путь и имя итогового файла .xlsx в поле «Сохранить итог» или нажмите «Выбрать» рядом с ним.
4. Нажмите «Запустить сверку».

Результат
Создается файл Excel с листом «Расхождения». В нем остаются документы с ненулевой разницей, включая документы, найденные только в одной из таблиц. Отсутствующая сумма считается равной нулю, разница округляется до двух знаков. Если в таблице несколько документов с одним номером, при обработке останется последний встретившийся.

Проверьте период выгрузок и корректность номеров перед запуском.""")
    instructions.config(state=tk.DISABLED)
    tk.Button(window, text="Закрыть", command=window.destroy).pack(pady=(0, 12))

root = tk.Tk()
root.title("Сверка документов 1С v0.3")
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

btn_run = tk.Button(frame, text="Запустить сверку", command=run_comparison, bg="#4F81BD", fg="white", font=("Arial", 10, "bold"))
btn_run.grid(row=3, column=0, columnspan=2, pady=15, ipadx=10, ipady=3)
tk.Button(frame, text="Инструкция", command=show_instructions).grid(row=3, column=2, pady=15, ipadx=8, ipady=3)

root.mainloop()
