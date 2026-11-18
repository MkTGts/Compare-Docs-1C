import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill

def extract_id_real(text):
    """Извлекает числовой номер документа из строки 1С, игнорируя нули и префиксы."""
    match = re.search(r'\s+([А-Яа-яA-Za-z-]*0*(\d+))\s+от\s+', str(text))
    if match:
        return match.group(2).lstrip('0')
    return None

def run_comparison():
    file_bp = entry_bp.get()
    file_real = entry_real.get()
    output_file = entry_out.get()

    if not file_bp or not file_real or not output_file:
        messagebox.showwarning("Внимание", "Пожалуйста, заполните пути ко всем файлам!")
        return

    # Блокируем кнопку на время работы скрипта
    btn_run.config(text="Обработка...", state=tk.DISABLED)
    root.update()

    try:
        # Чтение файла БП
        df1 = pd.read_excel(file_bp)
        df1['ID'] = df1['Номер'].astype(str).str.extract(r'(\d+)')[0]
        df1['ID_clean'] = df1['ID'].str.lstrip('0')
        df1['Сумма_БП'] = pd.to_numeric(df1['Сумма'], errors='coerce').fillna(0)

        # Чтение файла Реализаций
        df2_raw = pd.read_excel(file_real)
        df2 = df2_raw[['Продажи', 'Unnamed: 2']].copy()
        df2.columns = ['Text', 'Сумма_Реал']
        df2 = df2.dropna(subset=['Text'])
        df2 = df2[df2['Text'].astype(str).str.contains('Реализация', na=False)]
        
        df2['ID_clean'] = df2['Text'].apply(extract_id_real)
        df2['Сумма_Реал'] = pd.to_numeric(df2['Сумма_Реал'], errors='coerce').fillna(0)

        # Сверка
        merged = pd.merge(df1[['Номер', 'ID_clean', 'Сумма_БП']], 
                          df2[['Text', 'ID_clean', 'Сумма_Реал']], 
                          on='ID_clean', how='outer')

        merged['Разница'] = merged['Сумма_БП'].fillna(0) - merged['Сумма_Реал'].fillna(0)
        merged['Разница'] = merged['Разница'].round(2)
        discrepancies = merged[merged['Разница'] != 0].copy()

        report_df = discrepancies[['Номер', 'Сумма_БП', 'Text', 'Сумма_Реал', 'Разница']].copy()
        report_df.columns = ['Номер документа (из БП)', 'Сумма (БП)', 'Документ (из Окт. Реал)', 'Сумма (Окт. Реал)', 'Разница (БП - Реал)']

        report_df['Номер документа (из БП)'] = report_df['Номер документа (из БП)'].fillna('Отсутствует в БП')
        report_df['Документ (из Окт. Реал)'] = report_df['Документ (из Окт. Реал)'].fillna('Отсутствует в выгрузке реализаций')

        # Формирование Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "Расхождения"

        for r in dataframe_to_rows(report_df, index=False, header=True):
            ws.append(r)

        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for col in ws.columns:
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
            ws.column_dimensions[column_letter].width = min(max_length + 2, 60)

        wb.save(output_file)
        messagebox.showinfo("Готово", f"Сверка успешно завершена!\nНайдено расхождений: {len(report_df)}")

    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка при обработке данных:\n{e}")
    finally:
        # Возвращаем кнопку в исходное состояние
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

# Настройка главного окна
root = tk.Tk()
root.title("Сверка документов 1С")
root.geometry("540x200")
root.resizable(False, False)

frame = tk.Frame(root, padx=15, pady=15)
frame.pack(fill=tk.BOTH, expand=True)

# Элементы интерфейса
tk.Label(frame, text="Файл БП с суммой:").grid(row=0, column=0, sticky="w", pady=5)
entry_bp = tk.Entry(frame, width=40)
entry_bp.grid(row=0, column=1, padx=5, pady=5)
tk.Button(frame, text="Выбрать", command=lambda: select_file(entry_bp)).grid(row=0, column=2, pady=5)

tk.Label(frame, text="Файл Окт. Реал:").grid(row=1, column=0, sticky="w", pady=5)
entry_real = tk.Entry(frame, width=40)
entry_real.grid(row=1, column=1, padx=5, pady=5)
tk.Button(frame, text="Выбрать", command=lambda: select_file(entry_real)).grid(row=1, column=2, pady=5)

tk.Label(frame, text="Сохранить итог:").grid(row=2, column=0, sticky="w", pady=5)
entry_out = tk.Entry(frame, width=40)
entry_out.insert(0, "Расхождения_сверки.xlsx")
entry_out.grid(row=2, column=1, padx=5, pady=5)
tk.Button(frame, text="Выбрать", command=lambda: select_save_file(entry_out)).grid(row=2, column=2, pady=5)

btn_run = tk.Button(frame, text="Запустить сверку", command=run_comparison, bg="#4F81BD", fg="white", font=("Arial", 10, "bold"))
btn_run.grid(row=3, column=0, columnspan=3, pady=15, ipadx=10, ipady=3)

# Запуск цикла событий
root.mainloop()