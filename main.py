import _tkinter

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook


# Функция для обработки значений
def process_data(series):
    return series.apply(
        lambda x: x.strip().lower() if isinstance(x, str) else str(x).strip().lower()
    ).apply(
        lambda x: [
            num.strip().replace('8', '+7', 1).replace(' ', '').replace('(', '').replace(')', '').replace('-', '') if num.strip().startswith('8') else
            "+7" + num.strip()[1:].replace(' ', '').replace('(', '').replace(')', '').replace('-', '') if num.strip().startswith('7') else num.strip().replace(' ', '').replace('(', '').replace(')', '').replace('-', '')
            for num in x.split(';') if len(num.strip()) > 5 and '_' not in num
        ] if isinstance(x, str) else []
    )



def merge_excel(df1, df2, common_fields):
    for common_field1, common_field2 in common_fields:
        if common_field1 != 'ФИО':
            df1[common_field1] = process_data(df1[common_field1])
        if common_field2 != 'ФИО':
            df2[common_field2] = process_data(df2[common_field2])

        temp_df1 = df1.explode(common_field1)
        temp_df2 = df2.explode(common_field2)

        # Обработка пустых данных
        temp_df1[common_field1] = temp_df1[common_field1].apply(lambda x: ' ' if x == [] or x == 'nan' or str(x) is None else x)
        temp_df1.fillna(' ', inplace=True)

        pd.set_option('display.max_columns', 1000)  # or 1000
        pd.set_option('display.max_rows', 1000)  # or 1000
        pd.set_option('display.max_colwidth', 199)  # or 199
        print(temp_df1)
        print(temp_df2)
        # Объединение данных
        merged_df = pd.merge(temp_df1, temp_df2, left_on=common_field1, right_on=common_field2)

    return merged_df


class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Merger")

        self.file1 = ""
        self.file2 = ""
        self.columns_file1 = []
        self.columns_file2 = []
        self.common_fields = []

        # Кнопка для загрузки первого файла
        self.btn_load_file1 = tk.Button(root, text="Загрузить первый файл", command=self.load_file1)
        self.btn_load_file1.pack(pady=10)
        self.file1_label = tk.Label()
        self.file1_label.pack(pady=10)

        # Кнопка для загрузки второго файла
        self.btn_load_file2 = tk.Button(root, text="Загрузить второй файл", command=self.load_file2)
        self.btn_load_file2.pack(pady=10)
        self.file2_label = tk.Label()
        self.file2_label.pack(pady=10)

        # Фрейм для пар столбцов
        self.pair_frame = tk.Frame(root)
        self.pair_frame.pack(pady=10)

        self.add_pair_button = tk.Button(root, text="Добавить пару столбцов для объединения", command=self.add_column_pair)
        self.add_pair_button.pack(pady=5)

        self.btn_merge = tk.Button(root, text="Объединить файлы", command=self.merge)
        self.btn_merge.pack(pady=20)

    def add_column_pair(self):
        pair_frame = tk.Frame(self.pair_frame)
        pair_frame.pack(pady=5)

        combobox_file1 = ttk.Combobox(pair_frame, state="readonly")
        combobox_file1['values'] = self.columns_file1
        combobox_file1.pack(side=tk.LEFT)

        combobox_file2 = ttk.Combobox(pair_frame, state="readonly")
        combobox_file2['values'] = self.columns_file2
        combobox_file2.pack(side=tk.LEFT)

        remove_button = tk.Button(pair_frame, text="Удалить",
                                  command=lambda: self.remove_pair(pair_frame, combobox_file1, combobox_file2))
        remove_button.pack(side=tk.LEFT)

        # Сохраняем ссылки на комбобоксы
        self.common_fields.append((combobox_file1, combobox_file2))

    def remove_pair(self, frame, combobox_file1, combobox_file2):
        # Получаем значения из комбобоксов перед их удалением
        value1 = combobox_file1.get()
        value2 = combobox_file2.get()

        # Удаляем фрейм
        frame.destroy()

        # Обновляем список пар, исключая удаляемую
        self.common_fields = [(c1, c2) for c1, c2 in self.common_fields if (c1.get() != value1 or c2.get() != value2)]

    def load_file1(self):
        self.file1 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.file1:
            self.file1_label.config(text=self.file1)
            df1 = pd.read_excel(self.file1)
            self.columns_file1 = df1.columns.tolist()
            for pair in self.common_fields:
                pair[0]['values'] = self.columns_file1
            self.common_fields = [(pair[0], pair[1]) for pair in self.common_fields]

    def load_file2(self):
        self.file2 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.file2:
            self.file2_label.config(text=self.file2)
            df2 = pd.read_excel(self.file2)
            self.columns_file2 = df2.columns.tolist()
            for pair in self.common_fields:
                pair[1]['values'] = self.columns_file2
            self.common_fields = [(pair[0], pair[1]) for pair in self.common_fields]

    def merge(self):
        if not self.file1 or not self.file2:
            messagebox.showerror("Ошибка", "Пожалуйста, загрузите оба файла.")
            return

        if not self.common_fields:
            messagebox.showerror("Ошибка", "Пожалуйста, добавьте хотя бы одну пару столбцов для объединения.")
            return

        try:
            df1 = pd.read_excel(self.file1)
            df2 = pd.read_excel(self.file2)

            merged_df = pd.DataFrame()

            for combobox_file1, combobox_file2 in self.common_fields:
                try:
                    common_field1 = combobox_file1.get().strip()
                    common_field2 = combobox_file2.get().strip()
                    print(common_field1)
                    print(common_field2)
                    if common_field1 and common_field2:
                        temp_merged = merge_excel(df1, df2, [(common_field1, common_field2)])
                        merged_df = pd.concat([merged_df, temp_merged], ignore_index=True)

                        # Преобразование списков в строки
                        for col in merged_df.columns:
                            merged_df[col] = merged_df[col].apply(lambda x: ', '.join(x) if isinstance(x, list) else x)

                        # merged_df = merged_df.drop_duplicates(subset=[common_field1])  # Удаление дубликатов
                except _tkinter.TclError as e:
                    continue
            f1_rows = len(df1.index)
            f2_rows = len(df2.index)
            merged_rows = len(merged_df.index)
            # Сохранение результата в новый файл
            output_file = "merged_output.xlsx"
            merged_df.to_excel(output_file, index=False)

            # Открытие файла для форматирования
            wb = load_workbook(output_file)
            ws = wb.active

            # Автоматическая подгонка ширины столбцов
            for column in ws.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 5)  # Добавляем немного отступа
                ws.column_dimensions[column[0].column_letter].width = adjusted_width

            # Сохранение изменений
            wb.save(output_file)
            wb.close()

            messagebox.showinfo("Успех", "Файлы успешно объединены в 'merged_output.xlsx'.\n"
                                         f"Получено данных: {merged_rows}\n"
                                         f"Процент вхождения данных из первого файла: {merged_rows/f1_rows:.1%}\n"
                                         f"Процент вхождения данных из второго файла: {merged_rows/f2_rows:.1%}")

        except KeyError as e:
            messagebox.showerror("Ошибка", f"Столбец '{e}' не найден в одном из файлов.")
        except PermissionError as e:
            messagebox.showerror("Ошибка", f"К файлу '{e.filename}' запрещен доступ. Закройте файл или запустите программу с правами администратора.")
        # except Exception as e:
        #     messagebox.showerror("Ошибка", e)
        #     print(e)

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry('400x700')
    app = ExcelMergerApp(root)
    root.mainloop()
