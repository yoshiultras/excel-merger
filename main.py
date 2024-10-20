import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook


# Функция для обработки значений
def process_data(series):
    return series.apply(
        lambda x: x.strip().lower() if isinstance(x, str) else str(x).strip().lower()).apply(
        lambda x: [
            num.strip().replace('8', '+7', 1) if num.strip().startswith('8') else num.strip()
            for num in x.split(';') if num.strip() != '+7' and '_' not in num
        ] if isinstance(x, str) else []
    )



def merge_excel(df1, df2, common_field1, common_field2):
    if common_field1 != 'ФИО':
        df1[common_field1] = process_data(df1[common_field1])
    if common_field2 != 'ФИО':
        df2[common_field2] = process_data(df2[common_field2])

    df1 = df1.explode(common_field1)
    df2 = df2.explode(common_field2)

    # Обработка пустых данных
    # Следует переделать
    df1[common_field1] = df1[common_field1].apply(lambda x: ' ' if x == [] or x == 'nan' or str(x) is None else x)
    df1.fillna(' ', inplace=True)

    merged_df = pd.merge(df1, df2, left_on=common_field1, right_on=common_field2)
    return merged_df


class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Merger")

        self.file1 = ""
        self.file2 = ""
        self.columns_file1 = []
        self.columns_file2 = []

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

        # Селектор для выбора столбца из первого файла
        self.label_file1 = tk.Label(root, text="Выберите столбец объединения из первого файла:")
        self.label_file1.pack(pady=5)
        self.combobox_file1 = ttk.Combobox(root, state="readonly")
        self.combobox_file1.pack(pady=5)

        # Селектор для выбора столбца из второго файла
        self.label_file2 = tk.Label(root, text="Выберите столбец объединения из второго файла:")
        self.label_file2.pack(pady=5)
        self.combobox_file2 = ttk.Combobox(root, state="readonly")
        self.combobox_file2.pack(pady=5)

        # Кнопка для выполнения объединения
        self.btn_merge = tk.Button(root, text="Объединить файлы", command=self.merge)
        self.btn_merge.pack(pady=20)

    def load_file1(self):
        self.file1 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.file1:
            print(f"Выбран первый файл: {self.file1}")
            self.file1_label.config(text=self.file1)
            df1 = pd.read_excel(self.file1)
            self.columns_file1 = df1.columns.tolist()
            self.combobox_file1['values'] = self.columns_file1
            self.combobox_file1.set('')  # Сброс выбора

    def load_file2(self):
        self.file2 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.file2:
            print(f"Выбран второй файл: {self.file2}")
            self.file2_label.config(text=self.file2)
            df2 = pd.read_excel(self.file2)
            self.columns_file2 = df2.columns.tolist()
            self.combobox_file2['values'] = self.columns_file2
            self.combobox_file2.set('')  # Сброс выбора

    def merge(self):
        common_field1 = self.combobox_file1.get().strip()
        common_field2 = self.combobox_file2.get().strip()

        if not self.file1 or not self.file2:
            messagebox.showerror("Ошибка", "Пожалуйста, загрузите оба файла.")
            return

        if not common_field1 or not common_field2:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите столбцы для объединения.")
            return

        try:
            df1 = pd.read_excel(self.file1)
            df2 = pd.read_excel(self.file2)
            merged_df = merge_excel(df1, df2, common_field1, common_field2)
            f1_rows = len(df1.index)
            f2_rows = len(df2.index)
            merged_rows = len(merged_df.index)
            # Код отладки
            # pd.set_option('display.max_rows', None)
            # pd.set_option('display.max_columns', None)
            # pd.set_option('display.width', None)
            # pd.set_option('display.max_colwidth', None)
            # print(merged_df)

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
            messagebox.showerror("Ошибка", f"К файлу '{e.filename}' запрещен доступ. Закройте файл или запустите программу от имени администратора.")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry('620x420')
    app = ExcelMergerApp(root)
    root.mainloop()
