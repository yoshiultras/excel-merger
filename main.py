# -*- coding: utf-8 -*-

import pandas as pd
import locale
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

locale.setlocale(locale.LC_ALL, 'ru_RU.UTF8')


class CSVMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Merger")

        self.file1 = ""
        self.file2 = ""
        self.columns_file1 = []
        self.columns_file2 = []
        self.data_flow1 = None
        self.data_flow2 = None

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
        self.btn_merge = tk.Button(root, text="Объединить файлы", command=self.merge_csv)
        self.btn_merge.pack(pady=20)

    def load_file1(self):
        self.file1 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.file1:
            print(f"Выбран первый файл: {self.file1}")
            self.file1_label.config(text=self.file1)
            self.data_flow1 = pd.read_excel(self.file1)
            self.columns_file1 = self.data_flow1.columns.tolist()
            self.combobox_file1['values'] = self.columns_file1
            self.combobox_file1.set('')  # Сброс выбора

    def load_file2(self):
        self.file2 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.file2:
            print(f"Выбран второй файл: {self.file2}")
            self.file2_label.config(text=self.file2)
            self.data_flow2 = pd.read_excel(self.file2)
            self.columns_file2 = self.data_flow2.columns.tolist()
            self.combobox_file2['values'] = self.columns_file2
            self.combobox_file2.set('')  # Сброс выбора

    def merge_csv(self):
        common_field1 = self.combobox_file1.get().strip()
        common_field2 = self.combobox_file2.get().strip()

        if not self.file1 or not self.file2:
            messagebox.showerror("Ошибка", "Пожалуйста, загрузите оба файла.")
            return

        if not common_field1 or not common_field2:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите столбцы для объединения.")
            return

        try:
            # Объединение по выбранным столбцам
            merged_df = pd.merge(self.data_flow1, self.data_flow2, left_on=common_field1, right_on=common_field2)

            # Сохранение результата в новый CSV файл
            merged_df.to_excel("merged_output.xlsx", index=False)
            messagebox.showinfo("Успех", "Файлы успешно объединены в 'merged_output.xlsx'.")
        except KeyError as e:
            messagebox.showerror("Ошибка", f"Столбец '{e}' не найден в одном из файлов.")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry('520x420')
    app = CSVMergerApp(root)
    root.mainloop()
