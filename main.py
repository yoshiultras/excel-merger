# -*- coding: utf-8 -*-

import pandas as pd
import locale
import tkinter as tk
from tkinter import filedialog, messagebox

locale.setlocale(locale.LC_ALL, 'ru_RU.UTF8')


class CSVMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Merger")

        self.file1 = ""
        self.file2 = ""

        # Кнопка для загрузки первого файла
        self.btn_load_file1 = tk.Button(root, text="Загрузить первый CSV файл", command=self.load_file1)
        self.btn_load_file1.pack(pady=10)

        # Кнопка для загрузки второго файла
        self.btn_load_file2 = tk.Button(root, text="Загрузить второй CSV файл", command=self.load_file2)
        self.btn_load_file2.pack(pady=10)

        # Поле ввода для общего поля
        self.entry_common_field = tk.Entry(root)
        self.entry_common_field.pack(pady=10)
        self.entry_common_field.insert(0, "Введите общее поле (например, 'id')")

        # Кнопка для выполнения объединения
        self.btn_merge = tk.Button(root, text="Объединить файлы", command=self.merge_csv)
        self.btn_merge.pack(pady=20)

    def load_file1(self):
        self.file1 = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.file1:
            print(f"Выбран первый файл: {self.file1}")

    def load_file2(self):
        self.file2 = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.file2:
            print(f"Выбран второй файл: {self.file2}")

    def merge_csv(self):
        common_field = self.entry_common_field.get().strip()
        if not self.file1 or not self.file2:
            messagebox.showerror("Ошибка", "Пожалуйста, загрузите оба файла.")
            return

        try:
            df1 = pd.read_csv(self.file1, encoding = "cp1252")
            df2 = pd.read_csv(self.file2, encoding = "cp1252")

            # Объединение по общему полю
            merged_df = pd.merge(df1, df2, on=common_field)

            # Сохранение результата в новый CSV файл
            merged_df.to_csv("merged_output.csv", index=False)
            messagebox.showinfo("Успех", "Файлы успешно объединены в 'merged_output.csv'.")
        except KeyError as e:
            messagebox.showerror("Ошибка", f"Столбец '{e}' не найден в одном из файлов.")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry('520x300')
    app = CSVMergerApp(root)
    root.mainloop()
