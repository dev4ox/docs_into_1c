import pandas as pd
import docx
from openpyxl import load_workbook
import os
import re
from pathlib import Path


class StructuredExcelParser:
    PRODUCT_NAMES = [
        "Светильник", "Прожектор", "Лампа", "Осветительный прибор"
    ]

    def __init__(self, file_path):
        self.file_path = Path(file_path)
        self.data = []

    def is_product_name(self, text):
        """ Проверяет, является ли строка названием нового товара """
        return any(text.lower().startswith(name.lower()) for name in self.PRODUCT_NAMES)

    def parse_excel(self):
        print(f"Opening file: {self.file_path}")  # Отладочный вывод
        df = pd.read_excel(self.file_path, header=None, usecols=[0], engine='openpyxl')  # Читаем только 1-й столбец

        current_product = None
        product_data = {}

        for index, row in df.iterrows():
            cell_value = str(row[0]).strip()

            if self.is_product_name(cell_value):
                if current_product:
                    self.data.append(product_data)
                current_product = cell_value
                product_data = {"0": current_product}
            else:
                product_data[f"{len(product_data)}"] = cell_value

        if current_product:
            self.data.append(product_data)

    def print_data(self):
        if not self.data:
            print("No data parsed!")
        for product in self.data:
            print(product)

    def process(self):
        if not self.file_path.exists():
            print(f"File not found: {self.file_path}")
            return

        print(f"Processing file: {self.file_path.name}")
        self.parse_excel()
        self.print_data()


# Использование
file_path = Path("..", "test_data", "input", "ТЗ для GPT.xlsx")  # Конкретный файл Excel
parser = StructuredExcelParser(file_path)
parser.process()
