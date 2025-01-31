import pandas as pd
import docx
from openpyxl import load_workbook
import os
import re
from pathlib import Path


class StructuredXlsParser:
    PRODUCT_NAMES = ["Светильник", "Прожектор", "Лампа", "Осветительный прибор"]

    def __init__(self, file_path):
        self.file_path = Path(file_path)
        self.data = []

    def parse_xls(self):
        print(f"Opening file: {self.file_path}")
        df = pd.read_excel(self.file_path, header=None, engine='xlrd')  # Загружаем файл XLS

        # Найти столбец с "Наименование"
        name_column = None
        characteristics_column = None
        for col in df.columns:
            for row_idx in range(min(15, len(df))):  # Проверяем первые 15 строк
                cell_value = str(df.iloc[row_idx, col]).lower() if pd.notna(df.iloc[row_idx, col]) else ""
                if "наименование" in cell_value:
                    name_column = col
                    characteristics_column = col + 1  # Предполагаем, что характеристики справа
                    print(f"Found 'Наименование' column at index {col}")
                    break
            if name_column is not None:
                break

        if name_column is None or characteristics_column is None:
            print("Column 'Наименование' or characteristics column not found!")
            return

        current_product = None
        product_data = {}

        # Перебираем строки в найденном столбце
        for index, row in df.iterrows():
            if pd.isna(row[name_column]):
                continue

            cell_value = str(row[name_column]).strip().replace("\t", " ")
            char_value = str(row[name_column + 1]).strip().replace("\t", " ").replace("\n", ";")

            if any(name.lower() in cell_value.lower() for name in self.PRODUCT_NAMES):
                if current_product:
                    self.data.append(product_data)
                current_product = cell_value
                current_property = char_value
                product_data = {"0": current_product, "1": current_property}

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
        self.parse_xls()
        self.print_data()


# Использование
file_path = Path("..", "test_data", "input", "ТЗ для Татэн.xls")  # Конкретный файл XLS
parser = StructuredXlsParser(file_path)
parser.process()
