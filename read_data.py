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

    def __init__(self, input_path):
        self.input_path = Path(input_path)
        self.data = []

    def is_product_name(self, text):
        """ Проверяет, является ли строка названием нового товара """
        return any(text.lower().startswith(name.lower()) for name in self.PRODUCT_NAMES)

    def parse_excel(self, file_path):
        print(f"Opening file: {file_path}")  # Отладка
        df = pd.read_excel(file_path, header=None, usecols=[0], engine='openpyxl')
        # print(df.head(5))  # Отладка

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
            for id, value in product.items():
                print(f"{id}: {value}")
            print()

    def process_all(self):
        files = list(self.input_path.glob("*.xlsx"))
        if not files:
            print(f"No Excel files found in {self.input_path}")
            return

        for file_path in files:
            print(f"Processing file: {file_path.name}")
            self.parse_excel(file_path)
        self.print_data()


# Использование
base_path = Path(__file__).parent
input_path = base_path / "test_data" / "input"

parser = StructuredExcelParser(input_path)
parser.process_all()
