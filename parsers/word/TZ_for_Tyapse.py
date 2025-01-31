import docx
from pathlib import Path
import importlib.util

BASE_DIR = Path(__file__).resolve().parents[2]
SETTINGS_PATH = BASE_DIR / "settings.py"

spec = importlib.util.spec_from_file_location("settings", SETTINGS_PATH)
settings = importlib.util.module_from_spec(spec)
spec.loader.exec_module(settings)


class StructuredDocxParser:
    PRODUCT_NAMES = settings.product_names

    def __init__(self, file_path):
        self.file_path = Path(file_path)
        self.data = []

    def parse_doc(self):
        print(f"Opening file: {self.file_path}")
        doc = docx.Document(self.file_path)

        for table in doc.tables:
            for row in table.rows:
                row_text = [cell.text.strip().replace("\n", ";").replace("\t", " ") for cell in row.cells]
                row_combined = " | ".join(row_text)  # Разделяем ячейки для наглядности
                print(f"Row text: {row_combined}")  # Отладочный вывод

                # Проверяем, содержит ли строка товарное наименование
                if any(name.lower() in row_combined.lower() for name in self.PRODUCT_NAMES):
                    product_data = {"0": row_text}
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
        self.parse_doc()
        self.print_data()


# Использование
file_path = Path("..", "..", "test_data", "input", "ТЗ для Туапсе.docx")  # Конкретный файл DOC
parser = StructuredDocParser(file_path)
parser.process()
