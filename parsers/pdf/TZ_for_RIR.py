import pdfplumber
from pathlib import Path
import importlib.util

BASE_DIR = Path(__file__).resolve().parents[2]
SETTINGS_PATH = BASE_DIR / "settings.py"

spec = importlib.util.spec_from_file_location("settings", SETTINGS_PATH)
settings = importlib.util.module_from_spec(spec)
spec.loader.exec_module(settings)


class StructuredPdfParser:
    PRODUCT_NAMES = settings.product_names

    def __init__(self, file_path):
        self.file_path = Path(file_path)
        self.data = []

    def parse_pdf(self):
        print(f"Opening file: {self.file_path}")
        with pdfplumber.open(self.file_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                tables = page.extract_tables()
                if not tables:
                    continue

                for table_idx, table in enumerate(tables):
                    print(f"Processing table {table_idx + 1} on page {page_number}")
                    for row in table:
                        row_text = [str(cell).strip().replace("\n", " ") for cell in row if cell]
                        row_combined = " | ".join(row_text)

                        # Проверяем, содержится ли в строке название товара
                        if any(name.lower() in row_combined.lower() for name in self.PRODUCT_NAMES):
                            product_data = {"text": row_combined}
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
        self.parse_pdf()
        self.print_data()


file_path = Path("..", "..", "test_data", "input", "ТЗ для РИР.pdf")
parser = StructuredPdfParser(file_path)
parser.process()
