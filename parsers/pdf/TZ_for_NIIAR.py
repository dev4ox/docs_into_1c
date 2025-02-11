import pdfplumber
import re
from pathlib import Path
import importlib.util
import collections

BASE_DIR = Path(__file__).resolve().parents[2]
SETTINGS_PATH = BASE_DIR / "settings.py"

spec = importlib.util.spec_from_file_location("settings", SETTINGS_PATH)
settings = importlib.util.module_from_spec(spec)
spec.loader.exec_module(settings)


class StructuredPdfParser:
    PRODUCT_NAMES = settings.product_names
    EXCLUDE_WORDS = ["шт.", "шт", "штук"]

    def __init__(self, file_path):
        self.file_path = Path(file_path)
        self.data = []
        self.header_candidates = []  # список токенов из первых пяти найденных заголовков товара
        self.header_mask = None      # регулярное выражение для определения начала нового товара

    def determine_common_pattern(self):
        """
        Определяет общий паттерн из накопленных header_candidates.
        Если минимум 3 из 5 кандидатов соответствуют схеме "число + точка", возвращается паттерн r'^\s*\d+\s*\.'.
        Иначе выбирается самый часто встречающийся токен и формируется маска, проверяющая его наличие в начале строки.
        """
        if not self.header_candidates:
            return None

        # Проверяем, сколько кандидатов соответствуют схеме "число + точка"
        digit_dot_pattern = re.compile(r'^\d+\.$')
        count_digit_dot = sum(1 for token in self.header_candidates if digit_dot_pattern.match(token))
        if count_digit_dot >= 3:
            return re.compile(r'^\s*\d+\s*\.')

        # Иначе выбираем наиболее часто встречающийся токен
        counter = collections.Counter(self.header_candidates)
        most_common_token, freq = counter.most_common(1)[0]
        if most_common_token:
            escaped = re.escape(most_common_token)
            return re.compile(r'^\s*' + escaped)
        return None

    def update_header_mask(self):
        """
        Если накоплено не менее 5 кандидатов и паттерн ещё не установлен, вычисляем его.
        """
        if len(self.header_candidates) >= 5 and not self.header_mask:
            self.header_mask = self.determine_common_pattern()

    def is_new_header(self, row_combined):
        """
        Если паттерн определён, строка считается началом нового товара, если совпадает с паттерном.
        Если паттерн не установлен, в качестве критерия используется наличие наименования товара.
        """
        if self.header_mask:
            return bool(self.header_mask.match(row_combined))
        return any(name.lower() in row_combined.lower() for name in self.PRODUCT_NAMES)

    def parse_pdf(self):
        print(f"Opening file: {self.file_path}")
        current_record = None
        with pdfplumber.open(self.file_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                tables = page.extract_tables()
                if not tables:
                    continue

                for table_idx, table in enumerate(tables):
                    print(f"Processing table {table_idx + 1} on page {page_number}")
                    for row in table:
                        row_text = [str(cell).strip().replace("\n", " ") for cell in row if cell]
                        if not row_text:
                            continue
                        row_combined = " | ".join(row_text)
                        row_lower = row_combined.lower()

                        # Пропускаем строки, содержащие исключающие слова
                        if any(ex_word in row_lower for ex_word in self.EXCLUDE_WORDS):
                            continue

                        # Если строка содержит наименование товара, считаем её кандидатом на начало записи
                        if any(name.lower() in row_lower for name in self.PRODUCT_NAMES):
                            tokens = row_combined.split()
                            first_token = tokens[0] if tokens else ""
                            if first_token and first_token not in self.header_candidates:
                                self.header_candidates.append(first_token)
                            # Обновляем маску, если набрано достаточно кандидатов
                            self.update_header_mask()

                            # Если уже есть накопленная запись, сохраняем её
                            if current_record:
                                self.data.append({"0": current_record})
                            current_record = row_combined
                        else:
                            # Если паттерн установлен и строка соответствует началу нового товара,
                            # то считаем её новым заголовком
                            if self.header_mask and self.is_new_header(row_combined):
                                if current_record:
                                    self.data.append({"0": current_record})
                                current_record = row_combined
                            else:
                                # Иначе, строка считается продолжением предыдущего товара
                                if current_record:
                                    if current_record.endswith('-'):
                                        current_record = current_record.rstrip('-') + row_combined.lstrip()
                                    else:
                                        current_record += " " + row_combined
        if current_record:
            self.data.append({"0": current_record})

    def print_data(self):
        if not self.data:
            print("No data parsed!")
        else:
            for product in self.data:
                print(product)

    def process(self):
        if not self.file_path.exists():
            print(f"File not found: {self.file_path}")
            return

        print(f"Processing file: {self.file_path.name}")
        self.parse_pdf()
        self.print_data()


file_path = Path("..", "..", "test_data", "input", "ТЗ для НИИАР поз №158.pdf")
parser = StructuredPdfParser(file_path)
parser.process()
