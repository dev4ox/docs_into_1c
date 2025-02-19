import pandas as pd
from pathlib import Path
import importlib.util

# Загрузка настроек из settings.py
BASE_DIR = Path(__file__).resolve().parents[2]
SETTINGS_PATH = BASE_DIR / "settings.py"

spec = importlib.util.spec_from_file_location("settings", SETTINGS_PATH)
settings = importlib.util.module_from_spec(spec)
spec.loader.exec_module(settings)

class UnifiedExcelParser:
    PRODUCT_NAMES = settings.product_names  # список названий товаров из настроек

    def __init__(self, file_path):
        self.file_path = Path(file_path)
        self.data = []

    def is_product_name(self, text):
        """Проверяет, начинается ли строка с одного из имен товаров."""
        return any(text.lower().startswith(name.lower()) for name in self.PRODUCT_NAMES)

    def detect_engine(self):
        """Определяет движок для pd.read_excel по расширению файла."""
        ext = self.file_path.suffix.lower()
        if ext in ['.xlsx', '.xlsm']:
            return 'openpyxl'
        elif ext == '.xls':
            return 'xlrd'
        else:
            return 'openpyxl'

    def parse_excel(self):
        """Главный метод, определяющий стратегию парсинга в зависимости от структуры файла."""
        if not self.file_path.exists():
            print(f"File not found: {self.file_path}")
            return
        engine = self.detect_engine()
        print(f"Opening file: {self.file_path} with engine {engine}")
        df = pd.read_excel(self.file_path, header=None, engine=engine)

        # Если в файле только один столбец, применяем логику, аналогичную варианту 2
        if df.shape[1] == 1:
            self.parse_single_column(df)
        else:
            # Поиск столбца, содержащего "наименование" в первых 15 строках
            name_column = None
            for col in df.columns:
                for row_idx in range(min(15, len(df))):
                    cell_value = str(df.iloc[row_idx, col]).lower() if pd.notna(df.iloc[row_idx, col]) else ""
                    if "наименование" in cell_value:
                        name_column = col
                        print(f"Found 'Наименование' column at index {col}")
                        break
                if name_column is not None:
                    break

            if name_column is None:
                # Если столбец не найден, пробуем обработку как одностолбцового файла
                print("Столбец 'Наименование' не найден, пробуем обработку как одностолбцовый файл.")
                self.parse_single_column(df)
            else:
                # Если рядом с найденным столбцом есть дополнительные столбцы – используем мультиколоночный парсинг
                if name_column + 1 < df.shape[1]:
                    # Если имеется ещё и столбец после characteristics (name_column+2), то считаем, что данные о характеристиках распределены на два столбца (как в варианте 3)
                    extra_char = (name_column + 2 < df.shape[1])
                    self.parse_multi_column(df, name_column, extra_char)
                else:
                    self.parse_single_column(df)

    def parse_single_column(self, df):
        """
        Обработка файла с одним столбцом (логика варианта 2):
        Каждая строка проверяется на то, является ли она названием нового товара.
        """
        current_product = None
        product_line = ""
        for index, row in df.iterrows():
            cell_value = str(row[0]).strip()
            if self.is_product_name(cell_value):
                if current_product:
                    self.data.append(product_line)
                current_product = cell_value
                product_line = current_product
            else:
                product_line += f" {cell_value}"  # Добавляем текст в одну строку
        if current_product:
            self.data.append(product_line)

    def parse_multi_column(self, df, name_column, extra_char):
        """
        Обработка файлов с несколькими столбцами (логика вариантов 1, 3, 4 и 5).
        Если extra_char==True – предполагается, что характеристики распределены на два столбца.
        """
        current_product = None
        product_line = ""
        for index, row in df.iterrows():
            if pd.isna(row[name_column]):
                continue
            cell_value = str(row[name_column]).strip().replace("\t", " ")
            char_value = ""
            if extra_char:
                if name_column + 1 < df.shape[1]:
                    char_value += str(row[name_column + 1]).strip().replace("\t", " ").replace("\n", ";")
                if name_column + 2 < df.shape[1]:
                    char_value += " " + str(row[name_column + 2]).strip().replace("\t", " ").replace("\n", ";")
            else:
                if name_column + 1 < df.shape[1]:
                    char_value = str(row[name_column + 1]).strip().replace("\t", " ").replace("\n", ";")

            if any(name.lower() in cell_value.lower() for name in self.PRODUCT_NAMES):
                if current_product:
                    self.data.append(product_line)
                current_product = cell_value
                product_line = current_product
                if char_value:
                    product_line += f" {char_value}"
            else:
                product_line += f" {cell_value}"  # Добавляем дополнительную информацию в строку
        if current_product:
            self.data.append(product_line)

    def print_data(self):
        """Вывод полученных данных."""
        if not self.data:
            print("No data parsed!")
        else:
            for product in self.data:
                print(product)

    def process(self):
        print()
        """Основной метод для обработки файла."""
        if not self.file_path.exists():
            print(f"File not found: {self.file_path}")
            return
        print(f"Processing file: {self.file_path.name}")
        self.parse_excel()
        self.print_data()

# 1
file_path = Path("..", "..", "test_data", "input", "ТЗ для GPT.xlsx")
parser = UnifiedExcelParser(file_path)
parser.process()

# 2
file_path = Path("..", "..", "test_data", "input", "ТЗ для Рос Тюм.xlsm")
parser = UnifiedExcelParser(file_path)
parser.process()

# 3
file_path = Path("..", "..", "test_data", "input", "ТЗ для Ростов.xls")
parser = UnifiedExcelParser(file_path)
parser.process()

# 4
file_path = Path("..", "..", "test_data", "input", "ТЗ для Татэн.xls")
parser = UnifiedExcelParser(file_path)
parser.process()

# 5
file_path = Path("..", "..", "test_data", "input", "ТЗ для 213054.xlsx")
parser = UnifiedExcelParser(file_path)
parser.process()
