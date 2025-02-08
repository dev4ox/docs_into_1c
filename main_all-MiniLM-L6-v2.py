import os
import pandas as pd
from pathlib import Path
import importlib.util
from sentence_transformers import SentenceTransformer, util
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Загрузка настроек из settings.py
BASE_DIR = Path(__file__).resolve().parent
SETTINGS_PATH = BASE_DIR / "settings.py"

spec = importlib.util.spec_from_file_location("settings", SETTINGS_PATH)
settings = importlib.util.module_from_spec(spec)
spec.loader.exec_module(settings)


# -------------- Класс для объединённого парсинга Excel ---------------
class UnifiedExcelParser:
    PRODUCT_NAMES = settings.product_names  # список ключевых названий товаров

    def __init__(self, file_path):
        self.file_path = Path(file_path)
        self.data = []  # сюда будем записывать разобранные товары

    def is_product_name(self, text):
        """Проверяет, начинается ли строка с одного из названий товара."""
        return any(text.lower().startswith(name.lower()) for name in self.PRODUCT_NAMES)

    def detect_engine(self):
        """Определяет движок pd.read_excel по расширению файла."""
        ext = self.file_path.suffix.lower()
        if ext in ['.xlsx', '.xlsm']:
            return 'openpyxl'
        elif ext == '.xls':
            return 'xlrd'
        else:
            return 'openpyxl'

    def parse_excel(self):
        if not self.file_path.exists():
            print(f"File not found: {self.file_path}")
            return
        engine = self.detect_engine()
        print(f"Opening file: {self.file_path} with engine {engine}")
        df = pd.read_excel(self.file_path, header=None, engine=engine)
        # Если в файле только один столбец – используем одностолбцовый разбор
        if df.shape[1] == 1:
            self.parse_single_column(df)
        else:
            # Ищем столбец с ключевым словом "наименование" в первых 15 строках
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
                print("Столбец 'Наименование' не найден, пробуем обработку как одностолбцовый файл.")
                self.parse_single_column(df)
            else:
                # Если рядом с найденным столбцом есть ещё ячейки, определяем схему разбора
                extra_char = (name_column + 2 < df.shape[1])
                self.parse_multi_column(df, name_column, extra_char)

    def parse_single_column(self, df):
        """Разбор для файла с одним столбцом (логика варианта 2)."""
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

    def parse_multi_column(self, df, name_column, extra_char):
        """
        Разбор для файлов с несколькими столбцами (варианты 1, 3, 4, 5).
        Если extra_char==True, предполагается, что данные о характеристиках распределены по двум столбцам.
        """
        current_product = None
        product_data = {}
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
                    self.data.append(product_data)
                current_product = cell_value
                product_data = {"0": current_product}
                if char_value:
                    product_data["1"] = char_value
            else:
                product_data[f"{len(product_data)}"] = cell_value
        if current_product:
            self.data.append(product_data)

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
        self.parse_excel()
        self.print_data()


# -------------- Класс для заполнения формы ---------------
class FormFiller:
    def __init__(self):
        self.model = SentenceTransformer('models/all-MiniLM-L6-v2')
        # Целевые столбцы (в нужном порядке)
        self.columns = [
            "Номенклатура",
            "Мощность, Вт",
            "Св. поток, Лм",
            "IP",
            "Габариты",
            "Длина, мм",
            "Ширина, мм",
            "Высота, мм",
            "Рассеиватель",
            "Цвет. температура, К",
            "Вес, кг",
            "Напряжение, В",
            "Температура эксплуатации",
            "Срок службы (работы) светильника",
            "Тип КСС",
            "Род тока",
            "Гарантия",
            "Индекс цветопередачи (CRI, Ra)",
            "Цвет корпуса",
            "Коэффициент пульсаций",
            "Коэффициент мощности (Pf)",
            "Класс взрывозащиты (Ex)",
            "Класс пожароопасности",
            "Класс защиты от поражения электрическим током",
            "Материал корпуса",
            "Тип",
            "Прочее"
        ]
        # Словарь с синонимами для некоторых столбцов
        self.synonyms = {
            "Номенклатура": ["серия", "наименование", "товар"],
            "Мощность, Вт": ["мощность", "энергопотребление", "Вт", "W"],
            "Св. поток, Лм": ["световой поток", "Лм", "общий световой поток"],
            "IP": ["степень защиты", "IP", "защита от пыли и влаги"],
            "Габариты": ["размеры", "габариты"],
            "Цвет. температура, К": ["цветовая температура", "К", "температура света"],
            "Вес, кг": ["вес", "масса", "общий вес", "кг"],
            "Материал корпуса": ["материал корпуса", "материал изделия"],
            "Тип": ["тип крепления", "монтаж"],
            "Гарантия": ["гарантия", "гарантийный срок"],
        }
        # Предвычисляем эмбеддинги для всех целевых столбцов (fallback)
        self.column_embeddings = self.model.encode(self.columns, convert_to_tensor=True)

    def fill_form(self, product: dict) -> dict:
        """
        Заполняет форму по товару.
        Сначала пытается сопоставить каждое значение по словарю синонимов.
        Если сопоставление не найдено – используется семантическое сравнение через модель.
        """
        # Инициализация формы с базовыми значениями
        form = {col: "не указано" for col in self.columns}
        # Прямое присвоение для "Номенклатуры"
        form["Номенклатура"] = product.get("0", "не указано").replace(" или эквивалент.", "").strip()

        unassigned_texts = []
        # Перебор остальных атрибутов (ключи "1", "2", ...)
        for key in sorted(product.keys(), key=lambda x: int(x) if x.isdigit() else x):
            if key == "0":
                continue
            text = product[key]
            assigned = False

            # 1. Проверка по словарю синонимов (учтено, что регистр не важен)
            for target, syn_list in self.synonyms.items():
                for syn in syn_list:
                    if syn.lower() in text.lower():
                        if form[target] != "не указано":
                            form[target] += "; " + text
                        else:
                            form[target] = text
                        assigned = True
                        break
                if assigned:
                    break

            # 2. Если сопоставление по синонимам не удалось, используем семантическое сравнение
            if not assigned:
                text_embedding = self.model.encode(text, convert_to_tensor=True)
                # Рассматриваем кандидатов – все столбцы, кроме "Номенклатуры" и "Прочее"
                candidate_columns = [col for col in self.columns if col not in ["Номенклатура", "Прочее"]]
                candidate_indices = [self.columns.index(col) for col in candidate_columns]
                candidate_embeds = self.column_embeddings[candidate_indices]
                cos_scores = util.cos_sim(text_embedding, candidate_embeds)[0]
                best_idx = int(cos_scores.argmax())
                best_score = float(cos_scores[best_idx])
                matched_column = candidate_columns[best_idx]
                if best_score >= 0.3:
                    if form[matched_column] != "не указано":
                        form[matched_column] += "; " + text
                    else:
                        form[matched_column] = text
                else:
                    unassigned_texts.append(text)

        if unassigned_texts:
            form["Прочее"] = " ".join(unassigned_texts)
        # Специальная обработка для "Габаритов": объединяем, если заданы "Длина, мм" и "Ширина, мм"
        if form["Длина, мм"] != "не указано" and form["Ширина, мм"] != "не указано":
            deviation = ""
            if "отклонение" in form["Прочее"]:
                deviation = form["Прочее"]
            form["Габариты"] = f"{form['Длина, мм']}x{form['Ширина, мм']} мм {deviation}".strip()
        return form




# -------------- Парсинг файла и заполнение формы ---------------
if __name__ == "__main__":
    # Путь к входящему файлу (пример: "ТЗ для GPT.xlsx")
    input_file_path = Path("test_data", "input", "ТЗ для GPT.xlsx")
    parser = UnifiedExcelParser(input_file_path)
    parser.process()

    filler = FormFiller()
    filled_forms = []
    for product in parser.data:
        form = filler.fill_form(product)
        filled_forms.append(form)

    df_form = pd.DataFrame(filled_forms, columns=filler.columns)
    print("\nЗаполненная форма:")
    print(df_form.to_string(index=False))

    # Запись в output.xlsx: если файл существует, добавляем новые строки после последней заполненной строки.
    output_file = "output.xlsx"
    append_df_to_excel(output_file, df_form, sheet_name="Sheet1")
    print(f"\nДанные успешно добавлены в файл {output_file}.")
