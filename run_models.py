import os
import json
from llama_cpp import Llama
import pandas as pd
from pathlib import Path
import importlib.util
import re
import pdfplumber
import collections


BASE_DIR = Path(__file__).resolve().parent
SETTINGS_PATH = BASE_DIR / "settings.py"
spec = importlib.util.spec_from_file_location("settings", SETTINGS_PATH)
settings = importlib.util.module_from_spec(spec)
spec.loader.exec_module(settings)


def extract_gemma_2_2b_it_IQ3_M(text, initial_prompt, final_columns):
    llm = Llama(
        model_path="../../.lmstudio/models/lmstudio-community/gemma-2-2b-it-GGUF/gemma-2-2b-it-IQ3_M.gguf",
        n_ctx=4096,
        n_gpu_layers=-1,
        verbose=False,
    )
    prompt = initial_prompt + "\n\nText:\n" + text + "\n\nJSON:"
    output = llm(prompt=prompt, max_tokens=512, temperature=0.0)
    result_text = output["choices"][0]["text"].strip()
    # Извлекаем только JSON-объект с помощью регулярного выражения
    match = re.search(r'\{.*\}', result_text, re.DOTALL)
    if match:
        result_text = match.group(0)
    try:
        data = json.loads(result_text)
    except Exception as e:
        print("Error parsing JSON:", e)
        print("Raw output:", result_text)
        data = {col: "не указано" for col in final_columns}
    return data


def extract_gemma_2_2b_it_Q6_K(text, initial_prompt, final_columns):
    llm = Llama(
        model_path="../../.lmstudio/models/lmstudio-community/gemma-2-2b-it-GGUF/gemma-2-2b-it-Q6_K.gguf",
        n_ctx=8192,
        n_gpu_layers=-1,
        verbose=False
    )
    prompt = initial_prompt + "\n\nText:\n" + text + "\n\nJSON:"
    output = llm(prompt=prompt, max_tokens=512, temperature=0.0)
    result_text = output["choices"][0]["text"].strip()
    # Извлекаем только JSON-объект с помощью регулярного выражения
    match = re.search(r'\{.*\}', result_text, re.DOTALL)
    if match:
        result_text = match.group(0)
    try:
        data = json.loads(result_text)
    except Exception as e:
        print("Error parsing JSON:", e)
        print("Raw output:", result_text)
        data = {col: "не указано" for col in final_columns}
    return data


def extract_with_mistral(text, initial_prompt):
    llm = Llama(
        model_path="../../.lmstudio/models/lmstudio-community/Mistral-7B-Instruct-v0.3-GGUF/Mistral-7B-Instruct-v0.3-Q4_K_M.gguf",
        n_ctx=3072
    )
    prompt = initial_prompt + "\n\nДанные для обработки:\n" + text
    output = llm(prompt=prompt, max_tokens=3072, temperature=0.1)
    result_text = output["choices"][0]["text"].strip()
    try:
        data = json.loads(result_text)
    except Exception as e:
        print("Error parsing JSON:", e)
        print("Raw output:", result_text)
        data = {}
    return data


def append_df_to_excel(filename, df, sheet_name='Sheet1'):
    # Используем ExcelWriter в режиме добавления (append)
    if os.path.exists(filename):
        with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            writer.sheets = {ws.title: ws for ws in writer.book.worksheets}
            startrow = writer.sheets[sheet_name].max_row
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=startrow)
    else:
        df.to_excel(filename, index=False, sheet_name=sheet_name)


class UnifiedExcelParser:
    PRODUCT_NAMES = settings.product_names

    def __init__(self, file_path):
        self.file_path = Path(file_path)
        self.data = []

    def is_product_name(self, text):
        # Проверяет, начинается ли текст со слова из PRODUCT_NAMES
        return any(text.lower().startswith(name.lower()) for name in self.PRODUCT_NAMES)

    def contains_product_name(self, text):
        # Обязательная проверка: текст должен содержать хотя бы одно слово из PRODUCT_NAMES
        return any(name.lower() in text.lower() for name in self.PRODUCT_NAMES)

    def detect_engine(self):
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
        df = pd.read_excel(self.file_path, header=None, engine=engine)
        if df.shape[1] == 1:
            self.parse_single_column(df)
        else:
            name_column = None
            for col in df.columns:
                for row_idx in range(min(15, len(df))):
                    cell_value = str(df.iloc[row_idx, col]).lower() if pd.notna(df.iloc[row_idx, col]) else ""
                    if "наименование" in cell_value:
                        name_column = col
                        break
                if name_column is not None:
                    break
            if name_column is None:
                self.parse_single_column(df)
            else:
                extra_char = (name_column + 2 < df.shape[1])
                self.parse_multi_column(df, name_column, extra_char)

    def parse_single_column(self, df):
        current_text = ""
        for index, row in df.iterrows():
            cell_value = str(row[0]).strip()
            if self.is_product_name(cell_value):
                if current_text and self.contains_product_name(current_text):
                    self.data.append({"text": current_text})
                current_text = cell_value
            else:
                current_text += " " + cell_value
        if current_text and self.contains_product_name(current_text):
            self.data.append({"text": current_text})

    def parse_multi_column(self, df, name_column, extra_char):
        current_text = ""
        for index, row in df.iterrows():
            if pd.isna(row[name_column]):
                continue
            cell_value = str(row[name_column]).strip().replace("\t", " ")
            char_value = ""
            if extra_char:
                if name_column + 1 < df.shape[1]:
                    char_value += " " + str(row[name_column + 1]).strip().replace("\t", " ").replace("\n", " ")
                if name_column + 2 < df.shape[1]:
                    char_value += " " + str(row[name_column + 2]).strip().replace("\t", " ").replace("\n", " ")
            else:
                if name_column + 1 < df.shape[1]:
                    char_value = " " + str(row[name_column + 1]).strip().replace("\t", " ").replace("\n", " ")
            if self.is_product_name(cell_value):
                if current_text and self.contains_product_name(current_text):
                    self.data.append({"text": current_text})
                current_text = cell_value + char_value
            else:
                current_text += " " + cell_value + char_value
        if current_text and self.contains_product_name(current_text):
            self.data.append({"text": current_text})

    def process(self):
        self.parse_excel()








class StructuredPdfParser:
    PRODUCT_NAMES = settings.product_names
    EXCLUDE_WORDS = ["шт.", "шт", "штук"]

    def __init__(self, file_path):
        self.file_path = Path(file_path)
        self.data = []
        self.header_candidates = []  # список токенов из первых пяти найденных заголовков товара
        self.header_mask = None      # регулярное выражение для определения начала нового товара

    def determine_common_pattern(self):
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
                                self.data.append({"text": current_record})
                            current_record = row_combined
                        else:
                            # Если паттерн установлен и строка соответствует началу нового товара,
                            # то считаем её новым заголовком
                            if self.header_mask and self.is_new_header(row_combined):
                                if current_record:
                                    self.data.append({"text": current_record})
                                current_record = row_combined
                            else:
                                # Иначе, строка считается продолжением предыдущего товара
                                if current_record:
                                    if current_record.endswith('-'):
                                        current_record = current_record.rstrip('-') + row_combined.lstrip()
                                    else:
                                        current_record += " " + row_combined
        if current_record:
            self.data.append({"text": current_record})

    def process(self):
        if not self.file_path.exists():
            print(f"File not found: {self.file_path}")
            return

        print(f"Processing file: {self.file_path.name}")
        self.parse_pdf()


class StructuredPdfParser2:
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