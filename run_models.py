import os
import json
import pandas as pd
from pathlib import Path
import importlib.util
from llama_cpp import Llama
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

BASE_DIR = Path(__file__).resolve().parent
SETTINGS_PATH = BASE_DIR / "settings.py"
spec = importlib.util.spec_from_file_location("settings", SETTINGS_PATH)
settings = importlib.util.module_from_spec(spec)
spec.loader.exec_module(settings)


def extract_gemma_2_2b_it_IQ3_M(text, initial_prompt):
    llm = Llama(
        model_path="../../.lmstudio/models/lmstudio-community/gemma-2-2b-it-GGUF/gemma-2-2b-it-IQ3_M.gguf",
        n_ctx=3072
    )
    prompt = initial_prompt + "\n\nДанные для обработки:\n" + text
    output = llm(prompt=prompt, max_tokens=1024, temperature=0.1)
    result_text = output["choices"][0]["text"].strip()
    print(result_text)
    try:
        data = json.loads(result_text)
    except Exception as e:
        print("Error parsing JSON:", e)
        print("Raw output:", result_text)
        data = {}
    return data


def extract_gemma_2_2b_it_Q6_K(text, initial_prompt):
    llm = Llama(
        model_path="../../.lmstudio/models/lmstudio-community/gemma-2-2b-it-GGUF/gemma-2-2b-it-Q6_K.gguf",
        n_ctx=3072
    )
    prompt = initial_prompt + "\n\nДанные для обработки:\n" + text
    output = llm(prompt=prompt, max_tokens=1024, temperature=0.1)
    result_text = output["choices"][0]["text"].strip()
    print(result_text)
    try:
        data = json.loads(result_text)
    except Exception as e:
        print("Error parsing JSON:", e)
        print("Raw output:", result_text)
        data = {}
    return data


def extract_with_mistral(text, initial_prompt):
    llm = Llama(
        model_path="../../.lmstudio/models/lmstudio-community/Mistral-7B-Instruct-v0.3-GGUF/Mistral-7B-Instruct-v0.3-Q4_K_M.gguf",
        n_ctx=3072
    )
    prompt = initial_prompt + "\n\nДанные для обработки:\n" + text
    output = llm(prompt=prompt, max_tokens=1024, temperature=0.1)
    result_text = output["choices"][0]["text"].strip()
    print(result_text)
    try:
        data = json.loads(result_text)
    except Exception as e:
        print("Error parsing JSON:", e)
        print("Raw output:", result_text)
        data = {}
    return data


def append_df_to_excel(filename, df, sheet_name='Sheet1'):
    if os.path.exists(filename):
        wb = load_workbook(filename)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        for row in dataframe_to_rows(df, index=False, header=False):
            ws.append(row)
        wb.save(filename)
    else:
        df.to_excel(filename, index=False, sheet_name=sheet_name)


class UnifiedExcelParser:
    PRODUCT_NAMES = settings.product_names

    def __init__(self, file_path):
        self.file_path = Path(file_path)
        self.data = []

    def is_product_name(self, text):
        return any(text.lower().startswith(name.lower()) for name in self.PRODUCT_NAMES)

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
                if current_text:
                    self.data.append({"text": current_text})
                current_text = cell_value
            else:
                current_text += " " + cell_value
        if current_text:
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
            if any(name.lower() in cell_value.lower() for name in self.PRODUCT_NAMES):
                if current_text:
                    self.data.append({"text": current_text})
                current_text = cell_value + char_value
            else:
                current_text += " " + cell_value + char_value
        if current_text:
            print(current_text)
            self.data.append({"text": current_text})

    def process(self):
        self.parse_excel()