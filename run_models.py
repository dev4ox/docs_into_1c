from pathlib import Path
from datetime import datetime
from llama_cpp import Llama
from common.constants_prod import DIR_MODELS
import settings
import os
import json
import re
import pandas as pd


input_prompt = '''
Задача – анализ текста и извлечение параметров.
Выводи найденные товары, имеющие характеристики в формате JSON строго по инструкции. Обычно товары имеют следующие характеристики:
"Номенклатура", "Мощность, Вт", "Св. поток, Лм", "IP", "Длина, мм", "Ширина, мм", "Высота, мм", "Габариты", "Рассеиватель", "Цвет. температура, К", "Вес, кг", "Напряжение, В", "Срок службы (работы) светильника", "Температура эксплуатации", "Материал корпуса", "Тип", "Тип КСС", "Род тока", "Гарантия", "Индекс цветопередачи (CRI, Ra)", "Класс защиты от поражения электрическим током", "Коэффициент мощности (Pf)", "С регулятором яркости (диммирование)", "Ударопрочность", "Класс взрывозащиты (Ex)", "Класс пожароопасности", "Цвет корпуса", "Коэффициент пульсаций", "Прочее".
Может быть более одного товара на странице, выводи последовательно товары в формате JSON.

Если значение выражено диапазоном или с квалификаторами (например, "не более", "не менее", "от X до Y", "±10", "+-10", "около"), включай всю фразу с единицами измерения.
Если параметр отсутствует или его значение не может быть корректно извлечено, верни "не указано".
Если есть дополнительные характеристики товара, не подходящие под обычные характеристики, помести их в характеристику "Прочее".

Пример 1:
Входной текст:
"Наименование продукции: Прожектор светодиодный ASD СДО-2-20 20W или аналог. Энергопотребление, не более, Вт: 20; Входное напряжение: 85-265 В; Цветовая температура, К, не менее: 6500; Коэффициент пульсаций, не более: 5%; Угол свечения: 120°; Степень защиты, не менее IP: 65; Световой поток, Лм: не менее 1600; Габаритные размеры (L, b, h): 178*100*138; Время работы, не менее: 50 000 часов; Кронштейн крепления."
Вывод:
{
  "Номенклатура": "Прожектор светодиодный ASD СДО-2-20 20W или аналог",
  "Мощность, Вт": "не более 20 Вт",
  "Св. поток, Лм": "не менее 1600 Лм",
  "IP": "не менее 65",
  "Длина, мм": "178",
  "Ширина, мм": "100",
  "Высота, мм": "138",
  "Габариты": "178*100*138",
  "Рассеиватель": "не указано",
  "Цвет. температура, К": "не менее 6500",
  "Вес, кг": "не указано",
  "Напряжение, В": "85-265 В",
  "Срок службы (работы) светильника": "не менее 50 000 часов",
  "Температура эксплуатации": "не указано",
  "Материал корпуса": "не указано",
  "Тип": "не указано",
  "Тип КСС": "120°",
  "Род тока": "не указано",
  "Гарантия": "не указано",
  "Индекс цветопередачи (CRI, Ra)": "не указано",
  "Класс защиты от поражения электрическим током": "не указано",
  "Коэффициент мощности (Pf)": "не указано",
  "С регулятором яркости (диммирование)": "не указано",
  "Ударопрочность": "не указано",
  "Класс взрывозащиты (Ex)": "не указано",
  "Класс пожароопасности": "не указано",
  "Цвет корпуса": "не указано",
  "Коэффициент пульсаций": "не указано",
  "Прочее": "Кронштейн крепления"
}
'''


# НУЖНО Маленькая для OCR, 3Gb vram, с парсером работает збс!
def extract_gemma_2_2b_it_IQ3_M(text, final_columns):
    """
    Функция обрабатывает текст с помощью LLM модели Gemma 2, формирует корректный промпт,
    отправляет запрос и извлекает JSON-ответ.
    """
    llm = Llama(
        model_path=str(DIR_MODELS.joinpath("lmstudio-community", "gemma-2-2b-it-GGUF", "gemma-2-2b-it-IQ3_M.gguf")),
        n_ctx=8192,
        n_gpu_layers=-1,
        verbose=True,
    )
    prompt = f"<start_of_turn>user\n{input_prompt}\n\nText:\n{text}\n\nJSON:<end_of_turn>\n<start_of_turn>model\n"
    output = llm(
        prompt=prompt,
        max_tokens=2048,
        temperature=0.0,
        stop=["<end_of_turn>"]  # Останавливаем генерацию после ответа
    )

    # Извлекаем текст и определяем только JSON-объект с помощью регулярного выражения
    result_text = output["choices"][0]["text"].strip()
    match = re.search(r'\{.*\}', result_text, re.DOTALL)
    if match:
        result_text = match.group(0)

    try:
        data = json.loads(result_text)
    except Exception as e:
        print("Error parsing JSON:", e)
        print("Raw downloads:", result_text)
        data = {col: "не указано" for col in final_columns}
    return data


# НУЖНО генерирует уникальное имя файла
def generate_filename(prefix: str="Форма2", ext: str=".xlsx"):
    timestamp = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    return f"{prefix}-{timestamp}{ext}"


# НУЖНО Используем ExcelWriter в режиме добавления (append)
def append_df_to_excel(filename, df, sheet_name='Sheet1'):
    if os.path.exists(filename):
        with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            startrow = writer.sheets[sheet_name].max_row
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=startrow)
    else:
        df.to_excel(filename, index=False, sheet_name=sheet_name)

# НУЖНО (не)уникальный парсер excel
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
# todo: переделать на проверку светильника, а не наименование (пиздец сt циклом for)
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

