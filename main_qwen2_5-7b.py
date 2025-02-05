import os
import json
import pandas as pd
from pathlib import Path
import importlib.util
from llama_cpp import Llama
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# ---------------- Загрузка настроек из settings.py ----------------
SETTINGS_PATH = "settings.py"

spec = importlib.util.spec_from_file_location("settings", SETTINGS_PATH)
settings = importlib.util.module_from_spec(spec)
spec.loader.exec_module(settings)

# ---------------- Загрузка локальной модели Qwen2.5 ----------------
# Укажите корректный путь к файлу модели (формата GGUF)
llm = Llama(model_path="../models/Qwen2.5-7B-Instruct-1M-GGUF/Qwen2.5-7B-Instruct-1M-Q4_K_M.gguf", n_threads=4)

# ---------------- Начальный промпт для модели ----------------
initial_prompt = """
You are an expert data extractor specialized in extracting product information from text.
When given a text about a product, extract the following fields in JSON format:
"Номенклатура", "Мощность, Вт", "Св. поток, Лм", "IP", "Габариты", "Длина, мм", "Ширина, мм", "Высота, мм", 
"Рассеиватель", "Цвет. температура, К", "Вес, кг", "Напряжение, В", "Температура эксплуатации", "Срок службы (работы) светильника", 
"Тип КСС", "Род тока", "Гарантия", "Индекс цветопередачи (CRI, Ra)", "Цвет корпуса", "Коэффициент пульсаций", 
"Коэффициент мощности (Pf)", "Класс взрывозащиты (Ex)", "Класс пожароопасности", "Класс защиты от поражения электрическим током", 
"Материал корпуса", "Тип", "Прочее".
For each field, if no value is found, output "не указано".
Return only a JSON object.
"""


def extract_with_qwen(text):
    """
    Передаёт модели Qwen2.5 входной текст вместе с промптом и пытается получить структурированный JSON.
    """
    prompt = initial_prompt + "\n\nText:\n" + text + "\n\nExtracted JSON:"
    output = llm(prompt=prompt, max_tokens=300, temperature=0.0)
    result_text = output["choices"][0]["text"].strip()
    try:
        data = json.loads(result_text)
    except Exception as e:
        print("Error parsing JSON:", e)
        print("Raw output:", result_text)
        data = {}
    return data


def append_df_to_excel(filename, df, sheet_name='Sheet1'):
    """
    Если файл filename существует, функция находит последнюю заполненную строку
    и добавляет новые данные (без заголовков). Если файла нет, создается новый Excel-файл с заголовками.
    """
    if os.path.exists(filename):
        wb = load_workbook(filename)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.active
        for row in dataframe_to_rows(df, index=False, header=False):
            ws.append(row)
        wb.save(filename)
    else:
        df.to_excel(filename, index=False, sheet_name=sheet_name)


# ---------------- Основная логика: обработка входного словаря ----------------
if __name__ == "__main__":
    # Пример входного словаря, где ключ – номер товара, а значение – строка с информацией о товаре
    input_data = {
        "1": (
            "Мощность светильника 50 Вт, цветовая температура: 4000 К, защита от пыли и влаги IP65, "
            "масса: 1.2 кг; материал корпуса: алюминий, тип крепления: подвесной, угол излучения 120°, "
            "CRI>80, напряжение питания: 220 В, класс защиты I, гарантийный срок 36 месяцев."
        ),
        "2": (
            "Наименование товара: Светодиодный прожектор, мощность 150 Вт, "
            "световой поток 12000 Лм, IP65, размеры 325х155х440 мм, напряжение 220-240 В, срок службы 5 лет."
        )
    }

    # Итоговый порядок столбцов для заполненной формы
    final_columns = [
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

    filled_forms = []

    # Для каждого товара из входного словаря вызываем модель Qwen2.5 для извлечения параметров
    for prod_id, prod_text in input_data.items():
        extracted = extract_with_qwen(prod_text)
        # Гарантируем, что все поля присутствуют
        for col in final_columns:
            if col not in extracted:
                extracted[col] = "не указано"
        filled_forms.append(extracted)

    df_form = pd.DataFrame(filled_forms, columns=final_columns)
    print("\nЗаполненная форма:")
    print(df_form.to_string(index=False))

    output_file = "output.xlsx"
    append_df_to_excel(output_file, df_form, sheet_name="Sheet1")
    print(f"\nДанные успешно добавлены в файл {output_file}.")
