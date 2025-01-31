from sentence_transformers import SentenceTransformer, util
import pandas as pd

# Загружаем модель
model = SentenceTransformer('sentence-transformers/all-MiniLM-L6-v2')

# Определяем список столбцов шаблона (основные параметры)
columns = {
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

# Создаем эмбеддинги для заголовков столбцов
column_embeddings = {col: model.encode(terms) for col, terms in columns.items()}


# Функция для определения соответствия параметра колонке
def classify_parameter(text):
    text_embedding = model.encode(text)
    best_match = None
    best_score = 0.0

    for col, embeddings in column_embeddings.items():
        # Сравниваем входной текст со всеми синонимами столбца
        scores = util.pytorch_cos_sim(text_embedding, embeddings)
        max_score = scores.max().item()

        if max_score > best_score:
            best_score = max_score
            best_match = col

    return best_match if best_score > 0.5 else "Прочее"


# Пример парсинга и классификации текста
parsed_text = [
    "Светильник мощностью 50 Вт",
    "Защита от пыли и влаги IP65",
    "Цветовая температура 4000K",
    "Материал корпуса - алюминий"
]

classified_data = {col: "" for col in columns}
classified_data["Прочее"] = []

for text in parsed_text:
    col = classify_parameter(text)
    if col != "Прочее":
        classified_data[col] += text + "; "
    else:
        classified_data["Прочее"].append(text)

# Преобразуем в Pandas DataFrame
df = pd.DataFrame([classified_data])

# Сохраняем в Excel
df.to_excel("output.xlsx", index=False)
