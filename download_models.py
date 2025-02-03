from sentence_transformers import SentenceTransformer

# Загружаем и сохраняем модель локально
model = SentenceTransformer('sentence-transformers/all-MiniLM-L6-v2')
model.save("models/all-MiniLM-L6-v2")
