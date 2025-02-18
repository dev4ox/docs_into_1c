from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.templating import Jinja2Templates
from pathlib import Path
import shutil
import pandas as pd
from parsers import (
    StructuredPdfParser,
    extract_text_from_docx,
    UnifiedExcelParser,
    append_df_to_excel,
    final_columns
)
from run_models import (
    extract_with_mistral
)


app = FastAPI()
templates = Jinja2Templates(directory="templates")


# Главная страница
@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

# Эндпоинт загрузки файла
@app.post("/upload", response_class=HTMLResponse)
async def upload_file(request: Request, file: UploadFile = File(...)):
    upload_folder = Path("uploads")
    upload_folder.mkdir(exist_ok=True)
    file_path = upload_folder / file.filename
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    
    ext = file_path.suffix.lower()
    if ext in [".xlsx", ".xls", ".xlsm"]:
        extracted_text = UnifiedExcelParser(file_path)
        extracted_text.process()
    elif ext in [".doc", ".docx"]:
        extracted_text = extract_text_from_docx(file_path)
    elif ext == ".pdf":
        extracted_text = StructuredPdfParser(file_path)
        extracted_text.process()
    else:
        return templates.TemplateResponse("index.html", {"request": request, "message": "Неподдерживаемый формат файла."})
    
    # Передаём извлечённый текст в LLM для получения JSON с параметрами.
    # Если в тексте несколько товаров (например, для Excel), можно реализовать итерацию – здесь для простоты объединяем в один текст.
    result_json = extract_with_mistral(extracted_text.print_data())
    
    # Если LLM вернула пустой словарь, заменяем дефолтными значениями.
    if not result_json or not isinstance(result_json, dict) or len(result_json) == 0:
        result_json = {col: "не указано" for col in final_columns}
    else:
        for col in final_columns:
            if col not in result_json:
                result_json[col] = "не указано"
    
    # Создаем DataFrame из результата
    df_result = pd.DataFrame([result_json], columns=final_columns)
    
    # Сохраняем или дописываем результат в файл "Форма_2.xlsx" в папке "downloads"
    output_folder = Path("downloads")
    output_folder.mkdir(exist_ok=True)
    output_file = output_folder / "Форма_2.xlsx"
    if not output_file.exists():
        pd.DataFrame(columns=final_columns).to_excel(output_file, index=False, sheet_name="Sheet1")
    append_df_to_excel(str(output_file), df_result, sheet_name="Sheet1")
    
    return templates.TemplateResponse("result.html", {"request": request, "output_file": str(output_file)})

# Эндпоинт для скачивания файла
@app.get("/download", response_class=FileResponse)
async def download_file():
    file_path = Path("downloads") / "Форма_2.xlsx"
    return FileResponse(path=file_path, filename="Форма_2.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
