from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.templating import Jinja2Templates
from pathlib import Path
import pandas as pd
import shutil
import run_models
import settings


final_columns = ["Номенклатура", "Мощность, Вт", "Св. поток, Лм", "IP", "Габариты", "Длина, мм",
                 "Ширина, мм", "Высота, мм", "Рассеиватель", "Цвет. температура, К", "Вес, кг",
                 "Напряжение, В", "Температура эксплуатации", "Срок службы (работы) светильника",
                 "Тип КСС", "Род тока", "Гарантия", "Индекс цветопередачи (CRI, Ra)", "Цвет корпуса",
                 "Коэффициент пульсаций", "Коэффициент мощности (Pf)", "Класс взрывозащиты (Ex)",
                 "Класс пожароопасности", "Класс защиты от поражения электрическим током",
                 "Материал корпуса", "Тип", "Прочее"]


app = FastAPI()
templates = Jinja2Templates(directory="templates")


# Главная страница
@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/upload", response_class=HTMLResponse)
async def upload_file(request: Request, file: UploadFile = File(...)):
    upload_folder = Path("uploads")
    upload_folder.mkdir(exist_ok=True)
    input_file_path = upload_folder / file.filename
    with open(input_file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    ext = input_file_path.suffix.lower()
    if ext in [".xlsx", ".xls", ".xlsm"]:
        parser = run_models.UnifiedExcelParser(input_file_path)
        parser.process()
    elif ext in [".doc", ".docx"]:
        # parser = run_models.(input_file_path)
        pass
    elif ext == ".pdf":
        parser = run_models.StructuredPdfParser2(input_file_path)
        parser.process()
    else:
        return templates.TemplateResponse("index.html",
                                          {"request": request, "message": "Неподдерживаемый формат файла."})

    filled_forms = []
    for product in parser.data:
        product_text = product["text"]
        print(f"Распознанный товар: {product_text=}")
        extracted = run_models.extract_gemma_2_2b_it_IQ3_M(product_text, final_columns)

        if not extracted or not isinstance(extracted, dict) or len(extracted) == 0:
            extracted = {col: "не указано" for col in final_columns}
        else:
            # Проверка, что все ключи есть
            for col in final_columns:
                if col not in extracted:
                    extracted[col] = "не указано"
        print(f"Извлечённый товар: {extracted=}")
        filled_forms.append(extracted)

    df_form = pd.DataFrame(filled_forms, columns=final_columns)
    output_folder = Path("output")
    output_folder.mkdir(exist_ok=True)
    output_file = output_folder / "Форма_2.xlsx"
    if not output_file.exists():
        pd.DataFrame(columns=final_columns).to_excel(output_file, index=False, sheet_name="Sheet1")
    run_models.append_df_to_excel(output_file, df_form, sheet_name="Sheet1")
    print(f"\nДанные успешно добавлены в файл {output_file}.")
    return templates.TemplateResponse("result.html", {"request": request, "output_file": str(output_file)})


# Эндпоинт для скачивания файла
@app.get("/download", response_class=FileResponse)
async def download_file():
    file_path = Path("output") / "Форма_2.xlsx"
    return FileResponse(path=file_path, filename="Форма_2.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
