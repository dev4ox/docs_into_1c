import uvicorn
from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
import multiprocessing
import pandas as pd
import shutil
import run_models
from pathlib import Path
from common.constants import CWD

import main


final_columns = ["Номенклатура", "Мощность, Вт", "Св. поток, Лм", "IP", "Габариты", "Длина, мм",
                 "Ширина, мм", "Высота, мм", "Рассеиватель", "Цвет. температура, К", "Вес, кг",
                 "Напряжение, В", "Температура эксплуатации", "Срок службы (работы) светильника",
                 "Тип КСС", "Род тока", "Гарантия", "Индекс цветопередачи (CRI, Ra)", "Цвет корпуса",
                 "Коэффициент пульсаций", "Коэффициент мощности (Pf)", "Класс взрывозащиты (Ex)",
                 "Класс пожароопасности", "Класс защиты от поражения электрическим током",
                 "Материал корпуса", "Тип", "Прочее"]

allowed_extensions = {".doc", ".docx", ".xlsx", ".xls", ".xlsm", ".pdf"}

app = FastAPI()
templates = Jinja2Templates(directory="templates")


# Главная страница
@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/upload", response_class=HTMLResponse)
async def upload_file(request: Request, file: UploadFile = File(...)):
    # Если файл не загружен или неподдерживаемый код файла
    if not file or not any(file.filename.endswith(ext) for ext in allowed_extensions):
        return RedirectResponse(url="/?message=Невозможно обработать данный файл", status_code=301)

    upload_folder = Path("uploads")
    upload_folder.mkdir(exist_ok=True)
    # Сохранение файла во входящем
    input_file_path = upload_folder / run_models.generate_filename(Path(file.filename).stem,
                                                                   Path(file.filename).suffix.lower())
    print(f"{input_file_path=}")
    with open(input_file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Функция определения расширения файла (точка входа в парсер)
    ext = input_file_path.suffix.lower()
    if ext in [".xlsx", ".xls", ".xlsm"]:
        parser = run_models.UnifiedExcelParser(input_file_path)
        parser.process()
    elif ext in [".doc", ".docx", ".pdf"]:
        main.main(input_file_path)
        parser = run_models.UnifiedExcelParser(Path(CWD, 'test_data', 'output', 'intermediate.xlsx'))
        parser.process()
    else:
        return templates.TemplateResponse("index.html",
                                          {"request": request, "message": "Не удалось найти парсер."})

    filled_forms = []

    # На вход подаётся список со словарями, где [{"text": "Имя...характеристики"}, ...], 1 словарь = 1 позиция товара
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


    df_form_in = pd.DataFrame(filled_forms, columns=final_columns)
    # ФИЛЬТР!!!
    df_form_out = run_models.out_filter_dataframe(df_form_in)
    output_folder = Path("downloads")
    output_folder.mkdir(exist_ok=True)
    output_filename = run_models.generate_filename()
    output_file = output_folder / output_filename
    if not output_file.exists():
        pd.DataFrame(columns=final_columns).to_excel(output_file, index=False, sheet_name="Sheet1")
    run_models.append_df_to_excel(output_file, df_form_out, sheet_name="Sheet1")
    print(f"\nДанные успешно добавлены в файл {output_file}.")
    return templates.TemplateResponse("result.html",{
                                        "request": request,
                                        "output_file": str(output_file),
                                        "download_url": output_filename})


# Эндпоинт для скачивания файла
@app.get("/download/{filename}", response_class=FileResponse)
async def download_file(filename: str):
    file_path = Path("downloads") / filename
    if not file_path.exists():
        return {"error": "Файл не найден"}
    return FileResponse(path=file_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def run_server() -> None:
    uvicorn.run("run:app", host="0.0.0.0", port=8000)


if __name__ == "__main__":
    server = multiprocessing.Process(target=run_server)
    server.start()
    server.join()
