from pathlib import Path
import os

from openpyxl import load_workbook

import common.constants
from common.constants import DIR_DATA_INPUT, PATH_DATA_INTERMEDIATE_XLSX_FILE
from common.helpers import (
    convert_list_to_string_with_comma,
    resize_column_in_intermediate_xlsx,
    create_intermediate_xlsx,
)
from parsers.pdf import ParserPDF
from parsers.doc import DocParser


def activate_parsers(path_to_file: Path) -> None:
    file_type = path_to_file.suffix

    if file_type == ".pdf":
        parser_data = ParserPDF(path_to_file)
        print(f"\nПарс файла {path_to_file} - окончен успешно!\n")

    elif file_type == ".docx" or file_type == ".doc":
        parser_data = DocParser(path_to_file)
        print(f"\nПарс файла {path_to_file} - окончен успешно!\n")

    else:
        parser_data = {}
        print(f"\nНеизвестный формат файла: {file_type}\n")

    print(f"\nСохранение результатов парса в промежуточный файл")
    save_data_to_excel(parser_data, PATH_DATA_INTERMEDIATE_XLSX_FILE)


def save_data_to_excel(product_data: dict, path: Path) -> None:
    product_data = convert_list_to_string_with_comma(product_data)

    book = load_workbook(path)
    sheet = book.active

    # 🔹 Добавляем новую строку
    for name, value in product_data.items():
        sheet.append([name, value])

    # 🔹 Сохраняем изменения
    book.save(path)

    resize_column_in_intermediate_xlsx(path)
    print("Результаты парса сохранены!\n")


def main(input_file_path: Path | None = None) -> None:
    print(f"Промежуточный файл создан: {PATH_DATA_INTERMEDIATE_XLSX_FILE}\n")
    create_intermediate_xlsx(PATH_DATA_INTERMEDIATE_XLSX_FILE)
    if input_file_path:
        activate_parsers(input_file_path)
    else:
        paths_to_input_data = DIR_DATA_INPUT.glob("*.*")

        for path in paths_to_input_data:
            print(f"\nЗапуск парса файла: {path}\n")
            activate_parsers(path)


if __name__ == "__main__":
    try:
        main(Path(common.constants.CWD, 'uploads', 'ТЗ для Рос Волга-2025-02-19-11-19-07.doc'))

    except KeyboardInterrupt:
        os.remove(PATH_DATA_INTERMEDIATE_XLSX_FILE)

    except Exception as e:
        print(e)
        os.remove(PATH_DATA_INTERMEDIATE_XLSX_FILE)