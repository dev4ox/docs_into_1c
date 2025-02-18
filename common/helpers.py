from pathlib import Path

from openpyxl import load_workbook


def resize_column_in_output_xlsx(path: Path):
    # 🔹 Загружаем созданный файл
    book = load_workbook(path)
    sheet = book.active

    # 🔹 Автоматическая ширина колонок
    for col in sheet.columns:
        max_length = 0
        col_letter = col[0].column_letter  # Получаем букву колонки (A, B, C...)

        for cell in col:
            try:
                max_length = max(max_length, len(str(cell.value)))  # Максимальная длина значения

            except:
                pass

        adjusted_width = max_length + 2  # Небольшой запас по ширине
        sheet.column_dimensions[col_letter].width = adjusted_width  # Устанавливаем ширину

    # 🔹 Сохраняем изменения
    book.save(path)


def convert_list_to_string_with_comma(product_data: dict) -> dict:
    output_dict = {}

    for key, value in product_data.items():
        output_dict[key] = ", ".join(value)

    return output_dict