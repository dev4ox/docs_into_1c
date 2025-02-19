import os

from docx import Document
import pandas as pd

from common.helpers import convert_doc_to_docx
from .base import BaseParser


class DocParser(BaseParser):
    def __check_is_doc_file(self) -> bool:
        if self.path_to_file.suffix == ".doc":
            return True

        elif self.path_to_file.suffix == ".docx":
            return False

        else:
            raise TypeError(f"Unsupported file type: {self.path_to_file}")

    def get_dataframes(self) -> list[pd.DataFrame]:
        dataframes = []

        is_doc_file = self.__check_is_doc_file()

        if is_doc_file:
            self.path_to_file = convert_doc_to_docx(self.path_to_file)

        # 🔹 Открываем Word-файл
        doc = Document(str(self.path_to_file))

        for table in doc.tables:
            # 🔹 Извлекаем данные в список списков
            data = []
            for row in table.rows:
                data.append([cell.text.strip() for cell in row.cells])  # Убираем пробелы и символы новой строки

            # 🔹 Преобразуем в DataFrame
            dataframes.append(pd.DataFrame(data))

        if is_doc_file:
            os.remove(self.path_to_file)

        return dataframes