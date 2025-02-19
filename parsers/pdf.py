import pdfplumber
import pandas as pd

from .base import BaseParser


class ParserPDF(BaseParser):
    def get_dataframes(self) -> list[pd.DataFrame]:
        with pdfplumber.open(self.path_to_file) as file_pdf:
            dataframes = []

            for page in file_pdf.pages:
                table = page.extract_table()

                if table is not None:
                    dataframes.append(pd.DataFrame(table[1:], columns=table[0]))

        return dataframes