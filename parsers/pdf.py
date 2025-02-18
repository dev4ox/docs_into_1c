from pathlib import Path

import pdfplumber
import pandas as pd
from fuzzywuzzy import fuzz

from common.constants import DIR_TEST_DATA, PRODUCT_NAMES, SYNONYMS


PATH_TO_TEST_PDF = Path(DIR_TEST_DATA, "ТЗ для РИР.pdf")


class ParserPDF:
    def __init__(self, path_to_pdf: str | Path):
        if type(path_to_pdf) is str:
            self.path_to_pdf = Path(path_to_pdf)

        else:
            self.path_to_pdf = path_to_pdf

    def __new__(cls, *args, **kwargs) -> dict:
        instance = super().__new__(cls)
        instance.__init__(*args, **kwargs)

        result: dict = instance.__parse()

        return result

    @staticmethod
    def __check_characteristic_partial_ratio(string: str, synonym: str) -> bool:
        ratio = fuzz.partial_ratio(string.lower(), synonym.lower())

        if ratio >= 50:
            return True

        return False

    def check_characteristic(self, string: str) -> bool:
        is_characteristics: list[bool] = []

        for name_synonym, synonyms in SYNONYMS.items():

            if string is not None:
                is_characteristic = self.__check_characteristic_partial_ratio(string, name_synonym)
                is_characteristics.append(is_characteristic)

                for synonym in synonyms:
                    is_characteristic = self.__check_characteristic_partial_ratio(string, synonym)
                    is_characteristics.append(is_characteristic)

        if True in is_characteristics:
            return True

        else:
            return False

    @staticmethod
    def check_product_name(string: str) -> tuple[bool, int]:
        ratio = 0

        for name in PRODUCT_NAMES:
            if string is not None:
                ratio = fuzz.token_set_ratio(string.lower(), name.lower())

            if ratio == 100:
                return True, ratio

        return False, ratio

    def __parse(self) -> dict[str, list[str]]:
        with pdfplumber.open(self.path_to_pdf) as file_pdf:
            product_name: str | None = None
            product_data = {}

            for page in file_pdf.pages:
                table = page.extract_table()

                if table is not None:
                    df_values = pd.DataFrame(table[1:], columns=table[0]).values
                    print(df_values)

                    for value in df_values:
                        for i in value:
                            is_product_name, ratio = self.check_product_name(i)

                            if is_product_name:
                                print(f"\nname: {i}, ratio: {ratio}\n")
                                product_name = i
                                product_data[product_name] = []
                                continue

                            if self.check_characteristic(i) and product_name is not None:
                                print(f"\ncharacteristic: {i}\n")
                                product_data[product_name].append(i)

        return product_data