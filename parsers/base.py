from pathlib import Path

import pandas as pd
from fuzzywuzzy import fuzz
import Levenshtein
from common.constants import SYNONYMS, PRODUCT_NAMES


class BaseParser:
    def __init__(self, path_to_file: Path):
        if type(path_to_file) is str:
            self.path_to_file = Path(path_to_file)

        else:
            self.path_to_file = path_to_file

    def __new__(cls, *args, **kwargs):
        instance = super().__new__(cls)
        instance.__init__(*args, **kwargs)

        result: dict = instance.__parse()

        return result

    def get_dataframes(self) -> list[pd.DataFrame]:
        pass

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
        product_name: str | None = None
        product_data = {}
        dataframes = self.get_dataframes()

        for dataframe in dataframes:
            dataframe_values = dataframe.values
            print(dataframe_values)

            for value in dataframe_values:
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

            product_name: str | None = None

        return product_data