import string
from typing import List

import pandas as pd


def get_table(filename: str) -> list[dict]:
    df = pd.read_excel(filename, sheet_name=0, header=None)
    df.columns = ["{" + string.ascii_uppercase[_] + "}" for _ in df.columns]
    return df.reset_index().to_dict("records")


if __name__ == "__main__":
    filename = ".\\tests\\data\\test_table.xlsx"
    print(get_table(filename))
