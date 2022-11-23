import os
import argparse
import time

import openpyxl
import Levenshtein


translators = ['%DeepL', '%Yandex.Переводчик', '%Google Translate']
original_column = 1
translate_columns = (2, 5)
value_columns = (5, 8)


def arg_parse():
    parser = argparse.ArgumentParser()
    parser.add_argument('-f', dest='filepath', type=str)
    return parser.parse_args()


def get_levenshtein_value(original, translate):
    return 1 - Levenshtein.distance(original.value, translate.value) / min(len(original.value), len(translate.value))


def get_lev_values(original, translates):
    lev_values = [get_levenshtein_value(original, translate)
                  if translate else 0
                  for translate in translates]
    return lev_values


def xlsx_handle(filepath):
    if not os.path.exists(filepath):
        raise FileNotFoundError
    wb = openpyxl.load_workbook(filepath)
    wba = wb.active
    for row in wba.iter_rows(min_row=2):
        original = row[original_column]
        translates = [row[column] for column in range(*translate_columns)]
        lev_values = get_lev_values(original, translates)
        for column, lev_value in zip(range(*value_columns), lev_values):
            row[column].value = lev_value * 100
    wb.save(f"{''.join(filepath.split('.')[:-1])}_handled.xlsx")


if __name__ == '__main__':
    start_time = time.time()
    filepath = arg_parse().filepath
    xlsx_handle(filepath)
    print(f'Выполнение программы заняло {time.time() - start_time} секунд!')
