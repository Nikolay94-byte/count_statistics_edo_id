import logging
import os
import sys
from tabulate import tabulate
import pandas as pd

from check_input_data import check_files
from main_logic.create_report import create_report
from utils import constants
from utils.constants import (
    INPUT_DATA_DIRECTORY_PATH,
    OUTPUT_REPORTS_DIRECTORY_PATH,
)
from utils.utils import parse_bd_file, decoding_csv, expand_dataframe_data


def counting_statistics_table_pim() -> None:
    """Считает статистику качество извлечения таблиц при отправки в ПИМ"""

    os.makedirs(OUTPUT_REPORTS_DIRECTORY_PATH, exist_ok=True)

    etalon_file_path, recognized_file_path, doc_type = check_files(INPUT_DATA_DIRECTORY_PATH)

    logging.info(f"Читаем файл с эталонными данными {etalon_file_path} и разворачиваем значения")
    etalon_df = pd.read_excel(etalon_file_path, dtype=str).fillna("")
    expanded_etalon_df = expand_dataframe_data(etalon_df, constants.ETALON_VALUE)
    print(tabulate(expanded_etalon_df, headers='keys', tablefmt='grid', showindex=False))
    print("\n" + "=" * 80 + "\n")

    logging.info(f"Читаем файл с распознанными данными {recognized_file_path}")
    if os.path.splitext(recognized_file_path)[1].lower() == '.csv':
        recognized_df = decoding_csv(recognized_file_path)
    else:
        recognized_df = pd.read_excel(recognized_file_path, dtype=str).fillna("")
    logging.info(f"Парсим файл с распознанными данными {recognized_file_path} и разворачиваем значения")
    parsed_recognized_df = parse_bd_file(recognized_df)
    expanded_recognized_df = expand_dataframe_data(parsed_recognized_df, constants.RECOGINIZED_VALUE)
    print(tabulate(expanded_recognized_df, headers='keys', tablefmt='grid', showindex=False))
    print("\n" + "=" * 80 + "\n")

    logging.info(f'Создаем отчет')
    quality_percent = create_report(expanded_etalon_df, expanded_recognized_df, doc_type)
    logging.info(
        f"Среднее качество извлечения таблиц в шаблоне {doc_type}: {quality_percent}"
    )

if __name__ == '__main__':
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s-[%(levelname)s]- Module: %(module)s - Line:%(lineno)d - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(f'{__file__}.log', encoding='utf-8'),
        ]
    )
    counting_statistics_table_pim()
