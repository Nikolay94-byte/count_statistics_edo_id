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
from utils.utils import parse_bd_file, find_file_by_name, decoding_csv, expand_dataframe_data


def counting_statistics_table_pim() -> None:
    """Считает статистику качество извлечения таблиц при отправки в ПИМ"""

    os.makedirs(OUTPUT_REPORTS_DIRECTORY_PATH, exist_ok=True)

    etalon_file_path = find_file_by_name(INPUT_DATA_DIRECTORY_PATH, constants.INPUT_DATA_ETALON_FILE)
    recognized_file_path = find_file_by_name(INPUT_DATA_DIRECTORY_PATH, constants.INPUT_DATA_RECOGINIZED_FILE)

    check_files(INPUT_DATA_DIRECTORY_PATH)

    # Читаем эталонный файл
    etalon_df = pd.read_excel(etalon_file_path, dtype=str).fillna("")

    # Разворачиваем значения в эталонном файле
    expanded_etalon_df = expand_dataframe_data(etalon_df, constants.ETALON_VALUE)
    expanded_etalon_df = expanded_etalon_df.sort_values([constants.FILE_NAME, constants.ATTRIBUTE_NAME]).reset_index(
        drop=True)
    print(tabulate(expanded_etalon_df, headers='keys', tablefmt='grid', showindex=False))
    print("\n" + "=" * 80 + "\n")

    # Читаем файл с распознанными данными
    if os.path.splitext(recognized_file_path)[1].lower() == '.csv':
        recognized_df = decoding_csv(recognized_file_path)
    else:
        recognized_df = pd.read_excel(recognized_file_path, dtype=str).fillna("")

    # Парсим файл с распознанными данными
    parsed_recognized_df = parse_bd_file(recognized_df)

    # Разворачиваем значения в файле с распознанными данными
    expanded_recognized_df = expand_dataframe_data(parsed_recognized_df, constants.RECOGINIZED_VALUE)
    print(tabulate(expanded_recognized_df, headers='keys', tablefmt='grid', showindex=False))
    print("\n" + "=" * 80 + "\n")

    logging.info(f'Создаем отчет по {Path(filepath).stem}')
    quality_percent = create_report()

    #
    # quality_percent_main_document = check_input_data_and_create_reports(
    #     OUTPUT_INPUT_DATA_FORMAT_XLSX_MAIN_DIRECTORY_PATH
    # )
    # quality_percent_application = check_input_data_and_create_reports(
    #     OUTPUT_INPUT_DATA_FORMAT_XLSX_APPLICATION_DIRECTORY_PATH
    # )
    #
    # show_in_logs_document_statistic(MAIN_DOC, quality_percent_main_document)
    # show_in_logs_document_statistic(APPLICATION, quality_percent_application)


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
