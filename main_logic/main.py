import os
import logging
import sys
from pathlib import Path
from create_output_excels import convert_json_to_excel, find_column_index
from check_input_excel import check_excel
from create_report import create_report
from utils.utils import convert_csv_to_excel_in_folder
from utils.constants import (
    INPUT_DATA_COLUMN_NAME,
    OUTPUT_DATA_COLUMN_NAME,
    INPUT_DATA_DIRECTORY_PATH,
    OUTPUT_INPUT_DATA_FORMAT_CSV_DIRECTORY_PATH,
    OUTPUT_INPUT_DATA_FORMAT_XLSX_DIRECTORY_PATH,
    OUTPUT_REPORTS_DIRECTORY_PATH,
    OUTPUT_AUXILIARY_FILES_DIRECTORY_PATH
)


def counting_statistics_kpss_cnts() -> None:
    """Формирует три итоговых файлика - распознанные значения, верифицированные значения,
    а также отчет по качеству распознавания
    """
    logging.info(
        f'Преобразую исходные файлы из формата .csv в формат .xlsx, а также копирую исходники в итоговую папку OUTPUT'
    )
    convert_csv_to_excel_in_folder(
        INPUT_DATA_DIRECTORY_PATH,
        OUTPUT_INPUT_DATA_FORMAT_CSV_DIRECTORY_PATH,
        OUTPUT_INPUT_DATA_FORMAT_XLSX_DIRECTORY_PATH
    )
    # Создаем папки для репортов и вспомогательных файлов
    os.makedirs(OUTPUT_REPORTS_DIRECTORY_PATH, exist_ok=True)
    os.makedirs(OUTPUT_AUXILIARY_FILES_DIRECTORY_PATH, exist_ok=True)

    logging.info(f'Начинаем сбор статистики')

    for filepath in Path(OUTPUT_INPUT_DATA_FORMAT_XLSX_DIRECTORY_PATH).glob('*.xlsx'):
        check_excel(filepath)

    # input_request_filename = convert_json_to_excel(find_column_index(DOCUMENT_INPUT_REQUEST), INPUT_DATA_COLUMN_NAME)
    # verification_request_filename = convert_json_to_excel(find_column_index(DOCUMENT_VERIFICATION_REQUEST),
    #                                                          OUTPUT_DATA_COLUMN_NAME)
    #
    # filenames_for_prepare_report = input_request_filename, verification_request_filename
    #
    # create_report(filenames_for_prepare_report)
    # logging.info(f'Отчет создан')


if __name__ == '__main__':
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - [%(levelname)s] - %(message)s -Line: %(lineno)d',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(f'{__file__}.log', encoding='utf-8'),
        ]
    )
    counting_statistics_kpss_cnts()
