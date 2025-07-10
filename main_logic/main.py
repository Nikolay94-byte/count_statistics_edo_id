import logging
import os
import sys
from pathlib import Path

from check_input_excel import check_and_clean_file
from create_output_excels import convert_json_to_excel, find_column_index
from create_report import create_report
from utils.constants import (
    DOCUMENT_INPUT_REQUEST,
    DOCUMENT_VERIFICATION_REQUEST,
    INPUT_DATA_COLUMN_NAME,
    INPUT_DATA_DIRECTORY_PATH,
    OUTPUT_AUXILIARY_FILES_DIRECTORY_PATH,
    OUTPUT_DATA_COLUMN_NAME,
    OUTPUT_INPUT_DATA_FORMAT_CSV_DIRECTORY_PATH,
    OUTPUT_INPUT_DATA_FORMAT_XLSX_DIRECTORY_PATH,
    OUTPUT_REPORTS_DIRECTORY_PATH,
)
from utils.utils import convert_csv_to_excel_in_folder


def counting_statistics_kpss_cnts() -> None:
    """Формирует три итоговых файлика - распознанные значения, верифицированные значения,
    а также отчет по качеству распознавания
    """
    logging.info(
        f'Преобразуем исходные файлы из формата .csv в формат .xlsx, а также копируем исходники в итоговую папку OUTPUT'
    )
    convert_csv_to_excel_in_folder(
        INPUT_DATA_DIRECTORY_PATH,
        OUTPUT_INPUT_DATA_FORMAT_CSV_DIRECTORY_PATH,
        OUTPUT_INPUT_DATA_FORMAT_XLSX_DIRECTORY_PATH
    )
    # Создаем папки для отчетов и вспомогательных файлов
    os.makedirs(OUTPUT_REPORTS_DIRECTORY_PATH, exist_ok=True)
    os.makedirs(OUTPUT_AUXILIARY_FILES_DIRECTORY_PATH, exist_ok=True)

    for filepath in Path(OUTPUT_INPUT_DATA_FORMAT_XLSX_DIRECTORY_PATH).glob('*.xlsx'):
        logging.info(f'Начинаем проверку и редактирование исходника {Path(filepath).stem}')
        check_and_clean_file(filepath)

        logging.info(f'Создаем вспомогательный файл с распознанными данными по {Path(filepath).stem}')
        input_request_filename = convert_json_to_excel(filepath, find_column_index(filepath, DOCUMENT_INPUT_REQUEST), INPUT_DATA_COLUMN_NAME)
        logging.info(f'Создаем вспомогательный файл с верифицированными данными по {Path(filepath).stem}')
        verification_request_filename = convert_json_to_excel(filepath, find_column_index(filepath, DOCUMENT_VERIFICATION_REQUEST), OUTPUT_DATA_COLUMN_NAME)

        filenames_for_prepare_report = input_request_filename, verification_request_filename

        logging.info(f'Создаем отчет по {Path(filepath).stem}')
        create_report(filenames_for_prepare_report)
        logging.info(f'Отчет по {Path(filepath).stem} успешно создан')


if __name__ == '__main__':
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s-[%(levelname)s]- Module: %(module)s - Line:%(lineno)d - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(f'{__file__}.log', encoding='utf-8'),
        ]
    )
    counting_statistics_kpss_cnts()
