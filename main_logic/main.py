import logging
import os
import sys
from pathlib import Path

from main_logic.check_input_excel import check_and_clean_file
from main_logic.create_report import create_report
from utils.constants import (
    APPLICATION,
    INPUT_DATA_DIRECTORY_PATH,
    MAIN_DOC,
    OUTPUT_INPUT_DATA_FORMAT_CSV_DIRECTORY_PATH,
    OUTPUT_INPUT_DATA_FORMAT_XLSX_APPLICATION_DIRECTORY_PATH,
    OUTPUT_INPUT_DATA_FORMAT_XLSX_DIRECTORY_PATH,
    OUTPUT_INPUT_DATA_FORMAT_XLSX_MAIN_DIRECTORY_PATH,
    OUTPUT_REPORTS_DIRECTORY_PATH,
)
from utils.utils import (
    convert_csv_to_excel_in_folder,
    show_in_logs_document_statistic, prepare_input_data
)


def check_input_data_and_create_reports(input_data_directory: str) -> list:
    """Проверяет исходники и создает отчеты"""
    quality_percent_list = []
    for filepath in Path(input_data_directory).glob('*.xlsx'):
        logging.info(f'Начинаем проверку и редактирование исходника {Path(filepath).stem}')
        check_and_clean_file(filepath)
        logging.info(f'Создаем отчет по {Path(filepath).stem}')
        quality = create_report(filepath)
        quality_percent_list.append(quality)
        logging.info(f'Отчет по {Path(filepath).stem} успешно создан')
    return quality_percent_list


def counting_statistics_edo_id() -> None:
    """Считает статистику качество извлечения по выгрузке с прода"""
    logging.info(
        f'Преобразуем исходные файлы из формата .csv в формат .xlsx, а также копируем исходники в итоговую папку OUTPUT'
    )
    convert_csv_to_excel_in_folder(
        INPUT_DATA_DIRECTORY_PATH,
        OUTPUT_INPUT_DATA_FORMAT_CSV_DIRECTORY_PATH,
        OUTPUT_INPUT_DATA_FORMAT_XLSX_DIRECTORY_PATH
    )

    os.makedirs(OUTPUT_REPORTS_DIRECTORY_PATH, exist_ok=True)

    prepare_input_data(OUTPUT_INPUT_DATA_FORMAT_XLSX_DIRECTORY_PATH)

    quality_percent_main_document = check_input_data_and_create_reports(
        OUTPUT_INPUT_DATA_FORMAT_XLSX_MAIN_DIRECTORY_PATH
    )
    quality_percent_application = check_input_data_and_create_reports(
        OUTPUT_INPUT_DATA_FORMAT_XLSX_APPLICATION_DIRECTORY_PATH
    )

    show_in_logs_document_statistic(MAIN_DOC, quality_percent_main_document)
    show_in_logs_document_statistic(APPLICATION, quality_percent_application)


if __name__ == '__main__':
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s-[%(levelname)s]- Module: %(module)s - Line:%(lineno)d - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(f'{__file__}.log', encoding='utf-8'),
        ]
    )
    counting_statistics_edo_id()
