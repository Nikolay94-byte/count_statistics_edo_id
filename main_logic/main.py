import logging
import os
import sys
from pathlib import Path

import pandas as pd

from main_logic.check_input_excel import check_and_clean_file
from main_logic.create_report import create_report
from utils.constants import (
    APPLICATION,
    APPLICATION_FOR_THE_RECOVERY,
    COURT_ORDER,
    DOC_CLASS,
    INPUT_DATA_DIRECTORY_PATH,
    LABOUR_COMMISSION,
    MAIN_DOC,
    OUTPUT_INPUT_DATA_FORMAT_CSV_DIRECTORY_PATH,
    OUTPUT_INPUT_DATA_FORMAT_XLSX_APPLICATION_DIRECTORY_PATH,
    OUTPUT_INPUT_DATA_FORMAT_XLSX_DIRECTORY_PATH,
    OUTPUT_INPUT_DATA_FORMAT_XLSX_MAIN_DIRECTORY_PATH,
    OUTPUT_REPORTS_DIRECTORY_PATH,
    PERF_LIST,
)
from utils.utils import (
    convert_csv_to_excel_in_folder,
    show_in_logs_document_statistic,
)


def checks_sources_and_creates_reports(sources_directory: str) -> list:
    """Проверяет исходники и создает отчеты"""
    quality_percent_list = []
    for filepath in Path(sources_directory).glob('*.xlsx'):
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

    # Вычленяем заявление на взыскание
    for filepath in Path(OUTPUT_INPUT_DATA_FORMAT_XLSX_DIRECTORY_PATH).glob('*.xlsx'):
        input_data_df = pd.read_excel(filepath, index_col=None, dtype=str)
        doc_name = (input_data_df.loc[1, DOC_CLASS]).upper()
        if doc_name in [PERF_LIST, LABOUR_COMMISSION, COURT_ORDER]:
            os.makedirs(OUTPUT_INPUT_DATA_FORMAT_XLSX_MAIN_DIRECTORY_PATH, exist_ok=True)
            os.makedirs(OUTPUT_INPUT_DATA_FORMAT_XLSX_APPLICATION_DIRECTORY_PATH, exist_ok=True)
            input_data_df[DOC_CLASS] = APPLICATION_FOR_THE_RECOVERY
            application_for_the_recovery_file_name = f"{Path(filepath).stem}_{APPLICATION_FOR_THE_RECOVERY}.xlsx"
            with pd.ExcelWriter(OUTPUT_INPUT_DATA_FORMAT_XLSX_APPLICATION_DIRECTORY_PATH / application_for_the_recovery_file_name) as writer:
                input_data_df.to_excel(writer, sheet_name="Data", index=False)
            target_path = OUTPUT_INPUT_DATA_FORMAT_XLSX_MAIN_DIRECTORY_PATH / filepath.name
            if target_path.exists():
                os.remove(target_path)
            os.rename(filepath, target_path)
        else:
            break

    quality_percent_main_document = checks_sources_and_creates_reports(OUTPUT_INPUT_DATA_FORMAT_XLSX_MAIN_DIRECTORY_PATH)
    quality_percent_application = checks_sources_and_creates_reports(OUTPUT_INPUT_DATA_FORMAT_XLSX_APPLICATION_DIRECTORY_PATH)

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
