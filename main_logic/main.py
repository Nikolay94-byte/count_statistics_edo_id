import logging
import os
import statistics
import sys
from pathlib import Path

from check_input_excel import check_and_clean_file

from create_report import create_report
from utils.constants import (
    INPUT_DATA_DIRECTORY_PATH,
    OUTPUT_INPUT_DATA_FORMAT_CSV_DIRECTORY_PATH,
    OUTPUT_INPUT_DATA_FORMAT_XLSX_DIRECTORY_PATH,
    OUTPUT_REPORTS_DIRECTORY_PATH,
)
from utils.utils import convert_csv_to_excel_in_folder


def counting_statistics_edo_id() -> None:
    """Считает статистику качество извлечения по выгрузке с прода"""
    quality_percent = []

    logging.info(
        f'Преобразуем исходные файлы из формата .csv в формат .xlsx, а также копируем исходники в итоговую папку OUTPUT'
    )
    convert_csv_to_excel_in_folder(
        INPUT_DATA_DIRECTORY_PATH,
        OUTPUT_INPUT_DATA_FORMAT_CSV_DIRECTORY_PATH,
        OUTPUT_INPUT_DATA_FORMAT_XLSX_DIRECTORY_PATH
    )

    os.makedirs(OUTPUT_REPORTS_DIRECTORY_PATH, exist_ok=True)

    for filepath in Path(OUTPUT_INPUT_DATA_FORMAT_XLSX_DIRECTORY_PATH).glob('*.xlsx'):
        logging.info(f'Начинаем проверку и редактирование исходника {Path(filepath).stem}')
        check_and_clean_file(filepath)

        logging.info(f'Создаем отчет по {Path(filepath).stem}')
        quality = create_report(filepath)
        quality_percent.append(quality)
        logging.info(f'Отчет по {Path(filepath).stem} успешно создан')

    logging.info(quality_percent)
    if len(quality_percent) > 1:
        logging.info(f'Выборочная дисперсия {statistics.variance(quality_percent)}')
        logging.info(f'Стандартное отклонение {statistics.pstdev(quality_percent)}')
        logging.info(f'Размах {max(quality_percent) - min(quality_percent)}')

    logging.info(
        f"Среднее качество извлечения: {statistics.median(quality_percent)}"
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
    counting_statistics_edo_id()
