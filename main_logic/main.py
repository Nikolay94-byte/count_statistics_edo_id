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

    logging.info(f"Читаем файл с эталонными данными {etalon_file_path}")
    etalon_df = pd.read_excel(etalon_file_path, dtype=str).fillna("")

    # Проверяем, нужно ли разворачивать эталонные данные
    # Проверяем наличие колонок, которые должны быть после разворачивания
    has_expanded_structure = (
            constants.FILE_NAME in etalon_df.columns and
            constants.ATTRIBUTE_NAME in etalon_df.columns and
            constants.ATTRIBUTE_NAME_RUS in etalon_df.columns and
            constants.ETALON_VALUE in etalon_df.columns
    )

    if has_expanded_structure:
        logging.info("Эталонные данные уже имеют правильную структуру, используем как есть")
        expanded_etalon_df = etalon_df
    else:
        logging.info("Разворачиваем эталонные данные")
        expanded_etalon_df = expand_dataframe_data(etalon_df, constants.ETALON_VALUE)

    print(tabulate(expanded_etalon_df, headers='keys', tablefmt='grid', showindex=False))
    print("\n" + "=" * 80 + "\n")

    logging.info(f"Читаем файл с распознанными данными {recognized_file_path}")
    if os.path.splitext(recognized_file_path)[1].lower() == '.csv':
        recognized_df = decoding_csv(recognized_file_path)
    else:
        recognized_df = pd.read_excel(recognized_file_path, dtype=str).fillna("")

    # Проверяем, нужно ли парсить распознанные данные
    # Данные уже распарсены, если есть колонки из parse_bd_file результата
    needs_parsing = not all(col in recognized_df.columns for col in [
        constants.FILE_NAME,
        constants.ATTRIBUTE_NAME,
        constants.ATTRIBUTE_NAME_RUS,
        constants.RECOGINIZED_VALUE
    ])

    if needs_parsing:
        logging.info(f"Парсим файл с распознанными данными {recognized_file_path}")
        parsed_recognized_df = parse_bd_file(recognized_df)
    else:
        logging.info("Распознанные данные уже распаршены, используем как есть")
        parsed_recognized_df = recognized_df

    # Проверяем, нужно ли разворачивать распознанные данные
    # Проверяем, есть ли значения с разделителем ';' в колонке RECOGINIZED_VALUE
    needs_expanding = parsed_recognized_df[constants.RECOGINIZED_VALUE].astype(str).str.contains(';', na=False).any()

    if needs_expanding:
        logging.info("Разворачиваем распознанные данные")
        expanded_recognized_df = expand_dataframe_data(parsed_recognized_df, constants.RECOGINIZED_VALUE)
    else:
        logging.info("Распознанные данные уже развернуты, используем как есть")
        expanded_recognized_df = parsed_recognized_df

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