import logging
import os
import shutil
import statistics
from pathlib import Path

import openpyxl
import pandas as pd
from pandas import DataFrame

from utils import constants


def convert_csv_to_excel_in_folder(
        input_folder: str,
        output_csv_folder: str,
        output_excel_folder: str,
) -> None:
    """
    Конвертирует все CSV-файлы в XLSX и копирует исходные CSV.
    """
    # Создаём папки, если их нет
    os.makedirs(output_excel_folder, exist_ok=True)
    os.makedirs(output_csv_folder, exist_ok=True)

    for filename in os.listdir(input_folder):
        file_path = os.path.join(input_folder, filename)

        if filename.endswith('.csv'):
            # Обработка CSV
            # 1. Копируем оригинал в output_csv_folder
            csv_copy_path = os.path.join(output_csv_folder, filename)
            shutil.copy2(file_path, csv_copy_path)
            logging.info(f"[CSV] Скопирован оригинал: {filename} → {output_csv_folder}")

            # 2. Конвертируем в XLSX и сохраняем в output_excel_folder
            excel_filename = filename.replace('.csv', '.xlsx')
            excel_path = os.path.join(output_excel_folder, excel_filename)
            try:
                df = pd.read_csv(file_path, delimiter=',')
                df.to_excel(excel_path, index=False, engine='openpyxl')
                logging.info(f"[XLSX] Успешно преобразован: {filename} → {excel_filename}")
            except Exception as e:
                logging.error(f"Ошибка конвертации {filename}: {e}")
        else:
            # Копируем НЕ-CSV файлы в output_excel_folder
            shutil.copy2(file_path, output_excel_folder)
            logging.info(f"Конвертация в .xlsx не нужна! Файл {filename} скопирован → {output_excel_folder}")


def open_excel(filepath: str) -> openpyxl.worksheet.worksheet.Worksheet:
    """Открывает файл, создает рабочую книгу для работы с данными"""
    book = openpyxl.open(filepath)
    sheet = book.active
    return sheet


def normalize_dataframe(doc_attributes: dict, dataframe_for_formating: DataFrame) -> DataFrame:
    """Оставляет необходимые атрибуты, переименовывает их русские названия"""
    # оставляем только необходимые атрибуты
    new_dataframe = dataframe_for_formating[
        dataframe_for_formating[constants.SYSTEM_ATTRIBUTE_NAME_COLUNM_NAME].isin(doc_attributes.keys())
    ].copy()
    # переименовываем значения в колонке Наим.атрибута
    new_dataframe[constants.ATTRIBUTE_NAME_COLUNM_NAME] = new_dataframe[
        constants.SYSTEM_ATTRIBUTE_NAME_COLUNM_NAME
    ].map(doc_attributes)
    return new_dataframe


def prepare_input_data(input_data_directory: str) -> None:
    """Создает копию исходного файла, заменяя класс на Заявление на взыскание. Распределяет исходники по папкам"""
    for filepath in Path(input_data_directory).glob('*.xlsx'):
        input_data_df = pd.read_excel(filepath, index_col=None, dtype=str)
        doc_name = (input_data_df.loc[1, constants.DOC_CLASS]).upper()
        if doc_name in [constants.PERF_LIST, constants.LABOUR_COMMISSION, constants.COURT_ORDER]:
            os.makedirs(constants.OUTPUT_INPUT_DATA_FORMAT_XLSX_MAIN_DIRECTORY_PATH, exist_ok=True)
            os.makedirs(constants.OUTPUT_INPUT_DATA_FORMAT_XLSX_APPLICATION_DIRECTORY_PATH, exist_ok=True)
            input_data_df[constants.DOC_CLASS] = constants.APPLICATION_FOR_THE_RECOVERY
            application_for_the_recovery_file_name = \
                f"{Path(filepath).stem}_{constants.APPLICATION_FOR_THE_RECOVERY}.xlsx"
            with pd.ExcelWriter(
                    constants.OUTPUT_INPUT_DATA_FORMAT_XLSX_APPLICATION_DIRECTORY_PATH
                    / application_for_the_recovery_file_name
            ) as writer:
                input_data_df.to_excel(writer, sheet_name="Data", index=False)
            target_path = constants.OUTPUT_INPUT_DATA_FORMAT_XLSX_MAIN_DIRECTORY_PATH / filepath.name
            if target_path.exists():
                os.remove(target_path)
            os.rename(filepath, target_path)
        else:
            break


def show_in_logs_document_statistic(doc_type: str, quality_percent_list: list) -> None:
    """Выводит логи статистики в консоль"""
    logging.info(f'[{doc_type}] Список показателей качества извлечения: {quality_percent_list}')
    if len(quality_percent_list) > 1:
        logging.info(f'[{doc_type}] Выборочная дисперсия {statistics.variance(quality_percent_list)}')
        logging.info(f'[{doc_type}] Стандартное отклонение {statistics.pstdev(quality_percent_list)}')
        logging.info(f'[{doc_type}] Размах {max(quality_percent_list) - min(quality_percent_list)}')

    logging.info(
        f"[{doc_type}] Среднее качество извлечения: {statistics.median(quality_percent_list)}"
    )
