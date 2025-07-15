import os
import logging
import shutil
from pandas import DataFrame
import pandas as pd
import openpyxl
from . import constants


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
