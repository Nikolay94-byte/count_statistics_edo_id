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
    Конвертирует CSV в XLSX без пропуска битых строк.
    Определяет кодировку перебором стандартных вариантов.
    """
    os.makedirs(output_excel_folder, exist_ok=True)
    os.makedirs(output_csv_folder, exist_ok=True)

    # Приоритетные кодировки для кириллицы (можно добавить другие)
    ENCODINGS = ['utf-8', 'utf-8-sig', 'windows-1251', 'cp1251', 'iso-8859-5', 'cp866']

    for filename in os.listdir(input_folder):
        if not filename.endswith('.csv'):
            # Копируем не-CSV файлы без изменений
            shutil.copy2(
                os.path.join(input_folder, filename),
                output_excel_folder
            )
            logging.info(f"[SKIP] Не CSV: {filename} → {output_excel_folder}")
            continue

        file_path = os.path.join(input_folder, filename)
        csv_copy_path = os.path.join(output_csv_folder, filename)
        excel_path = os.path.join(
            output_excel_folder,
            filename.replace('.csv', '.xlsx')
        )

        # 1. Копируем оригинальный CSV
        shutil.copy2(file_path, csv_copy_path)
        logging.info(f"[COPY] CSV сохранён: {filename} → {output_csv_folder}")

        # 2. Пытаемся конвертировать с разными кодировками
        success = False
        last_error = None

        for encoding in ENCODINGS:
            try:
                # Чтение без пропуска ошибок (on_bad_lines=None)
                df = pd.read_csv(
                    file_path,
                    encoding=encoding,
                    delimiter=',',
                    engine='python',
                    quotechar='"',
                    on_bad_lines=None  # Не пропускать битые строки!
                )

                # Сохранение в Excel
                df.to_excel(
                    excel_path,
                    index=False,
                    engine='openpyxl'
                )
                logging.info(f"[SUCCESS] Конвертирован ({encoding}): {filename}")
                success = True
                break

            except UnicodeDecodeError:
                continue
            except pd.errors.ParserError as e:
                last_error = f"Ошибка формата ({encoding}): {str(e)}"
                continue
            except Exception as e:
                last_error = f"Неизвестная ошибка ({encoding}): {str(e)}"
                continue

        if not success:
            error_msg = f"[FAIL] Не удалось конвертировать {filename}. Последняя ошибка: {last_error}"
            logging.error(error_msg)
            raise ValueError(error_msg)


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
        #logging.info(f'[{doc_type}] Выборочная дисперсия {statistics.variance(quality_percent_list)}')
        logging.info(f'[{doc_type}] Стандартное отклонение {statistics.pstdev(quality_percent_list)}')
        logging.info(f'[{doc_type}] Размах {max(quality_percent_list) - min(quality_percent_list)}')

    logging.info(
        f"[{doc_type}] Среднее качество извлечения: {statistics.median(quality_percent_list)}"
    )
