import os
import logging
import shutil

import pandas as pd
import openpyxl


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
