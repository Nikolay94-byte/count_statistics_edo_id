import os
import shutil
from pathlib import Path

import pandas as pd
import openpyxl


def convert_csv_to_excel_in_folder(
        input_folder: str | Path,
        output_csv_folder: str | Path,
        output_excel_folder: str | Path,
) -> None:
    """
    Конвертирует все CSV-файлы из input_folder в XLSX (в output_excel_folder)
    и копирует исходные CSV в output_csv_folder.
    """
    # Создаём папки, если их нет
    os.makedirs(output_excel_folder, exist_ok=True)
    os.makedirs(output_csv_folder, exist_ok=True)

    for filename in os.listdir(input_folder):
        if filename.endswith('.csv'):
            csv_path = os.path.join(input_folder, filename)

            # 1. Копируем CSV в OUTPUT_INPUT_DATA_FORMAT_CSV
            output_csv_path = os.path.join(output_csv_folder, filename)
            shutil.copy2(csv_path, output_csv_path)
            print(f"[CSV] Скопирован: {filename} → {output_csv_folder}")

            # 2. Конвертируем в Excel
            excel_filename = filename.replace('.csv', '.xlsx')
            excel_path = os.path.join(output_excel_folder, excel_filename)

            try:
                df = pd.read_csv(csv_path, delimiter=',')
                df.to_excel(excel_path, index=False, engine='openpyxl')
                print(f"[XLSX] Успешно преобразован: {filename} → {excel_filename}")
            except Exception as e:
                print(f"Ошибка при обработке {filename}: {e}")


def open_excel(filepath: str) -> openpyxl.worksheet.worksheet.Worksheet:
    """Открывает файл, создает рабочую книгу для работы с данными"""
    book = openpyxl.open(filepath)
    sheet = book.active
    return sheet


def convert_file_attributes_to_dict(file_path) -> dict[str, str]:
    """Возвращает словарь атрибутов, необходимый для подсчета статистики."""
    result_dict = {}

    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                if not line or line.startswith('#'):
                    continue
                if ':' in line:
                    key_part, value_part = line.split(':', 1)
                    result_dict[key_part] = value_part

    except FileNotFoundError:
        raise FileNotFoundError(f"Ошибка: Файл '{file_path}' не найден.")
    except Exception as error:
        raise Exception(f"Произошла ошибка при чтении файла: {error}")

    return result_dict


def convert_file_products_to_dict(file_path) -> dict[str, list[str | int]]:
    """Возвращает словарь продуктов запроса"""
    result_dict = {}

    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                line = line.strip()
                if not line or line.startswith('#'):
                    continue

                key, value = line.split(':', 1)
                key = key.strip()
                value = value.strip()

                items = value.split(',')
                processed_items = []
                for item in items:
                    item = item.strip()
                    if item.isdigit():
                        processed_items.append(int(item))
                    else:
                        processed_items.append(item)

                result_dict[key] = processed_items

    except FileNotFoundError:
        raise FileNotFoundError(f"Ошибка: Файл '{file_path}' не найден.")
    except Exception as error:
        raise Exception(f"Произошла ошибка при чтении файла: {error}")

    return result_dict
