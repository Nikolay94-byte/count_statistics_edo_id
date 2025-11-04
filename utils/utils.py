import json
import logging
from pathlib import Path
from typing import Optional

import pandas as pd

from utils import constants


def find_file_by_name(directory: Path, base_name: str) -> Optional[Path]:
    """Ищет файл в директории по базовому имени без учета расширения."""
    if not directory.exists():
        return None

    for file_path in directory.iterdir():
        if file_path.is_file() and file_path.stem == base_name:
            return file_path

    return None


def decoding_csv(recognized_file_path: Path) -> pd.DataFrame:
    """Чтение CSV."""

    ENCODINGS = ['utf-8', 'utf-8-sig', 'windows-1251', 'cp1251', 'iso-8859-5', 'cp866']
    recognized_df = None
    result_encoding = None

    for encoding in ENCODINGS:
        try:
            recognized_df = pd.read_csv(recognized_file_path, encoding=encoding)
            result_encoding = encoding
            break
        except UnicodeDecodeError:
            continue

    if recognized_df is None:
        error_msg = f"Не удалось прочитать файл {recognized_file_path} с доступными кодировками"
        logging.error(error_msg)
        raise ValueError(error_msg)

    logging.info(f"Файл {recognized_file_path} успешно прочитан (кодировка: {result_encoding})")
    logging.debug(f"Колонки: {recognized_df.columns.tolist()}, размер: {recognized_df.shape}")

    return recognized_df

def parse_bd_file(recognized_df: pd.DataFrame) -> pd.DataFrame:
    """Парсит данные из файла с данными из бд."""

    results = []

    for _, row in recognized_df.iterrows():
        filename = row['name']
        data_object = json.loads(row['data_object'])

        file_data = {}
        for column in data_object:
            for row_obj in column.get('objects', []):
                for attribute in row_obj.get('attributes', []):
                    attr_name = attribute['name']

                    if attr_name in constants.TABLE_ATTRIBUTES:
                        value = None
                        if attribute.get('value') and attribute['value'].get('text'):
                            value = attribute['value']['text']

                        if attr_name in file_data:
                            if file_data[attr_name] and value:
                                file_data[attr_name] = f"{file_data[attr_name]};{value}"
                            elif value:
                                file_data[attr_name] = value
                        else:
                            file_data[attr_name] = value

        for attr_name in constants.TABLE_ATTRIBUTES.keys():
            results.append({
                constants.FILE_NAME: filename,
                constants.ATTRIBUTE_NAME: attr_name,
                constants.ATTRIBUTE_NAME_RUS: constants.TABLE_ATTRIBUTES[attr_name],
                constants.RECOGINIZED_VALUE: file_data.get(attr_name, None)
            })

    df_result = pd.DataFrame(results)

    parsed_recognized_df = df_result.sort_values(
        [constants.FILE_NAME, constants.ATTRIBUTE_NAME]
    ).reset_index(drop=True)

    return parsed_recognized_df


def expand_dataframe_data(df: pd.DataFrame, column_name: str) -> pd.DataFrame:
    """Разворачивает(переносит на новую строку) значения ячеек, где перечисленно больше одного инстанса колонки."""

    expanded_df = df.copy()
    expanded_df[column_name] = expanded_df[column_name].apply(
        lambda x: str(x).split(';') if pd.notna(x) and ';' in str(x) else [x]
    )
    expanded_df = expanded_df.explode(column_name, ignore_index=True)
    expanded_df[column_name] = expanded_df[column_name].apply(
        lambda x: None if pd.isna(x) or str(x).strip() == 'None' else str(x).strip()
    )
    expanded_df = expanded_df.sort_values([constants.FILE_NAME, constants.ATTRIBUTE_NAME]).reset_index(
        drop=True)
    return expanded_df
