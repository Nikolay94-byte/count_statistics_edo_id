import logging
from pathlib import Path
from typing import List

import pandas as pd

from utils import constants
from utils.utils import find_file_by_name


def check_columns_in_file(file_path: Path, required_columns: List[str], file_type: str) -> None:
    """Проверяет наличие необходимых колонок в файле."""
    try:
        if file_path.suffix.lower() == '.csv':
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path, sheet_name=0)

        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            error_msg = f"В файле {file_type} отсутствуют обязательные колонки: {missing_columns}"
            logging.error(error_msg)
            logging.info(f"Доступные колонки в файле: {list(df.columns)}")
            raise ValueError(error_msg)

        logging.info(f"✓ Файл {file_type}: все необходимые колонки присутствуют")

    except Exception as e:
        error_msg = f"Ошибка при чтении файла {file_type}: {str(e)}"
        logging.error(error_msg)
        raise ValueError(error_msg)


def check_files(filepath: str):
    """
    1 Проверяет существование файлов по имени без учета расширения:
    - INPUT_DATA_ETALON
    - INPUT_DATA_RECOGINIZED
    2 Есть ли в файле INPUT_DATA_ETALON необходимые колонки
    """

    data_path = Path(filepath)

    # 1. Проверяем существование файлов по имени без учета расширения
    logging.info(f"Поиск файлов по имени в директории: {filepath}")

    etalon_file = find_file_by_name(data_path, constants.INPUT_DATA_ETALON_FILE)
    recognized_file = find_file_by_name(data_path, constants.INPUT_DATA_RECOGINIZED_FILE)

    available_files = [f.stem for f in data_path.iterdir() if f.is_file()]

    missing_files = []

    if etalon_file is None:
        missing_files.append(constants.INPUT_DATA_ETALON_FILE)

    if recognized_file is None:
        missing_files.append(constants.INPUT_DATA_RECOGINIZED_FILE)

    if missing_files:
        error_msg = f"Файлы не найдены: {', '.join(missing_files)}"
        logging.error(error_msg)
        logging.info(f"Доступные файлы в директории: {available_files}")
        raise FileNotFoundError(f"{error_msg} в директории {filepath}")

    logging.info(f"✓ Найден файл ETALON: {etalon_file.name}")
    logging.info(f"✓ Найден файл RECOGINIZED: {recognized_file.name}")

    # 2. Проверяем колонки в файле INPUT_DATA_ETALON
    check_columns_in_file(
        etalon_file,
        [constants.FILE_NAME, constants.ATTRIBUTE_NAME, constants.ATTRIBUTE_NAME_RUS, constants.ETALON_VALUE],
        constants.INPUT_DATA_ETALON_FILE
    )
    logging.info("✓ Все проверки пройдены успешно!")
