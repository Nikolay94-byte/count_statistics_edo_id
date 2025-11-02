import pandas as pd
import logging
import re
import datetime
from pathlib import Path
import pandas as pd

from utils import constants
from utils.constants import (
    OUTPUT_REPORTS_DIRECTORY_PATH,
)


def calculate_metrics(etalon_df: pd.DataFrame, recognized_df: pd.DataFrame) -> pd.DataFrame:
    """Рассчитывает метрики качества распознавания на основе эталонных и распознанных значений"""

    merged_df = pd.merge(
        etalon_df,
        recognized_df,
        on=[constants.FILE_NAME, constants.ATTRIBUTE_NAME, constants.ATTRIBUTE_NAME_RUS]
    )

    # 1. Оценка ячейка - сравнение значений
    def compare_cells(reference, recognized):
        if pd.isna(reference) and pd.isna(recognized):
            return 1
        elif pd.isna(reference) or pd.isna(recognized):
            return 0
        else:
            ref_str = str(reference).strip().lower()
            rec_str = str(recognized).strip().lower()
            return 1 if ref_str == rec_str else 0

    merged_df[constants.CELL_SCORE] = merged_df.apply(
        lambda row: compare_cells(row[constants.ETALON_VALUE], row[constants.RECOGINIZED_VALUE]),
        axis=1
    )

    # 2. Оценка столбец - среднее по группам (файл + параметр)
    merged_df[constants.COLUMN_SCORE] = merged_df.groupby(
        [constants.FILE_NAME, constants.ATTRIBUTE_NAME]
    )[constants.CELL_SCORE].transform('mean')

    # 3. Оценка пакет - среднее по файлам
    package_scores = merged_df.groupby(constants.FILE_NAME)[constants.COLUMN_SCORE].mean()
    merged_df[constants.PACKAGE_SCORE] = merged_df[constants.FILE_NAME].map(package_scores)

    final_columns = [
        constants.FILE_NAME, constants.ATTRIBUTE_NAME, constants.ATTRIBUTE_NAME_RUS,
        constants.ETALON_VALUE, constants.RECOGINIZED_VALUE,
        constants.CELL_SCORE, constants.COLUMN_SCORE, constants.PACKAGE_SCORE
    ]

    return merged_df[final_columns]


def create_report(etalon_df: pd.DataFrame, recognized_df: pd.DataFrame) -> pd.DataFrame:
    """Создает отчет по качеству распознавания."""

    # четвертый лист - 'эталонные значения'
    report_etalon_df = etalon_df

    # пятый лист - 'распознанные значения'
    report_recognized_df = recognized_df

    # второй лист - 'детализация попакетно'
    paket_statistics_report_df = calculate_metrics(report_etalon_df, report_recognized_df)

    # третий лист - 'детализация по колонкам' (отражает в каких колонках больше всего ошибок)
    column_quality = (
        paket_statistics_report_df
            .groupby(constants.ATTRIBUTE_NAME_RUS)[constants.CELL_SCORE]
            .mean()
            .mul(100)
            .round(2)
            .reset_index()
    )
    column_statistics_report_df = pd.DataFrame({
        constants.COLUMN_NAME: column_quality[constants.ATTRIBUTE_NAME_RUS],
        constants.COLUMN_QUALITY: column_quality[constants.CELL_SCORE]
    })
    column_statistics_report_df = column_statistics_report_df.sort_values(constants.COLUMN_QUALITY)

    # первый лист 'общая статистика'
    document_name = f"{Path(filepath).stem}"
    count_date = datetime.date.today()
    amount_examples = paket_statistics_report_df[constants.FILE_NAME_COLUMN_NAME].nunique()
    true_attribute_amount = (paket_statistics_report_df[constants.COMPARISON_COLUMN_NAME] == True).sum()
    false_attribute_amount = (paket_statistics_report_df[constants.COMPARISON_COLUMN_NAME] == False).sum()
    quality_percent = round(true_attribute_amount * 100/(true_attribute_amount + false_attribute_amount), 1)
    final_report_df = pd.DataFrame({
        constants.DOCUMENT_NAME: [document_name],
        constants.COUNT_DATE: [count_date],
        constants.AMOUNT_EXAMPLES: [amount_examples],
        constants.TRUE_ATTRIBUTE_AMOUNT: [true_attribute_amount],
        constants.FALSE_ATTRIBUTE_AMOUNT: [false_attribute_amount],
        constants.QUALITY_PERCENT: [quality_percent],
    })

    report_file_name = f"{document_name}_report.xlsx"

    with pd.ExcelWriter(OUTPUT_REPORTS_DIRECTORY_PATH / report_file_name) as writer:
        final_report_df.to_excel\
            (writer, sheet_name=constants.FINAL_REPORT_SHEET_NAME, index=False)
        paket_statistics_report_df.to_excel\
            (writer, sheet_name=constants.PAKET_STATISTICS_REPORT_SHEET_NAME, index=False)
        column_statistics_report_df.to_excel\
            (writer, sheet_name=constants.COLUMN_STATISTICS_REPORT_SHEET_NAME, index=False)
        report_etalon_df.to_excel\
            (writer, sheet_name=constants.ETALON_DATA_SHEET_NAME, index=False)
        report_recognized_df.to_excel\
            (writer, sheet_name=constants.RECOGINIZED_DATA_SHEET_NAME, index=False)

    return quality_percent
