import datetime
from pathlib import Path

import pandas as pd

from utils import constants
from utils.constants import (
    CLASS_ATTRIBUTE_MAPPING,
    OUTPUT_REPORTS_DIRECTORY_PATH,
)
from utils.utils import normalize_dataframe


def create_report(filepath: str) -> float:
    """Создает отчет по качеству распознавания."""

    # четвертый лист - 'исходые данные'
    input_data_df = pd.read_excel(filepath, index_col=None, dtype=str)

    # второй лист - 'детализация попакетно'
    paket_statistics_report_df = input_data_df.copy()
    # определяем класс документа
    doc_name = (paket_statistics_report_df.loc[1, constants.DOC_CLASS]).upper()
    # оставляем только необходимые колонки и переименовывем их
    columns_to_keep = [constants.REGNUMBER, constants.ATTRIBUTE_NAME, constants.RUS_ATTRIBUTE_NAME,
                       constants.TEXT_NORMALIZED, constants.TEXT_VERIFICATION]
    new_columns_name = [constants.FILE_NAME_COLUMN_NAME, constants.SYSTEM_ATTRIBUTE_NAME_COLUNM_NAME,
                        constants.ATTRIBUTE_NAME_COLUNM_NAME, constants.INPUT_DATA_COLUMN_NAME,
                        constants.OUTPUT_DATA_COLUMN_NAME]
    paket_statistics_report_df = paket_statistics_report_df[columns_to_keep]
    paket_statistics_report_df = (paket_statistics_report_df[columns_to_keep].set_axis(new_columns_name, axis=1))
    # оставляем необходимые атрибуты согласно классу документа
    doc_type = CLASS_ATTRIBUTE_MAPPING.get(doc_name, doc_name)
    paket_statistics_report_df = normalize_dataframe(doc_type, paket_statistics_report_df)
    # производим подсчет
    paket_statistics_report_df = paket_statistics_report_df.fillna('')
    paket_statistics_report_df[constants.COMPARISON_COLUMN_NAME] = \
        paket_statistics_report_df[constants.OUTPUT_DATA_COLUMN_NAME].str.replace(' ', '') \
        == paket_statistics_report_df[constants.INPUT_DATA_COLUMN_NAME].str.replace(' ', '')

    # третий лист - 'детализация поатрибутивно' (отражает в каких атрибутах больше всего ошибок)
    attribute_statistics_report_df = paket_statistics_report_df.copy()
    attribute_statistics_report_df = attribute_statistics_report_df.drop\
        ([constants.FILE_NAME_COLUMN_NAME,
          constants.SYSTEM_ATTRIBUTE_NAME_COLUNM_NAME,
          constants.INPUT_DATA_COLUMN_NAME,
          constants.OUTPUT_DATA_COLUMN_NAME,
          ], axis=1)
    attribute_statistics_report_series = \
        attribute_statistics_report_df[attribute_statistics_report_df[constants.COMPARISON_COLUMN_NAME] == False]\
            .groupby(constants.ATTRIBUTE_NAME_COLUNM_NAME).size().sort_values(ascending=False)
    attribute_statistics_report_df_counted = attribute_statistics_report_series.reset_index\
        (name=constants.FALSE_ATTRIBUTE_AMOUNT)

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
            (writer, sheet_name=constants.PAKET_REPORT_SHEET_NAME, index=False)
        attribute_statistics_report_df_counted.to_excel\
            (writer, sheet_name=constants.ATTRIBUTE_STATISTICS_REPORT_SHEET_NAME, index=False)
        input_data_df.to_excel\
            (writer, sheet_name=constants.INPUT_DATA_SHEET_NAME, index=False)

    return quality_percent
