import pandas as pd
import datetime
import constants
from settings import DATA_PATH


def create_report(filenames_for_prepare_report: tuple):
    """Создает отчет по качеству распознавания."""
    input_request_filename, verification_request_filename = filenames_for_prepare_report

    input_request_df = pd.read_excel(DATA_PATH / input_request_filename, index_col=None, dtype=str)
    verification_request_df = pd.read_excel(DATA_PATH / verification_request_filename, index_col=None, dtype=str)

    # пятый лист - "ст-ка по всем атр."
    general_report_df = pd.merge(
        verification_request_df, input_request_df,
        left_on=[constants.FILE_NAME_COLUMN_NAME, constants.SYSTEM_ATTRIBUTE_NAME_COLUNM_NAME,
                 constants.ATTRIBUTE_NAME_COLUNM_NAME],
        right_on=[constants.FILE_NAME_COLUMN_NAME, constants.SYSTEM_ATTRIBUTE_NAME_COLUNM_NAME,
                  constants.ATTRIBUTE_NAME_COLUNM_NAME]
    )
    general_report_df = general_report_df.fillna('')
    general_report_df[constants.COMPARISON_COLUMN_NAME] = \
        general_report_df[constants.OUTPUT_DATA_COLUMN_NAME].str.replace(' ', '') \
        == general_report_df[constants.INPUT_DATA_COLUMN_NAME].str.replace(' ', '')

    # второй лист - "только по зап-ым атр." (статистика по атрибутам, которые были заполнены верификаторами)
    only_completed_report_df = general_report_df.copy()
    only_completed_report_df_cleaned = \
        only_completed_report_df[only_completed_report_df[constants.OUTPUT_DATA_COLUMN_NAME] != ""]

    # третий лист - "детализация поатрибутивно" (отражает в каких атрибутах больше всего ошибок)
    attribute_statistics_report_df = only_completed_report_df_cleaned.copy()
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

    # четвертый лист "ложно извлеч атр." (отражает статистику по ложно извлеченным атрибутам, которые верификатор
    # удалил полностью)
    falsely_completed_report_df = general_report_df.copy()
    falsely_completed_report_df = falsely_completed_report_df[falsely_completed_report_df
                                                              [constants.OUTPUT_DATA_COLUMN_NAME] == ""]
    falsely_completed_report_series = \
        falsely_completed_report_df[falsely_completed_report_df[constants.COMPARISON_COLUMN_NAME] == False] \
            .groupby(constants.ATTRIBUTE_NAME_COLUNM_NAME).size().sort_values(ascending=False)
    falsely_completed_report_df_counted = falsely_completed_report_series.reset_index \
        (name=constants.FALSELY_COMPLETED_AMOUNT_ATTRIBUTE_COLUMN_NAME)

    # первый лист "общая статистика"
    document_name = input_request_filename.split("_document_")[0]
    count_date = datetime.date.today()
    amount_examples = general_report_df[constants.FILE_NAME_COLUMN_NAME].nunique()
    true_attribute_amount = (only_completed_report_df_cleaned[constants.COMPARISON_COLUMN_NAME] == True).sum()
    false_attribute_amount = (only_completed_report_df_cleaned[constants.COMPARISON_COLUMN_NAME] == False).sum()
    quality_percent = round(true_attribute_amount * 100/(true_attribute_amount + false_attribute_amount), 1)
    falsely_completed_average_amount = \
        round(falsely_completed_report_df_counted[constants.FALSELY_COMPLETED_AMOUNT_ATTRIBUTE_COLUMN_NAME]
              .sum()/amount_examples, 1)
    final_report_df = pd.DataFrame({
        constants.DOCUMENT_NAME: [document_name],
        constants.COUNT_DATE: [count_date],
        constants.AMOUNT_EXAMPLES: [amount_examples],
        constants.TRUE_ATTRIBUTE_AMOUNT: [true_attribute_amount],
        constants.FALSE_ATTRIBUTE_AMOUNT: [false_attribute_amount],
        constants.QUALITY_PERCENT: [quality_percent],
        constants.FALSELY_COMPLETED_AVERAGE_AMOUNT: [falsely_completed_average_amount],
    })

    report_file_name = document_name + '_' + 'report.xlsx'

    with pd.ExcelWriter(DATA_PATH / report_file_name) as writer:
        final_report_df.to_excel\
            (writer, sheet_name=constants.FINAL_REPORT_SHEET_NAME, index=False)
        attribute_statistics_report_df_counted.to_excel\
            (writer, sheet_name=constants.ATTRIBUTE_STATISTICS_REPORT_SHEET_NAME, index=False)
        falsely_completed_report_df_counted.to_excel\
            (writer, sheet_name=constants.FALSELY_COMPLETED_REPORT_SHEET_NAME, index=False)
        only_completed_report_df_cleaned.to_excel\
            (writer, sheet_name=constants.ONLY_COMPLETED_REPORT_SHEET_NAME, index=False)
        general_report_df.to_excel\
            (writer, sheet_name=constants.GENERAL_REPORT_SHEET_NAME, index=False)
