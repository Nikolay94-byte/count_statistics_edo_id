import datetime
import pandas as pd
from utils import constants


def calculate_metrics(etalon_df: pd.DataFrame, recognized_df: pd.DataFrame) -> pd.DataFrame:
    """Рассчитывает метрики качества распознавания на основе эталонных и распознанных значений"""

    def compare_cells(reference, recognized):
        """Сравнивает две ячейки и возвращает оценку 0 или 1"""
        ref_str = str(reference) if pd.notna(reference) else ""
        rec_str = str(recognized) if pd.notna(recognized) else ""

        ref_clean = ref_str.strip()
        rec_clean = rec_str.strip()

        if not ref_clean and not rec_clean:
            return 1
        elif not ref_clean or not rec_clean:
            return 0
        else:
            return 1 if ref_str == rec_str else 0

    def count_rows_for_attribute(df, file_name, attribute):
        """Считает количество строк для конкретного файла и атрибута"""
        return len(df[
                       (df[constants.FILE_NAME] == file_name) &
                       (df[constants.ATTRIBUTE_NAME] == attribute)
                       ])

    # Создаем полный набор всех возможных файлов
    all_files = set(etalon_df[constants.FILE_NAME]).union(set(recognized_df[constants.FILE_NAME]))

    # Берем только те атрибуты из константы, которые есть в эталонных данных
    etalon_attributes = set(etalon_df[constants.ATTRIBUTE_NAME])
    available_attributes = set(constants.TABLE_ATTRIBUTES.keys()) & etalon_attributes

    # Создаем полный DataFrame с максимальным количеством строк для каждой комбинации
    full_rows = []

    for file_name in all_files:
        for attribute in available_attributes:
            rus_name = constants.TABLE_ATTRIBUTES[attribute]

            # Определяем максимальное количество строк для этой комбинации
            etalon_count = count_rows_for_attribute(etalon_df, file_name, attribute)
            recognized_count = count_rows_for_attribute(recognized_df, file_name, attribute)
            max_rows = max(etalon_count, recognized_count)

            # Если есть хотя бы одна строка с данными, создаем строки
            if max_rows > 0:
                for row_num in range(max_rows):
                    full_rows.append({
                        constants.FILE_NAME: file_name,
                        constants.ATTRIBUTE_NAME: attribute,
                        constants.ATTRIBUTE_NAME_RUS: rus_name,
                        'row_num': row_num
                    })

    full_df = pd.DataFrame(full_rows)

    # Добавляем порядковый номер к исходным DataFrame
    etalon_df_with_num = etalon_df.copy()
    recognized_df_with_num = recognized_df.copy()

    for df in [etalon_df_with_num, recognized_df_with_num]:
        df['row_num'] = df.groupby([constants.FILE_NAME, constants.ATTRIBUTE_NAME]).cumcount()

    # Объединяем с полным DataFrame
    merge_columns = [constants.FILE_NAME, constants.ATTRIBUTE_NAME, 'row_num']

    merged_df = pd.merge(
        full_df,
        etalon_df_with_num[merge_columns + [constants.ETALON_VALUE]],
        on=merge_columns,
        how='left'
    )

    merged_df = pd.merge(
        merged_df,
        recognized_df_with_num[merge_columns + [constants.RECOGINIZED_VALUE]],
        on=merge_columns,
        how='left'
    )

    # Удаляем временную колонку и рассчитываем оценки
    merged_df = merged_df.drop('row_num', axis=1)
    merged_df[constants.CELL_SCORE] = merged_df.apply(
        lambda row: compare_cells(row[constants.ETALON_VALUE], row[constants.RECOGINIZED_VALUE]),
        axis=1
    )

    # 2. Оценка столбец - среднее по группам (файл + параметр)
    column_scores = merged_df.groupby(
        [constants.FILE_NAME, constants.ATTRIBUTE_NAME]
    )[constants.CELL_SCORE].mean().round(2)

    # Создаем колонку с оценками столбцов, но оставляем только в первой строке для каждого параметра
    merged_df[constants.COLUMN_SCORE] = merged_df.apply(
        lambda row: column_scores.get((row[constants.FILE_NAME], row[constants.ATTRIBUTE_NAME]), 0),
        axis=1
    )

    # Оставляем значение только в первой строке для каждого параметра
    mask_column = merged_df.duplicated(subset=[constants.FILE_NAME, constants.ATTRIBUTE_NAME], keep='first')
    merged_df.loc[mask_column, constants.COLUMN_SCORE] = None

    # 3. Оценка пакет - среднее по файлам (по уникальным колонкам)
    package_scores = merged_df.drop_duplicates(
        subset=[constants.FILE_NAME, constants.ATTRIBUTE_NAME]
    ).groupby(constants.FILE_NAME)[constants.COLUMN_SCORE].mean().round(2)

    # Создаем колонку с оценками пакетов, но оставляем только в первой строке для каждого файла
    merged_df[constants.PACKAGE_SCORE] = merged_df[constants.FILE_NAME].map(package_scores)

    # Оставляем значение только в первой строке для каждого файла
    mask_package = merged_df.duplicated(subset=[constants.FILE_NAME], keep='first')
    merged_df.loc[mask_package, constants.PACKAGE_SCORE] = None

    final_columns = [
        constants.FILE_NAME, constants.ATTRIBUTE_NAME, constants.ATTRIBUTE_NAME_RUS,
        constants.ETALON_VALUE, constants.RECOGINIZED_VALUE,
        constants.CELL_SCORE, constants.COLUMN_SCORE, constants.PACKAGE_SCORE
    ]

    return merged_df[final_columns]


def create_report(etalon_df: pd.DataFrame, recognized_df: pd.DataFrame, doc_type: str) -> pd.DataFrame:
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
            .dropna(subset=[constants.COLUMN_SCORE])
            .groupby(constants.ATTRIBUTE_NAME_RUS)[constants.COLUMN_SCORE]
            .mean()
            .mul(100)
            .round(2)
            .reset_index()
    )

    column_statistics_report_df = pd.DataFrame({
        constants.COLUMN_NAME: column_quality[constants.ATTRIBUTE_NAME_RUS],
        constants.COLUMN_QUALITY: column_quality[constants.COLUMN_SCORE]
    })
    column_statistics_report_df = column_statistics_report_df.sort_values(constants.COLUMN_QUALITY)

    # первый лист 'общая статистика'
    report_doc_type = doc_type
    count_date = datetime.date.today()
    amount_examples = paket_statistics_report_df[constants.FILE_NAME].nunique()
    quality_percent = paket_statistics_report_df[constants.PACKAGE_SCORE].mean() * 100
    quality_percent = round(quality_percent, 2)

    final_report_df = pd.DataFrame({
        constants.DOCUMENT_NAME: [report_doc_type],
        constants.COUNT_DATE: [count_date],
        constants.AMOUNT_EXAMPLES: [amount_examples],
        constants.QUALITY_PERCENT: [quality_percent],
    })

    report_file_name = f"{doc_type}_report.xlsx"

    with pd.ExcelWriter(constants.OUTPUT_REPORTS_DIRECTORY_PATH / report_file_name) as writer:
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
