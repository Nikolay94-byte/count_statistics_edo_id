import pandas as pd
from pandas import DataFrame

from utils import constants


def format_excel(value_column_name: str, dataframe_for_formating: DataFrame) -> DataFrame:
    """Форматирует эксель под необходимый формат (формат загрузки эталонки)."""
    # делаем анпивот таблицы
    unpivot_df = (
        pd.melt(
            dataframe_for_formating,
            id_vars=[dataframe_for_formating.columns[0]],
            var_name='attribute',
            value_name='value'
        )
        .sort_values(by=[dataframe_for_formating.columns[0], 'attribute'])  # Сортировка по двум столбцам
        .reset_index(drop=True)
    )
    unpivot_df.columns = [constants.FILE_NAME_COLUMN_NAME, constants.COMMON_ATTRIBUTE_NAME_COLUMN_NAME,
                          value_column_name]
    # разделяем колонку Общее наим.атрибута на две
    divide_atribute_name_df = unpivot_df[constants.COMMON_ATTRIBUTE_NAME_COLUMN_NAME].str.split('#', expand=True)
    divide_atribute_name_df.columns = [constants.SYSTEM_ATTRIBUTE_NAME_COLUNM_NAME,
                                       constants.ATTRIBUTE_NAME_COLUNM_NAME]
    # Объединяем датафреймы и расставляем колонки в нужном порядке
    final_df = pd.concat([unpivot_df, divide_atribute_name_df], axis=1)
    final_df = final_df.drop(constants.COMMON_ATTRIBUTE_NAME_COLUMN_NAME, axis=1)
    final_df = final_df.reindex(columns=[constants.FILE_NAME_COLUMN_NAME, constants.SYSTEM_ATTRIBUTE_NAME_COLUNM_NAME,
                                         constants.ATTRIBUTE_NAME_COLUNM_NAME, value_column_name])
    # Удаляем строки с не нужными атрибутами (по которым не считается статистика)
    remove_unnecessary_atr_final_df = final_df.dropna(subset=[constants.ATTRIBUTE_NAME_COLUNM_NAME])
    return remove_unnecessary_atr_final_df
