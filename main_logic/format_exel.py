import pandas as pd
import constants
from settings import DATA_PATH


def format_exel(value_column_name: str, new_book_name: str) -> str:
    """Форматирует эксель под необходимый формат (формат загрузки эталонки)."""
    df = pd.read_excel(DATA_PATH / new_book_name, index_col=0, dtype=str)
    # делаем анпивот таблицы
    unpivot_df = df.stack(dropna=False).reset_index()
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
    remove_unnecessary_atr_final_df.to_excel(DATA_PATH / new_book_name, index=False)
    return new_book_name
