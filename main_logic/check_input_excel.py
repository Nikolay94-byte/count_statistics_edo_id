import os.path

from utils.constants import INPUT_FILE_PATH
from utils.utils import open_excel
from utils import constants


def check_excel(filepath: str):
    """Проверяет, соответствует ли excel-файл необходимому формату:
    1 Есть ли необходимый файл с соответствующим названием - "Выгрузка с прода ФОИВ.xlsx"
    2 Есть ли 4 обязательные колонки:
    - reg_number
    - document_type
    - document_input_request
    - document_verification_request
    3 Есть ли ячейки, содержащие 32767 символов (32767 - максимально допустимое количество символов в ячейке excel,
    что означает, что json не полный, сломанный
    """

    if not os.path.exists(filepath):
       raise FileNotFoundError(
            f'Необходимый файл отсутствует, убедитесь, что файл назван верно и находится по пути {INPUT_FILE_PATH}'
        )

    sheet = open_excel(filepath)
    mandatory_column_name_list = [
        constants.REG_NUMBER, constants.DOCUMENT_TYPE, constants.DOCUMENT_INPUT_REQUEST,
        constants.DOCUMENT_VERIFICATION_REQUEST]
    current_column_name_list = [cell.value for cell in sheet['1'] if cell.value]
    if not set(mandatory_column_name_list).issubset(current_column_name_list):
        raise ValueError(
            f'В файле отсутствуют обязательные колонки, убедитесь что все 4 есть в файле '
            f'{constants.REG_NUMBER}, {constants.DOCUMENT_TYPE}, '
            f'{constants.DOCUMENT_INPUT_REQUEST}, {constants.DOCUMENT_VERIFICATION_REQUEST}'
        )

    for row in range(2, sheet.max_row+1):
        for cell in sheet[str(row)]:
            if len(str(cell.value)) == 32767:
                raise ValueError(
                    f'Ячейка (строка {row}, колонка {sheet[cell.coordinate].column}) содержит 32767 символов'
                )
