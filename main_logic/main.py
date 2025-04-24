import logging
import sys
from utils.constants import INPUT_FILE_PATH, DOCUMENT_INPUT_REQUEST, DOCUMENT_VERIFICATION_REQUEST
from create_output_excels import convert_json_to_excel, find_column_index
from check_input_excel import check_excel
from create_report import create_report
from utils.constants import INPUT_DATA_COLUMN_NAME, OUTPUT_DATA_COLUMN_NAME


def counting_statistics_kpss_cnts() -> None:
    """Формирует три итоговых файлика - распознанные значения, верифицированные значения,
    а также отчет по качеству распознавания
    """
    logging.info(f'Начинаем сбор статистики')
    check_excel(INPUT_FILE_PATH)

    input_request_filename = convert_json_to_excel(find_column_index(DOCUMENT_INPUT_REQUEST), INPUT_DATA_COLUMN_NAME)
    verification_request_filename = convert_json_to_excel(find_column_index(DOCUMENT_VERIFICATION_REQUEST),
                                                             OUTPUT_DATA_COLUMN_NAME)

    filenames_for_prepare_report = input_request_filename, verification_request_filename

    create_report(filenames_for_prepare_report)
    logging.info(f'Отчет создан')


if __name__ == '__main__':
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - [%(levelname)s] - %(message)s -Line: %(lineno)d',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(f'{__file__}.log', encoding='utf-8'),
        ]
    )
    counting_statistics_kpss_cnts()
