import logging
import sys
from constants import INPUT_FILE_PATH, DOCUMENT_INPUT_REQUEST, DOCUMENT_VERIFICATION_REQUEST
from create_output_exels import convert_json_to_exel, find_column_index
from check_input_exel import check_exel
from create_report import create_report
from utils import CheckError
from constants import INPUT_DATA_COLUMN_NAME, OUTPUT_DATA_COLUMN_NAME


def counting_statistics_kpss_cnts() -> None:
    """Формирует три итоговых файлика - распознанные значения, верифицированные значения,
    а также отчет по качеству распознавания
    """
    logging.info(f'Начинаем сбор статистики')
    check_result, verdict = check_exel(INPUT_FILE_PATH)

    if check_result:
        input_request_filename = convert_json_to_exel(find_column_index(DOCUMENT_INPUT_REQUEST), INPUT_DATA_COLUMN_NAME)
        verification_request_filename = convert_json_to_exel(find_column_index(DOCUMENT_VERIFICATION_REQUEST),
                                                             OUTPUT_DATA_COLUMN_NAME)
    else:
        logging.critical(f'Входные данные не прошли проверку по причине: {verdict}')
        raise CheckError(verdict)

    logging.info(
        f'Подготовительные файлы {input_request_filename} и {verification_request_filename} успешно сформированы'
    )
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
