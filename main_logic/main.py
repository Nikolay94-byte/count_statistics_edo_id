from create_output_exels import convert_json_to_exel, find_column
from check_input_exel import check_exel
from create_report import create_report
from settings import INPUTFILEPATH
from utils import CheckError
from constants import INPUT_DATA_COLUMN_NAME, OUTPUT_DATA_COLUMN_NAME


def convert_json_to_exel_final() -> None:
    """Формирует три итоговых файлика - распознанные значения, верифицированные значения,
    а также отчет по качеству распознавания
    """

    check_result, fail_reason = check_exel(INPUTFILEPATH)

    if check_result:
        input_request_filename = convert_json_to_exel(find_column('document_input_request'), INPUT_DATA_COLUMN_NAME)
        verification_request_filename = convert_json_to_exel(find_column('document_verification_request'),
                                                             OUTPUT_DATA_COLUMN_NAME)
    else:
        raise CheckError(fail_reason)

    filenames_for_prepare_report = input_request_filename, verification_request_filename

    create_report(filenames_for_prepare_report)


if __name__ == '__main__':
    convert_json_to_exel_final()
