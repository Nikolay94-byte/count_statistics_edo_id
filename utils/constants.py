from settings import DATA_PATH, ROOT_PATH


# INPUT, путь к папке с исходными файлами в формате .csv
INPUT_DATA_DIRECTORY_PATH = DATA_PATH / "INPUT"

# OUTPUT, пути к папкам с исходящими данными
# путь к папке с исходными файлами в формате .csv (копируются в том же самом виде из папки INPUT_DATA_DIRECTORY_PATH
OUTPUT_INPUT_DATA_FORMAT_CSV_DIRECTORY_PATH = DATA_PATH / "OUTPUT" / "INPUT_DATA_FORMAT_CSV"
# путь к папке к исходными файлами в формате .xlsx
OUTPUT_INPUT_DATA_FORMAT_XLSX_DIRECTORY_PATH = DATA_PATH / "OUTPUT" / "INPUT_DATA_FORMAT_XLSX"
# путь к папке с отчетами
OUTPUT_REPORTS_DIRECTORY_PATH = DATA_PATH / "OUTPUT" / "OUTPUT_REPORTS"
# путь к папке со вспомогательными файлами
OUTPUT_AUXILIARY_FILES_DIRECTORY_PATH = DATA_PATH / "OUTPUT" / "OUTPUT_AUXILIARY_FILES"

# путь к файлу, где описаны все атрибуты (у которых есть извлечение)
ATTRIBUTE_DICT_FILE_PATH = ROOT_PATH / "utils" / "attributes.txt"

# путь к файлу, где описаны все продукты (у которых есть извлечение)
PRODUCT_DICT_FILE_PATH = ROOT_PATH / "utils" / "products.txt"

# названия обязательных колонок в исходном excel-файле от заказчика
REG_NUMBER = 'reg_number'
DOCUMENT_INPUT_REQUEST = 'document_input_request'
DOCUMENT_VERIFICATION_REQUEST = 'document_verification_request'

# названия колонок с данными в вспомогательных excel-файлах
FILE_NAME_COLUMN_NAME = 'Наим.файла'
SYSTEM_ATTRIBUTE_NAME_COLUNM_NAME = 'Сис наим.атрибута'
ATTRIBUTE_NAME_COLUNM_NAME = 'Наим.атрибута'
OUTPUT_DATA_COLUMN_NAME = 'Верифицированное значение'
INPUT_DATA_COLUMN_NAME = 'Распознанное значение'
COMPARISON_COLUMN_NAME = 'Сравнение значений'
COMMON_ATTRIBUTE_NAME_COLUMN_NAME = 'Общее наим.атрибута'

# названия колонок на листе - "ложно извлеч атр."
FALSELY_COMPLETED_AMOUNT_ATTRIBUTE_COLUMN_NAME = 'Количество ложно извлеченных атрибутов'

# названия колонок на итоговом листе отчета
DOCUMENT_NAME = 'Шаблон'
COUNT_DATE = 'Дата подсчета'
AMOUNT_EXAMPLES = 'Количество примеров'
TRUE_ATTRIBUTE_AMOUNT = 'Количество верно извлеченных атрибутов'
FALSE_ATTRIBUTE_AMOUNT = 'Количество неверно извлеченных атрибутов'
QUALITY_PERCENT = 'Качество извлечения (процент)'
FALSELY_COMPLETED_AVERAGE_AMOUNT = 'Среднее количество ложно извлеченных на документ'

# названия листов в excel-файле отчета
FINAL_REPORT_SHEET_NAME = 'общая статистика'
ONLY_COMPLETED_REPORT_SHEET_NAME = 'только по зап-ым атр.'
ATTRIBUTE_STATISTICS_REPORT_SHEET_NAME = 'детализация поатрибутивно'
FALSELY_COMPLETED_REPORT_SHEET_NAME = 'ложно извлеч атр.'
GENERAL_REPORT_SHEET_NAME = 'ст-ка по всем атр.'
