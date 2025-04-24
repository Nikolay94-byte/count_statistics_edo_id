from settings import DATA_PATH, ROOT_PATH


# путь к исходному файлу от заказчика
INPUT_FILE_PATH = DATA_PATH / "Выгрузка с прода ФОИВ.xlsx"

# путь к файлу, где описаны все атрибуты (у которых есть извлечение)
ATTRIBUTE_DICT_FILE_PATH = ROOT_PATH / "utils" / "attributes.txt"

# путь к файлу, где описаны все продукты (у которых есть извлечение)
PRODUCT_DICT_FILE_PATH = ROOT_PATH / "utils" / "products.txt"

# названия колонок в исходном excel-файле от заказчика
REG_NUMBER = 'reg_number'
DOCUMENT_TYPE = 'document_type'
DOCUMENT_INPUT_REQUEST = 'document_input_request'
DOCUMENT_VERIFICATION_REQUEST = 'document_verification_request'

# названия колонок с данными в подготовительных excel-файлах
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
