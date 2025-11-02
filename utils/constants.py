from settings import DATA_PATH


# Путь к папке с исходными файлами
INPUT_DATA_DIRECTORY_PATH = DATA_PATH / 'INPUT'

# путь к папке с отчетами
OUTPUT_REPORTS_DIRECTORY_PATH = DATA_PATH / 'OUTPUT' / 'OUTPUT_REPORTS'

# Наименование файла с эталонными данными из таблиц
INPUT_DATA_ETALON_FILE = 'INPUT_DATA_ETALON'

# Наименование файла с извлеченными данными из таблиц
INPUT_DATA_RECOGINIZED_FILE = 'INPUT_DATA_RECOGINIZED'

# Наименования колонок в файлах
FILE_NAME = 'Наименование файла с расширением'
ATTRIBUTE_NAME = 'Параметр'
ATTRIBUTE_NAME_RUS = 'Параметр с переводом'
ETALON_VALUE = 'Эталонное значение'
RECOGINIZED_VALUE = 'Распознанное значение'
CELL_SCORE = 'Оценка ячейка'
COLUMN_SCORE = 'Оценка столбец'
PACKAGE_SCORE = 'Оценка пакет'

# Словарь наименования колонок таблиц, по которым ведется подсчет
TABLE_ATTRIBUTES = {
    'table_description_column_row_attribute': 'Наименование товара/услуги',
    'table_qty_column_row_attribute': 'Количество товара/услуг по позициям',
    'table_unit_column_row_attribute': 'Ед.изм',
    'table_cost_column_row_attribute': 'Цена по позициям',
    'table_cost_without_tax_column_row_attribute': 'Сумма без НДС по позициям',
    'table_cost_with_tax_column_row_attribute': 'Сумма с НДС по позициям'
}

# Названия листов в excel-файле отчета
FINAL_REPORT_SHEET_NAME = 'общая статистика'
PAKET_STATISTICS_REPORT_SHEET_NAME = 'детализация попакетно'
COLUMN_STATISTICS_REPORT_SHEET_NAME = 'детализация по колонкам'
ETALON_DATA_SHEET_NAME = 'эталонные значения'
RECOGINIZED_DATA_SHEET_NAME = 'распознанные значения'

# Названия колонок на итоговом листе отчета
DOCUMENT_NAME = 'Шаблон'
COUNT_DATE = 'Дата подсчета'
AMOUNT_EXAMPLES = 'Количество примеров'
QUALITY_PERCENT = 'Качество извлечения таблицы (процент)'

# Названия колонок на листе 'детализация по колонкам'
COLUMN_NAME = 'Наименование колонки'
COLUMN_QUALITY = 'Качество извлечения колонки'

# Наименования шаблонов с таблицами
