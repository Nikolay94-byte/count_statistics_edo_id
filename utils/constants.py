from settings import DATA_PATH, ROOT_PATH


# INPUT, путь к папке с исходными файлами в формате .csv
INPUT_DATA_DIRECTORY_PATH = DATA_PATH / "INPUT"

# OUTPUT, пути к папкам с исходящими данными
# путь к папке с исходными файлами в формате .csv (копируются в том же самом виде из папки INPUT_DATA_DIRECTORY_PATH)
OUTPUT_INPUT_DATA_FORMAT_CSV_DIRECTORY_PATH = DATA_PATH / "OUTPUT" / "INPUT_DATA_FORMAT_CSV"
# путь к папке к исходными файлами в формате .xlsx
OUTPUT_INPUT_DATA_FORMAT_XLSX_DIRECTORY_PATH = DATA_PATH / "OUTPUT" / "INPUT_DATA_FORMAT_XLSX"
# путь к папке с отчетами
OUTPUT_REPORTS_DIRECTORY_PATH = DATA_PATH / "OUTPUT" / "OUTPUT_REPORTS"

# названия обязательных колонок в исходном excel-файле выгрузке
REGNUMBER = 'regnumber'
ATTRIBUTE_NAME = 'attribute_name'
RUS_ATTRIBUTE_NAME = 'rus_attribute_name'
TEXT_NORMALIZED = 'text_normalized'
TEXT_VERIFICATION = 'text_verification'

# названия колонок на итоговом листе отчета
DOCUMENT_NAME = 'Шаблон'
COUNT_DATE = 'Дата подсчета'
AMOUNT_EXAMPLES = 'Количество примеров'
TRUE_ATTRIBUTE_AMOUNT = 'Количество верно извлеченных атрибутов'
FALSE_ATTRIBUTE_AMOUNT = 'Количество неверно извлеченных атрибутов'
QUALITY_PERCENT = 'Качество извлечения (процент)'

# названия колонок на листе статистики по пакетно
FILE_NAME_COLUMN_NAME = 'Наим.файла'
SYSTEM_ATTRIBUTE_NAME_COLUNM_NAME = 'Сис наим.атрибута'
ATTRIBUTE_NAME_COLUNM_NAME = 'Наим.атрибута'
OUTPUT_DATA_COLUMN_NAME = 'Верифицированное значение'
INPUT_DATA_COLUMN_NAME = 'Распознанное значение'
COMPARISON_COLUMN_NAME = 'Сравнение значений'

# названия листов в excel-файле отчета
FINAL_REPORT_SHEET_NAME = 'общая статистика'
PAKET_REPORT_SHEET_NAME = 'детализация попакетно'
ATTRIBUTE_STATISTICS_REPORT_SHEET_NAME = 'детализация поатрибутивно'
INPUT_DATA_SHEET_NAME = 'исходые данные'

# список атрибутов для подсчета по Исполнительному листу
PERF_LIST_ATTRIBUTES = {
    'document_number': '01.Документ-основание - 02.Номер документа',
    'document_date': '01.Документ-основание - 03.Дата документа',
    'document_case_number': '01.Документ-основание - 05.Номер дела',
    'document_case_date': '01.Документ-основание - 06.Дата дела',
    'document_authority': '01.Документ-основание - 07.Орган, выдавший документ',
    'claimer_entity_name': '02.Взыскатель - 01.Взыскатель - 03.Наименование на 5 стр',
    'debtor_entity_type': '03.Должник - 01.Должник - 01.Тип должника',
    'debtor_entity_name': '03.Должник - 01.Должник - 03.Наименование на 5 стр',
    'debtor_entity_birthdate': '03.Должник - 01.Должник - 04.Дата рождения',
    'debtor_entity_inn': '03.Должник - 01.Должник - 05.ИНН/КИО'
}
# список атрибутов для подсчета по КТС
LABOUR_COMMISSION_ATTRIBUTES = {
    'document_number': '01.Документ-основание - 02.Номер документа',
    'document_date': '01.Документ-основание - 03.Дата документа',
    'document_case_number': '01.Документ-основание - 04.Номер дела',
    'document_case_date': '01.Документ-основание - 05.Дата дела',
    'claimer_entity_name': '02.Взыскатель - 01.Взыскатель - 02.Наименование',
    'sum_total_amount_val':	'02.Взыскатель - 01.Взыскатель - 02.Сумма - Общая сумма - 01.Значение',
    'debtor_entity_type': '03.Должник - 01.Должник - 01.Тип должника',
    'debtor_entity_name': '03.Должник - 01.Должник - 02.Наименование',
    'debtor_entity_inn': '03.Должник - 01.Должник - 03.ИНН/КИО',
}
# список атрибутов для подсчета по Судебному приказу
COURT_ORDER_ATTRIBUTES = {
    'document_case_number': '01.Документ-основание - 04.Номер дела',
    'document_case_date': '01.Документ-основание - 05.Дата дела',
    'document_authority': '01.Документ-основание - 06.Орган, выдавший документ',
    'claimer_entity_name': '02.Взыскатель - 01.Взыскатель - 02.Наименование',
    'debtor_entity_type': '03.Должник - 01.Должник - 01.Тип должника',
    'debtor_entity_name': '03.Должник - 01.Должник - 02.Наименование',
    'debtor_entity_birthdate': '03.Должник - 01.Должник - 03.Дата рождения',
    'debtor_entity_inn': '03.Должник - 01.Должник - 04.ИНН/КИО'
}
# список атрибутов для подсчета по Заявлению на отзыв
APPLICATION_FOR_WITHDRAWAL_ATTRIBUTES = {
    'document_revocable_type': '01. Документ-основание - 02.Тип отзываемого документа',
    'document_revocable_number': '01. Документ-основание - 03.Номер отзываемого документа',
    'document_revocable_date': '01. Документ-основание - 04.Дата отзываемого документа',
    'document_revocable_case_number': '01.Документ-основание - 05.Номер отзываемого дела',
    'document_revocable_case_date': '01.Документ-основание - 06.Дата отзываемого дела',
    'claimer_entity_type': '02.Взыскатель - 01.Взыскатель - 01.Тип взыскателя',
    'claimer_entity_name': '02.Взыскатель - 01.Взыскатель - 01.Наименование',
    'claimer_entity_address': '02.Взыскатель - 01.Взыскатель - 02.Адрес',
    'claimer_entity_inn': '02.Взыскатель - 01.Взыскатель - 04.ИНН',
    'debtor_entity_type': '03.Должник - 01.Должник - 01.Тип должника',
    'debtor_entity_name': '03.Должник - 01.Должник - 01.Наименование',
    'debtor_entity_birthdate': '03.Должник - 01.Должник - 02.Дата рождения',
    'debtor_entity_inn': '03.Должник - 01.Должник - 03.ИНН'
}
# список атрибутов для подсчета по Заявлению на взыскание
APPLICATION_FOR_THE_RECOVERY_ATTRIBUTES = {
    'recovery_claimer_entity_type': '02.Взыскатель - 01.Взыскатель - 01.Тип взыскателя',
    'recovery_claimer_entity_name': '02.Взыскатель - 01.Взыскатель - 02.Наименование получателя',
    'recovery_claimer_entity_inn': '02.Взыскатель - 01.Взыскатель - 03.ИНН',
    'recovery_claimer_entity_kpp': '02.Взыскатель - 01.Взыскатель - 04.КПП',
    'recovery_claimer_entity_address': '02.Взыскатель - 01.Взыскатель - 05.Адрес',
    'recovery_claimer_entity_phone': '02.Взыскатель - 01.Взыскатель - 06.Телефон',
    'recovery_claimer_entity_bik': '02.Взыскатель - 01.Взыскатель - 07.БИК',
    'recovery_claimer_entity_transfer_account': '02.Взыскатель - 01.Взыскатель - 08.Расчетный счет',
    'recovery_claimer_entity_kbk': '02.Взыскатель - 01.Взыскатель - 09.КБК',
    'recovery_claimer_entity_oktmo': '02.Взыскатель - 01.Взыскатель - 10.ОКТМО',
    'recovery_claimer_entity_uin': '02.Взыскатель - 01.Взыскатель - 11.УИН'
}
