import openpyxl
from typing import Any, Optional, Dict, Union, List
from utils.utils import convert_file_attributes_to_dict, convert_file_products_to_dict
from utils.constants import ATTRIBUTE_DICT_FILE_PATH, PRODUCT_DICT_FILE_PATH


def write_headers_to_excel(sheet: openpyxl.worksheet.worksheet.Worksheet) -> None:
    """Записывает заголовки в excel-файл заготовку."""
    for cell_coordinate, attribute_name in convert_file_attributes_to_dict(ATTRIBUTE_DICT_FILE_PATH).items():
        sheet[cell_coordinate] = attribute_name


def _set_value(
        sheet: openpyxl.worksheet.worksheet.Worksheet,
        row: int,
        col: int,
        value: Optional[Union[str, int, float, bool, list, dict]]
) -> None:
    """Устанавливает значение в ячейку, если оно не None."""
    if value is not None:
        if sheet[row][col].value is None:
            sheet[row][col].value = str(value)
        else:
            sheet[row][col].value = f"{sheet[row][col].value};{value}"


def _get_nested_value(
        data: Dict[str, Any],
        *keys: str,
        default: Optional[Any] = None
) -> Optional[Any]:
    """Получает значение из вложенного словаря по цепочке ключей."""
    for key in keys:
        if isinstance(data, dict) and key in data:
            data = data[key]
        else:
            return default
    return data


def write_rows_to_excel(
        row_number: int,
        json_body: Dict[str, Any],
        sheet: openpyxl.worksheet.worksheet.Worksheet
) -> None:
    """Записывает строки в Excel-файл заготовку."""

    # Блок Документ
    _set_value(sheet, row_number, 0, _get_nested_value(json_body, 'regNumber'))
    _set_value(sheet, row_number, 1, _get_nested_value(json_body, 'is_anticor') if json_body.get('is_anticor') else '')
    _set_value(sheet, row_number, 4, _get_nested_value(json_body, 'repeated') if json_body.get('repeated') else '')

    # Блок Кто запрашивает
    signer_pos = _get_nested_value(json_body, 'inquirySignature', 'signerPosition', 'name')
    approver_pos = _get_nested_value(json_body, 'inquirySignature', 'approverPosition', 'name')
    _set_value(sheet, row_number, 5, ';'.join(filter(None, [signer_pos, approver_pos])))

    signer_name = _get_nested_value(json_body, 'inquirySignature', 'signerName')
    approver_name = _get_nested_value(json_body, 'inquirySignature', 'approverName')
    _set_value(sheet, row_number, 6, ';'.join(filter(None, [signer_name, approver_name])))

    _set_value(sheet, row_number, 7, _get_nested_value(json_body, 'applicant', 'address'))
    _set_value(sheet, row_number, 8, _get_nested_value(json_body, 'applicant', 'name'))

    # Оптимизированный парсинг targets
    def _process_targets(
            targets: List[Dict[str, Any]],
            target_type: str,
            field_mapping: Dict[str, int]
    ) -> None:
        """Обрабатывает targets заданного типа и заполняет ячейки согласно field_mapping"""
        for target in targets:
            if target['targetType']['codeName'] != target_type:
                continue

            for field, col in field_mapping.items():
                if '.' in field:
                    # Для вложенных полей (например, identityDocument.issueDate)
                    parts = field.split('.')
                    value = target
                    try:
                        for part in parts:
                            value = value[part]
                    except (KeyError, TypeError):
                        value = None
                else:
                    value = target.get(field)

                if value is not None:
                    if isinstance(value, list):
                        value = ';'.join(value)
                    _set_value(sheet, row_number, col, value)

    # Маппинг полей для разных типов targets
    target_mappings = {
        'TARGET_INDIVIDUAL': {
            'surname,name,patronymic': 10,  # Специальный случай для ФИО
            'regAddress': 11,
            'birthDate': 12,
            'deathDate': 13,
            'inn': 14,
            'birthPlace': 15,
            'phones': 16,
            'factAddress': 17,
            'identityDocument.issueDate': 18,
            'identityDocument.issueCode': 19,
            'identityDocument.series': 20,
            'identityDocument.number': 21,
            'identityDocument.issuedBy': 22,
            'identityDocument.identityDocumentType.name': 23
        },
        'TARGET_INDIVIDUAL_ENTREPRENEUR': {
            'fullName': 24,
            'regAddress': 25,
            'birthDate': 26,
            'inn': 27,
            'phones': 28,
            'factAddress': 29
        },
        'TARGET_COMPANY': {
            'fullName': 30,
            'regAddress': 31,
            'inn': 32,
            'kpp': 33,
            'ogrn': 34,
            'phones': 35,
            'factAddress': 36
        }
    }

    if json_body.get('targets'):
        targets = json_body['targets']

        # Обработка специального случая для ФИО (TARGET_INDIVIDUAL)
        for target in targets:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL':
                parts = []
                if target.get('surname'):
                    parts.append(target['surname'])
                if target.get('name'):
                    parts.append(target['name'])
                if target.get('patronymic'):
                    parts.append(target['patronymic'])

                if parts:
                    _set_value(sheet, row_number, 10, ' '.join(parts))

        # Обработка остальных полей
        for target_type, field_mapping in target_mappings.items():
            _process_targets(targets, target_type, field_mapping)

    # Блок Основание для запроса
    basis_type_mapping = {
        'BASIS_ADMINISTRATIVE_OFFENCE': 38,
        'BASIS_ARBITRATION_CASE': 39,
        'BASIS_CIVIL_CASE': 40,
        'BASIS_ENFORCEMENT_PROCEEDING': 42,
        'BASIS_PRELIMINARY_INQUIRY': 43,
        'BASIS_INHERITANCE_CASE': 44,
        'BASIS_COURT_ORDER': 45,
        'BASIS_CLAIM': 46,
        'BASIS_CRIMINAL_CASE': 47
    }

    if json_body.get('requestBasis'):
        basis = json_body['requestBasis']
        _set_value(sheet, row_number, 37, basis.get('date'))
        _set_value(sheet, row_number, 41, _get_nested_value(basis, 'position', 'name'))

        if basis.get('requestBasisType') and basis.get('values'):
            code = basis['requestBasisType']['codeName']
            if code in basis_type_mapping:
                _set_value(sheet, row_number, basis_type_mapping[code], ';'.join(basis['values']))

        _set_value(sheet, row_number, 48, ';'.join(basis.get('legalClauseTypes', [])))
        _set_value(sheet, row_number, 49, ';'.join(basis.get('legalClauseTypes', [])))
        _set_value(sheet, row_number, 50, ';'.join(basis.get('legalClauseTypes', [])))

    # Блок Значения продукта запроса - номер карты или номер счета
    def _write_product_value(
            product_type: str,
            header_number: int
    ):
        """Записывает значения продукта запроса из json в excel-файл"""
        if not json_body.get('targets'):
            return

        values = []
        for target in json_body['targets']:
            if target.get('products'):
                for product in target['products']:
                    if product['productType']['codeName'] == product_type and product.get('values'):
                        values.extend(product['values'])

        if values:
            _set_value(sheet, row_number, header_number, ';'.join(values))

    _write_product_value('PROD_CARD', 57)
    _write_product_value('PROD_ACCOUNT', 58)

    # Блок Продукты запроса
    def _write_productType(
            row_number: int,
            json_body: Dict[str, Any],
            sheet: openpyxl.worksheet.worksheet.Worksheet,
            product_name: str,
            attr_name: str,
            header_number: int
    ):
        """Записывает продукт запроса из json в excel-файл"""
        targets = json_body.get('targets', [])
        for target in targets:
            products = target.get('products', [])
            for product in products:
                if product.get('productType', {}).get('codeName') != product_name:
                    continue

                attributes = product.get('attributes', [])
                for attribute in attributes:
                    if attribute.get('attributeType', {}).get('codeName') == attr_name:
                        sheet[row_number][header_number].value = 'true'
                        break

    for attribute, parameters in convert_file_products_to_dict(PRODUCT_DICT_FILE_PATH).items():
        _write_productType(row_number, json_body, sheet, parameters[0], parameters[1], parameters[2])
