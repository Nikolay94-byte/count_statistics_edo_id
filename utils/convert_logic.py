import openpyxl
from typing import Dict, Any
from utils.utils import convert_file_attributes_to_dict, convert_file_products_to_dict
from utils.constants import ATTRIBUTE_DICT_FILE_PATH, PRODUCT_DICT_FILE_PATH


def write_headers_to_exel(sheet: openpyxl.worksheet.worksheet.Worksheet) -> None:
    """Записывает заголовки в Exel-файл заготовку."""
    for cell_coordinate, attribute_name in convert_file_attributes_to_dict(ATTRIBUTE_DICT_FILE_PATH).items():
        sheet[cell_coordinate] = attribute_name


def write_rows_to_exel(
        row_number: int,
        json_body: Dict[str, Any],
        sheet: openpyxl.worksheet.worksheet.Worksheet
) -> None:
    """Записывает строки в Exel-файл заготовку."""

    # Блок Документ
    # file_name
    sheet[row_number][0].value = json_body['regNumber'] if json_body.get('regNumber') is not None else ''
    # document_anticorr
    sheet[row_number][1].value = json_body['is_anticor'] if json_body.get('is_anticor') is not None and json_body['is_anticor'] != False else ''
    # document_outgoing_date
    #sheet[row_number][2].value = json_body['documentDate'] if json_body.get('documentDate') is not None else ''
    # document_outgoing_number
    #sheet[row_number][3].value = json_body['documentNumber'] if json_body.get('documentNumber') is not None else ''
    # document_repeatedly
    sheet[row_number][4].value = json_body['repeated'] if json_body.get('repeated') is not None and json_body['repeated'] != False else ''

    # Блок Кто запрашивает
    if json_body.get('inquirySignature') is not None:
        if json_body['inquirySignature'].get('signerPosition') is not None:
            if json_body['inquirySignature']['signerPosition'].get('name') is not None:
                sheet[row_number][5].value = json_body['inquirySignature']['signerPosition']['name']
    else:
        sheet[row_number][5].value = ''
    # applicant_signer_position(approver)
    if json_body.get('inquirySignature') is not None:
        if json_body['inquirySignature'].get('approverPosition') is not None:
            if json_body['inquirySignature']['approverPosition'].get('name') is not None:
                if sheet[row_number][5].value is None:
                    sheet[row_number][5].value = json_body['inquirySignature']['approverPosition']['name']
                else:
                    sheet[row_number][5].value = sheet[row_number][5].value + ';' + json_body['inquirySignature']['approverPosition']['name']
    else:
        sheet[row_number][5].value = None
    # applicant_signer_name
    if json_body.get('inquirySignature') is not None:
        if json_body['inquirySignature'].get('signerName') is not None:
            sheet[row_number][6].value = json_body['inquirySignature']['signerName']
    else:
        sheet[row_number][6].value = ''
    # applicant_signer_name(approver)
    if json_body.get('inquirySignature') is not None:
        if json_body['inquirySignature'].get('approverName') is not None:
            if sheet[row_number][6].value is None:
                sheet[row_number][6].value = json_body['inquirySignature']['approverName']
            else:
                sheet[row_number][6].value = sheet[row_number][6].value + ';' + json_body['inquirySignature']['approverName']
    else:
        sheet[row_number][6].value = None
    # applicant_foiv_address
    sheet[row_number][7].value = (json_body['applicant']['address']
        if json_body.get('applicant') is not None and
        json_body['applicant'].get('address') is not None
        else '')
    # applicant_foiv_name
    sheet[row_number][8].value = (json_body['applicant']['name']
        if json_body.get('applicant') is not None and
        json_body['applicant'].get('name') is not None
        else '')
    # # applicant_foiv_phone
    # sheet[row_number][9].value = (';'.join(json_body['applicant']['phones'])
    #     if json_body.get('applicant') is not None and
    #     json_body['applicant'].get('phones') is not None
    #     else '')
# Блок Объект запроса ФЛ
    # target_individual_name
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL':
                if target.get('surname') is not None and target.get('name') is not None and target.get('patronymic') is not None:
                    if sheet[row_number][10].value is None:
                        sheet[row_number][10].value = target['surname'] + ' ' + target['name'] + ' ' + target['patronymic']
                    else:
                        sheet[row_number][10].value = sheet[row_number][10].value + ';' + target['surname'] + ' ' + target['name'] + ' ' + target['patronymic']
                elif target.get('surname') is not None and target.get('name') is not None and target.get('patronymic') is None:
                    if sheet[row_number][10].value is None:
                        sheet[row_number][10].value = target['surname'] + ' ' + target['name']
                    else:
                        sheet[row_number][10].value = sheet[row_number][10].value + ';' + target['surname'] + ' ' + target['name']
    # target_individual_reg_address
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL':
                if target.get('regAddress') is not None:
                    if sheet[row_number][11].value is None:
                        sheet[row_number][11].value = target['regAddress']
                    else:
                        sheet[row_number][11].value = sheet[row_number][11].value + ';' + target['regAddress']
                else:
                    sheet[row_number][11].value = None
    #target_individual_date_of_birth
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL':
                if target.get('birthDate') is not None:
                    if sheet[row_number][12].value is None:
                        sheet[row_number][12].value = target['birthDate']
                    else:
                        sheet[row_number][12].value = sheet[row_number][12].value + ';' + target['birthDate']
                else:
                    sheet[row_number][12].value = None
    # target_individual_death_date
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL':
                if target.get('deathDate') is not None:
                    if sheet[row_number][13].value is None:
                        sheet[row_number][13].value = target['deathDate']
                    else:
                        sheet[row_number][13].value = sheet[row_number][13].value + ';' + target['deathDate']
                else:
                    sheet[row_number][13].value = None
    # target_individual_inn
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL':
                if target.get('inn') is not None:
                    if sheet[row_number][14].value is None:
                        sheet[row_number][14].value = target['inn']
                    else:
                        sheet[row_number][14].value = sheet[row_number][14].value + ';' + target['inn']
                else:
                    sheet[row_number][14].value = None
    # target_individual_birth_place
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL':
                if target.get('birthPlace') is not None:
                    if sheet[row_number][15].value is None:
                        sheet[row_number][15].value = target['birthPlace']
                    else:
                        sheet[row_number][15].value = sheet[row_number][15].value + ';' + target['birthPlace']
                else:
                    sheet[row_number][15].value = None
    # target_individual_phones
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL':
                if target.get('phones') is not None:
                    if sheet[row_number][16].value is None:
                        sheet[row_number][16].value = ';'.join(target['phones'])
                    else:
                        sheet[row_number][16].value = sheet[row_number][16].value + ';' + ';'.join(target['phones'])
                else:
                    sheet[row_number][16].value = None
    # target_individual_fact_address
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL':
                if target.get('factAddress') is not None:
                    if sheet[row_number][17].value is None:
                        sheet[row_number][17].value = target['factAddress']
                    else:
                        sheet[row_number][17].value = sheet[row_number][17].value + ';' + target['factAddress']
                else:
                    sheet[row_number][17].value = None
# Блок Объект запроса ФЛ Идентиф.документ
    # target_individual_dul_date
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL':
                if target.get('identityDocument') is not None:
                    if target['identityDocument'].get('issueDate') is not None:
                        if sheet[row_number][18].value is None:
                            sheet[row_number][18].value = target['identityDocument']['issueDate']
                        else:
                            sheet[row_number][18].value = sheet[row_number][18].value + ';' + target['identityDocument']['issueDate']
                else:
                    sheet[row_number][18].value = None
    # target_individual_dul_issue_code
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL':
                if target.get('identityDocument') is not None:
                    if target['identityDocument'].get('issueCode') is not None:
                        if sheet[row_number][19].value is None:
                            sheet[row_number][19].value = target['identityDocument']['issueCode']
                        else:
                            sheet[row_number][19].value = sheet[row_number][19].value + ';' + target['identityDocument']['issueCode']
                else:
                    sheet[row_number][19].value = None
    # target_individual_identitydocument_series
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL':
                if target.get('identityDocument') is not None:
                    if target['identityDocument'].get('series') is not None:
                        if sheet[row_number][20].value is None:
                            sheet[row_number][20].value = target['identityDocument']['series']
                        else:
                            sheet[row_number][20].value = sheet[row_number][20].value + ';' + target['identityDocument']['series']
                else:
                    sheet[row_number][20].value = None
    # target_individual_identitydocument_number
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL':
                if target.get('identityDocument') is not None:
                    if target['identityDocument'].get('number') is not None:
                        if sheet[row_number][21].value is None:
                            sheet[row_number][21].value = target['identityDocument']['number']
                        else:
                            sheet[row_number][21].value = sheet[row_number][21].value + ';' + target['identityDocument']['number']
                else:
                    sheet[row_number][21].value = None
    # target_individual_dul_org
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL':
                if target.get('identityDocument') is not None:
                    if target['identityDocument'].get('issuedBy') is not None:
                        if sheet[row_number][22].value is None:
                            sheet[row_number][22].value = target['identityDocument']['issuedBy']
                        else:
                            sheet[row_number][22].value = sheet[row_number][22].value + ';' + target['identityDocument']['issuedBy']
                else:
                    sheet[row_number][22].value = None
    # target_individual_dul_type
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL':
                if target.get('identityDocument') is not None:
                    if target['identityDocument'].get('identityDocumentType') is not None:
                        if target['identityDocument']['identityDocumentType'].get('name') is not None:
                            if sheet[row_number][23].value is None:
                                sheet[row_number][23].value = target['identityDocument']['identityDocumentType']['name']
                            else:
                                sheet[row_number][23].value = sheet[row_number][23].value + ';' + target['identityDocument']['identityDocumentType']['name']
                else:
                    sheet[row_number][23].value = None
# Блок Объект запроса ИП
    # target_individual_entrepreneur_name
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL_ENTREPRENEUR':
                if target.get('fullName') is not None:
                    if sheet[row_number][24].value is None:
                        sheet[row_number][24].value = target['fullName']
                    else:
                        sheet[row_number][24].value = sheet[row_number][24].value + ';' + target['fullName']
                else:
                    sheet[row_number][24].value = None
    # target_individual_entrepreneur_regadress
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL_ENTREPRENEUR':
                if target.get('regAddress') is not None:
                    if sheet[row_number][25].value is None:
                        sheet[row_number][25].value = target['regAddress']
                    else:
                        sheet[row_number][25].value = sheet[row_number][25].value + ';' + target['regAddress']
                else:
                    sheet[row_number][25].value = None
    # target_individual_entrepreneur_date_of_birth
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL_ENTREPRENEUR':
                if target.get('birthDate') is not None:
                    if sheet[row_number][26].value is None:
                        sheet[row_number][26].value = target['birthDate']
                    else:
                        sheet[row_number][26].value = sheet[row_number][26].value + ';' + target['birthDate']
                else:
                    sheet[row_number][26].value = None
    # target_individual_entrepreneur_inn
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL_ENTREPRENEUR':
                if target.get('inn') is not None:
                    if sheet[row_number][27].value is None:
                        sheet[row_number][27].value = target['inn']
                    else:
                        sheet[row_number][27].value = sheet[row_number][27].value + ';' + target['inn']
                else:
                    sheet[row_number][27].value = None
    # target_individual_entrepreneur_phones
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL_ENTREPRENEUR':
                if target.get('phones') is not None:
                    if sheet[row_number][28].value is None:
                        sheet[row_number][28].value = ';'.join(target['phones'])
                    else:
                        sheet[row_number][28].value = sheet[row_number][28].value + ';' + ';'.join(target['phones'])
                else:
                    sheet[row_number][28].value = None
    # target_individual_entrepreneur_factadress
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_INDIVIDUAL_ENTREPRENEUR':
                if target.get('factAddress') is not None:
                    if sheet[row_number][29].value is None:
                        sheet[row_number][29].value = target['factAddress']
                    else:
                        sheet[row_number][29].value = sheet[row_number][29].value + ';' + target['factAddress']
                else:
                    sheet[row_number][29].value = None
# Блок Объект запроса ЮЛ
    # target_company_name
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_COMPANY':
                if target.get('fullName') is not None:
                    if sheet[row_number][30].value is None:
                        sheet[row_number][30].value = target['fullName']
                    else:
                        sheet[row_number][30].value = sheet[row_number][30].value + ';' + target['fullName']
                else:
                    sheet[row_number][30].value = None
    # target_company_regadress
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_COMPANY':
                if target.get('regAddress') is not None:
                    if sheet[row_number][31].value is None:
                        sheet[row_number][31].value = target['regAddress']
                    else:
                        sheet[row_number][31].value = sheet[row_number][31].value + ';' + target['regAddress']
                else:
                    sheet[row_number][31].value = None
    # target_company_taxpayer_number
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_COMPANY':
                if target.get('inn') is not None:
                    if sheet[row_number][32].value is None:
                        sheet[row_number][32].value = target['inn']
                    else:
                        sheet[row_number][32].value = sheet[row_number][32].value + ';' + target['inn']
                else:
                    sheet[row_number][32].value = None
    # target_company_kpp
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_COMPANY':
                if target.get('kpp') is not None:
                    if sheet[row_number][33].value is None:
                        sheet[row_number][33].value = target['kpp']
                    else:
                        sheet[row_number][33].value = sheet[row_number][33].value + ';' + target['kpp']
                else:
                    sheet[row_number][33].value = None
    # target_company_ogrn
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_COMPANY':
                if target.get('ogrn') is not None:
                    if sheet[row_number][34].value is None:
                        sheet[row_number][34].value = target['ogrn']
                    else:
                        sheet[row_number][34].value = sheet[row_number][34].value + ';' + target['ogrn']
                else:
                    sheet[row_number][34].value = None
    # target_company_phones
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_COMPANY':
                if target.get('phones') is not None:
                    if sheet[row_number][35].value is None:
                        sheet[row_number][35].value = ';'.join(target['phones'])
                    else:
                        sheet[row_number][35].value = sheet[row_number][35].value + ';' + ';'.join(target['phones'])
                else:
                    sheet[row_number][35].value = None
    # target_company_factadress
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target['targetType']['codeName'] == 'TARGET_COMPANY':
                if target.get('factAddress') is not None:
                    if sheet[row_number][36].value is None:
                        sheet[row_number][36].value = target['factAddress']
                    else:
                        sheet[row_number][36].value = sheet[row_number][36].value + ';' + target['factAddress']
                else:
                    sheet[row_number][36].value = None
# Блок Основание для запроса
    # basis_date_court_act
    if json_body.get('requestBasis') is not None:
        if json_body['requestBasis'].get('date') is not None:
            sheet[row_number][37].value = json_body['requestBasis']['date']
    else:
        sheet[row_number][37].value = ''
    # basis_custom_offence_value
    if json_body.get('requestBasis') is not None:
        if json_body['requestBasis'].get('values') is not None:
            if json_body['requestBasis'].get('requestBasisType') is not None:
                if json_body['requestBasis']['requestBasisType']['codeName'] == 'BASIS_ADMINISTRATIVE_OFFENCE':
                    sheet[row_number][38].value = ';'.join(json_body['requestBasis']['values'])
        else:
            sheet[row_number][38].value = ''
    # basis_arbitration_case_value
    if json_body.get('requestBasis') is not None:
        if json_body['requestBasis'].get('values') is not None:
            if json_body['requestBasis'].get('requestBasisType') is not None:
                if json_body['requestBasis']['requestBasisType']['codeName'] == 'BASIS_ARBITRATION_CASE':
                    sheet[row_number][39].value = ';'.join(json_body['requestBasis']['values'])
        else:
            sheet[row_number][39].value = ''
    # basis_civil_case_value
    if json_body.get('requestBasis') is not None:
        if json_body['requestBasis'].get('values') is not None:
            if json_body['requestBasis'].get('requestBasisType') is not None:
                if json_body['requestBasis']['requestBasisType']['codeName'] == 'BASIS_CIVIL_CASE':
                    sheet[row_number][40].value = ';'.join(json_body['requestBasis']['values'])
        else:
            sheet[row_number][40].value = ''
    # basis_position_value
    if json_body.get('requestBasis') is not None:
        if json_body['requestBasis'].get('position') is not None:
            if json_body['requestBasis']['position']['name'] is not None:
                sheet[row_number][41].value = json_body['requestBasis']['position']['name']
    else:
        sheet[row_number][41].value = ''
    # basis_inforcement_proceeding_value
    if json_body.get('requestBasis') is not None:
        if json_body['requestBasis'].get('values') is not None:
            if json_body['requestBasis'].get('requestBasisType') is not None:
                if json_body['requestBasis']['requestBasisType']['codeName'] == 'BASIS_ENFORCEMENT_PROCEEDING':
                    sheet[row_number][42].value = ';'.join(json_body['requestBasis']['values'])
        else:
            sheet[row_number][42].value = ''
    # basis_preliminary_inquiry_value
    if json_body.get('requestBasis') is not None:
        if json_body['requestBasis'].get('values') is not None:
            if json_body['requestBasis'].get('requestBasisType') is not None:
                if json_body['requestBasis']['requestBasisType']['codeName'] == 'BASIS_PRELIMINARY_INQUIRY':
                    sheet[row_number][43].value = ';'.join(json_body['requestBasis']['values'])
        else:
            sheet[row_number][43].value = ''
    # basis_inheritance_case_value
    if json_body.get('requestBasis') is not None:
        if json_body['requestBasis'].get('values') is not None:
            if json_body['requestBasis'].get('requestBasisType') is not None:
                if json_body['requestBasis']['requestBasisType']['codeName'] == 'BASIS_INHERITANCE_CASE':
                    sheet[row_number][44].value = ';'.join(json_body['requestBasis']['values'])
        else:
            sheet[row_number][44].value = ''
    # basis_court_order_value
    if json_body.get('requestBasis') is not None:
        if json_body['requestBasis'].get('values') is not None:
            if json_body['requestBasis'].get('requestBasisType') is not None:
                if json_body['requestBasis']['requestBasisType']['codeName'] == 'BASIS_COURT_ORDER':
                    sheet[row_number][45].value = ';'.join(json_body['requestBasis']['values'])
        else:
            sheet[row_number][45].value = ''
    # basis_claim_value
    if json_body.get('requestBasis') is not None:
        if json_body['requestBasis'].get('values') is not None:
            if json_body['requestBasis'].get('requestBasisType') is not None:
                if json_body['requestBasis']['requestBasisType']['codeName'] == 'BASIS_CLAIM':
                    sheet[row_number][46].value = ';'.join(json_body['requestBasis']['values'])
        else:
            sheet[row_number][46].value = ''
    # basis_criminal_case_value
    if json_body.get('requestBasis') is not None:
        if json_body['requestBasis'].get('values') is not None:
            if json_body['requestBasis'].get('requestBasisType') is not None:
                if json_body['requestBasis']['requestBasisType']['codeName'] == 'BASIS_CRIMINAL_CASE':
                    sheet[row_number][47].value = ';'.join(json_body['requestBasis']['values'])
        else:
            sheet[row_number][47].value = ''
    # basis_legal_clause_lc_fz311
    if json_body.get('requestBasis') is not None:
        if json_body['requestBasis'].get('legalClauseTypes') is not None:
            sheet[row_number][48].value = ';'.join(json_body['requestBasis']['legalClauseTypes'])
    else:
        sheet[row_number][48].value = ''
    # basis_legal_clause_lc_a15_notary
    if json_body.get('requestBasis') is not None:
        if json_body['requestBasis'].get('legalClauseTypes') is not None:
            sheet[row_number][49].value = ';'.join(json_body['requestBasis']['legalClauseTypes'])
    else:
        sheet[row_number][49].value = ''
    # basis_legal_clause_lc_a23_fz173
    if json_body.get('requestBasis') is not None:
        if json_body['requestBasis'].get('legalClauseTypes') is not None:
            sheet[row_number][50].value = ';'.join(json_body['requestBasis']['legalClauseTypes'])
    else:
        sheet[row_number][50].value = ''
# Блок Особые условия
#     # conditions_give_out_on_purpose
#     sheet[row_number][51].value = json_body['isPersonally'] if json_body.get('isPersonally') is not None and json_body['isPersonally'] != False else ''
#     # conditions_whom_send_response
#     sheet[row_number][52].value = json_body['responseRecipient'] if json_body.get('responseRecipient') is not None else ''
#     # conditions_where_send_response
#     sheet[row_number][53].value = json_body['responseRecipientAddress'] if json_body.get('responseRecipientAddress') is not None else ''
#     # conditions_in_digital_format
#     sheet[row_number][54].value = json_body['isNeedDisk'] if json_body.get('isNeedDisk') is not None and json_body['isNeedDisk'] != False else ''
#     # conditions_response_deadline
#     sheet[row_number][55].value = json_body['dueDate'] if json_body.get('dueDate') is not None else ''
#     # conditions_in_excel_format
#     sheet[row_number][56].value = json_body['isNeedExcelFormat'] if json_body.get('isNeedExcelFormat') is not None and json_body['isNeedExcelFormat'] != False else ''
# Блок Продукты запроса Значения продукта
    # product_card_attr_number
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target.get('products') is not None:
                for product in target['products']:
                    if product['productType']['codeName'] == 'PROD_CARD':
                        if product.get('values') is not None:
                            if sheet[row_number][57].value is None:
                                sheet[row_number][57].value = ';'.join(product['values'])
                            else:
                                sheet[row_number][57].value = sheet[row_number][57].value + ';' + ';'.join(product['values'])
    # product_account_number
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target.get('products') is not None:
                for product in target['products']:
                    if product['productType']['codeName'] == 'PROD_ACCOUNT':
                        if product.get('values') is not None:
                            if sheet[row_number][58].value is None:
                                sheet[row_number][58].value = ';'.join(product['values'])
                            else:
                                sheet[row_number][58].value = sheet[row_number][58].value + ';' + ';'.join(product['values'])

    # Заполняем продукты запроса
    for attribute, parameters in convert_file_products_to_dict(PRODUCT_DICT_FILE_PATH).items():
        write_productType(row_number, json_body, sheet, parameters[0], parameters[1], parameters[2])

def write_productType(
        row_number: int,
        json_body: Dict[str, Any],
        sheet: openpyxl.worksheet.worksheet.Worksheet,
        product_name: str,
        attr_name: str,
        header_number: int):
    """Записывает продукт запроса из json в Exel-файл"""
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
