import openpyxl
from typing import Dict, Any

def write_headers_to_exel(sheet: openpyxl.worksheet.worksheet.Worksheet) -> None:
    """Записывает заголовки в exel файл заготовку."""
    # Блок Документ
    sheet['A1'] = 'file_name#имя файла'  # 0
    sheet['B1'] = 'document_anticorr#Документ - Антикоррупционный'  # 1
    #sheet['C1'] = 'document_outgoing_date#Документ - Исх. дата'  # 2
    #sheet['D1'] = 'document_outgoing_number#Документ - Исх. номер'  # 3
    sheet['E1'] = 'document_repeatedly#Документ - Повторно'  # 4

# Блок Кто запрашивает
    sheet['F1'] = 'applicant_signer_position#Кто запрашивает - Подписант - Должность'  # 5
    sheet['G1'] = 'applicant_signer_name#Кто запрашивает - Подписант - ФИО'  # 6
    sheet['H1'] = 'applicant_foiv_address#Кто запрашивает - ФОИВ - Адрес'  # 7
    sheet['I1'] = 'applicant_foiv_name#Кто запрашивает - ФОИВ - Наименование'  # 8
    #sheet['J1'] = 'applicant_foiv_phone#Кто запрашивает - ФОИВ - Телефон'  # 9

# Блок Объект запроса ФЛ
    sheet['K1'] = 'target_individual_name#Объект запроса - ФЛ - ФИО'  # 10
    sheet['L1'] = 'target_individual_reg_address#Объект запроса - ФЛ - Адрес регистрации'  # 11
    sheet['M1'] = 'target_individual_date_of_birth#Объект запроса - ФЛ - Дата рождения'  # 12
    sheet['N1'] = 'target_individual_death_date#Объект запроса - ФЛ - Дата смерти'  # 13
    sheet['O1'] = 'target_individual_inn#Объект запроса - ФЛ - ИНН'  # 14
    sheet['P1'] = 'target_individual_birth_place#Объект запроса - ФЛ - Место рождения'  # 15
    sheet['Q1'] = 'target_individual_phones#Объект запроса - ФЛ - Телефонные номера'  # 16
    sheet['R1'] = 'target_individual_fact_address#Объект запроса - ФЛ - Фактический адрес'  # 17

# Блок Объект запроса ФЛ Идентиф.документ
    sheet['S1'] = 'target_individual_dul_date#Объект запроса - ФЛ - Идентификационный документ - 1_Дата выдачи'  # 18
    sheet['T1'] = 'target_individual_dul_issue_code#Объект запроса - ФЛ - Идентификационный документ - 2_Код подразделения'  # 19
    sheet['U1'] = 'target_individual_identitydocument_series#Объект запроса - ФЛ - Идентификационный документ - 3_Серия'  # 20
    sheet['V1'] = 'target_individual_identitydocument_number#Объект запроса - ФЛ - Идентификационный документ - 4_Номер'  # 21
    sheet['W1'] = 'target_individual_dul_org#Объект запроса - ФЛ - Идентификационный документ - 5_Орган выдачи'  # 22
    sheet['X1'] = 'target_individual_dul_type#Объект запроса - ФЛ - Идентификационный документ - 6_Тип'  # 23

# Блок Объект запроса ИП
    sheet['Y1'] = 'target_individual_entrepreneur_name#Объект запроса - ИП - ФИО'  # 24
    sheet['Z1'] = 'target_individual_entrepreneur_regadress#Объект запроса - ИП - Адрес регистрации'  # 25
    sheet['AA1'] = 'target_individual_entrepreneur_date_of_birth#Объект запроса - ИП - Дата рождения'  # 26
    sheet['AB1'] = 'target_individual_entrepreneur_inn#Объект запроса - ИП - ИНН'  # 27
    sheet['AC1'] = 'target_individual_entrepreneur_phones#Объект запроса - ИП - Телефонные номера'  # 28
    sheet['AD1'] = 'target_individual_entrepreneur_factadress#Объект запроса - ИП - Фактический адрес'  # 29

# Блок Объект запроса ЮЛ
    sheet['AE1'] = 'target_company_name#Объект запроса - ЮЛ - Наименование'  # 30
    sheet['AF1'] = 'target_company_regadress#Объект запроса - ЮЛ - Адрес регистрации'  # 31
    sheet['AG1'] = 'target_company_taxpayer_number#Объект запроса - ЮЛ - ИНН'  # 32
    sheet['AH1'] = 'target_company_kpp#Объект запроса - ЮЛ - КПП'  # 33
    sheet['AI1'] = 'target_company_ogrn#Объект запроса - ЮЛ - ОГРН'  # 34
    sheet['AJ1'] = 'target_company_phones#Объект запроса - ЮЛ - Телефонные номера'  # 35
    sheet['AK1'] = 'target_company_factadress#Объект запроса - ЮЛ - Фактический адрес'  # 36

# Блок Основание для запроса
    sheet['AL1'] = 'basis_date_court_act#Основание для запроса - Дата основания запроса'  # 37
    sheet['AM1'] = 'basis_custom_offence_value#Основание для запроса - Административное правонарушение - Номер дела об АП'  # 38
    sheet['AN1'] = 'basis_arbitration_case_value#Основание для запроса - Арбитражное дело - Номер арбитражного дела'  # 39
    sheet['AO1'] = 'basis_civil_case_value#Основание для запроса - Гражданское дело - Номер гражданского дела'  # 40
    sheet['AP1'] = 'basis_position_value#Основание для запроса - Должность - Название должности'  # 41
    sheet['AQ1'] = 'basis_inforcement_proceeding_value#Основание для запроса - Исполнительное производство - Номер исполнительного производства'  # 42
    sheet['AR1'] = 'basis_preliminary_inquiry_value#Основание для запроса - Материал проверки - Номер материала проверки'  # 43
    sheet['AS1'] = 'basis_inheritance_case_value#Основание для запроса - Наследственное дело - Номер наследственного дела'  # 44
    sheet['AT1'] = 'basis_court_order_value#Основание для запроса - Постановление суда - Номер постановления суда'  # 45
    sheet['AU1'] = 'basis_claim_value#Основание для запроса - Претензия/жалоба - Номер претензии/жалобы'  # 46
    sheet['AV1'] = 'basis_criminal_case_value#Основание для запроса - Уголовное дело - Номер уголовного дела'  # 47
    sheet['AW1'] = 'basis_legal_clause_lc_fz311#Основание для запроса - Статьи - 311-ФЗ'  # 48
    sheet['AX1'] = 'basis_legal_clause_lc_a15_notary#Основание для запроса - Статьи - ст. 15 "О нотариате"'  # 49
    sheet['AY1'] = 'basis_legal_clause_lc_a23_fz173#Основание для запроса - Статьи - ст. 23 (173-ФЗ)'  # 50

# Блок Особые условия
    #sheet['AZ1'] = 'conditions_give_out_on_purpose#Особые условия - Выдать нарочно'  # 51
    #sheet['BA1'] = 'conditions_whom_send_response#Особые условия - Кому направить ответ'  # 52
    #sheet['BB1'] = 'conditions_where_send_response#Особые условия - Куда направить ответ'  # 53
    #sheet['BC1'] = 'conditions_in_digital_format#Особые условия - Предоставить в электронном виде'  # 54
    #sheet['BD1'] = 'conditions_response_deadline#Особые условия - Предоставить до'  # 55
    #sheet['BE1'] = 'conditions_in_excel_format#Особые условия - Формат сведений Excel'  # 56

# Блок Продукты запроса Значения продукта
    sheet['BF1'] = 'product_card_attr_number#Продукты запроса - Карта - Значения продукта - Номер карты'  # 57
    sheet['BG1'] = 'product_account_number#Продукты запроса - Счет - Значения продукта - Номер счета'  # 58

# Блок Продукты запроса Аккредитив
    sheet['BH1'] = 'product_credit_letter_attr_credit_letter_data#Продукты запроса - Аккредитив - Атрибуты продукта - Данные по аккредитивам - Данные по аккредитивам'  # 59

# Блок Продукты запроса Банковские гарантии
    sheet['BI1'] = 'product_bank_guarantees_attr_bank_guarantees_data#Продукты запроса - Банковские гарантии - Атрибуты продукта - Данные по банковским гарантиям - Данные по банковским гарантиям'  # 60
    sheet['BJ1'] = 'product_bank_guarantees_attr_contract_file#Продукты запроса - Банковские гарантии - Атрибуты продукта - Досье по договорам (продуктовое досье) - Досье по договорам (продуктовое досье)'  # 61

# Блок Продукты запроса Вексель
    sheet['BK1'] = 'product_promissory_note_attr_promissory_note_data#Продукты запроса - Вексель - Атрибуты продукта - Данные по векселям - Данные по векселям'  # 62

# Блок Продукты запроса Вклад
    sheet['BL1'] = 'product_investment_attr_client_pers_data#Продукты запроса - Вклад - Атрибуты продукта - Анкетные данные владельца счета'  # 63
    sheet['BM1'] = 'product_investment_attr_opened_account_dept_data#Продукты запроса - Вклад - Атрибуты продукта - Данные о наименовании/адресе ТП, где открыт счет'  # 64
    sheet['BN1'] = 'product_investment_attr_deposit_paid_out_percent#Продукты запроса - Вклад - Атрибуты продукта - Данные по уплаченным процентам по вкладу (депозиту)'  # 65
    sheet['BO1'] = 'product_investment_attr_account_file#Продукты запроса - Вклад - Атрибуты продукта - Досье по счету/карте (копии документов)'  # 66
    sheet['BP1'] = 'product_investment_attr_bank_statement#Продукты запроса - Вклад - Атрибуты продукта - Выписки - Выписки'  # 67
    sheet['BQ1'] = 'product_investment_attr_opened_account_pers_data#Продукты запроса - Вклад - Атрибуты продукта - Данные о лицах, открывших счета - Данные о лицах, открывших счета'  # 68
    sheet['BR1'] = 'product_investment_attr_file#Продукты запроса - Вклад - Атрибуты продукта - Картотека - Картотека'  # 69
    sheet['BS1'] = 'product_investment_attr_power_of_attorney_copy#Продукты запроса - Вклад - Атрибуты продукта - Копии доверенностей - Копии доверенностей'  # 70
    sheet['BT1'] = 'product_investment_attr_exist_account_card_cert#Продукты запроса - Вклад - Атрибуты продукта - Сведения о наличии счетов и банковских карт - Сведения о наличии счетов и банковских карт (справки)'  # 71
    sheet['BU1'] = 'product_investment_attr_phone_to_account_info#Продукты запроса - Вклад - Атрибуты продукта - Сведения о подключении телефона к счетам/картам клиента из запроса - Сведения о подключении телефона к счетам/картам клиента из запроса'  # 72
    sheet['BV1'] = 'product_investment_attr_account_balance_cert#Продукты запроса - Вклад - Атрибуты продукта - Справки (сведения) об остатках - Справки (сведения) об остатках'  # 73
    sheet['BW1'] = 'product_investment_attr_photo_video_office#Продукты запроса - Вклад - Атрибуты продукта - Фото/видео из отделений Банка - Фото/видео из отделений Банка'  # 74

# Блок Продукты запроса ДБО
    sheet['BX1'] = 'product_dbo_attr_dbo_connection_data#Продукты запроса - ДБО - Атрибуты продукта - Данные о подключении ДБО'  # 75
    sheet['BY1'] = 'product_dbo_attr_ip_address#Продукты запроса - ДБО - Атрибуты продукта - IP адреса/Log файлы - IP адреса/Log файлы'  # 76
    sheet['BZ1'] = 'product_dbo_attr_used_ip_address_pers_data#Продукты запроса - ДБО - Атрибуты продукта - Данные о лицах использовавших ip-адреса - Данные о лицах использовавших ip-адреса'  # 77
    sheet['CA1'] = 'product_dbo_attr_dbo_connected_pers_data#Продукты запроса - ДБО - Атрибуты продукта - Данные о лицах подключивших ДБО - Данные о лицах, подключивших ДБО'  # 78
    sheet['CB1'] = 'product_dbo_attr_dbo_document_copy#Продукты запроса - ДБО - Атрибуты продукта - Копии документов по ДБО (акты, сертификаты ключей) - Копии документов по ДБО (акты, сертификаты ключей)'  # 79

# Блок Продукты запроса ИБС
    sheet['CC1'] = 'product_ibs_attr_ibs_data#Продукты запроса - ИБС - Атрибуты продукта - Данные по ИБС (ячейкам) - Данные по ИБС (ячейкам)'  # 80
    sheet['CD1'] = 'product_ibs_attr_contract_file#Продукты запроса - ИБС - Атрибуты продукта - Досье по договорам (продуктовое досье) - Досье по договорам (продуктовое досье)'  # 81

# Блок Продукты запроса Инкассация
    sheet['CE1'] = 'product_encashment_attr_encashment_contract_copy#Продукты запроса - Инкассация - Атрибуты продукта - Данные по договору инкассации (копии документов) - Данные по договору инкассации (копии документов)'  # 82

# Блок Продукты запроса Карта
    sheet['CF1'] = 'product_card_attr_client_pers_data#Продукты запроса - Карта - Атрибуты продукта - Анкетные данные владельца счета'  # 83
    sheet['CG1'] = 'product_card_attr_opened_account_dept_data#Продукты запроса - Карта - Атрибуты продукта - Данные о наименовании/адресе ТП, где открыт счет'  # 84
    sheet['CH1'] = 'product_card_attr_account_file#Продукты запроса - Карта - Атрибуты продукта - Досье по счету/карте (копии документов)'  # 85
    sheet['CI1'] = 'product_card_attr_recipient_pers_data#Продукты запроса - Карта - Атрибуты продукта - Анкетные данные получателя перевода - Анкетные данные получателя перевода'  # 86
    sheet['CJ1'] = 'product_card_attr_bank_statement#Продукты запроса - Карта - Атрибуты продукта - Выписки - Выписки'  # 87
    sheet['CK1'] = 'product_card_attr_opened_account_pers_data#Продукты запроса - Карта - Атрибуты продукта - Данные о лицах, открывших счета - Данные о лицах, открывших счета'  # 88
    sheet['CL1'] = 'product_card_attr_atm_data#Продукты запроса - Карта - Атрибуты продукта - Данные по банкоматам/терминалам - Данные по банкоматам/терминалам'  # 89
    sheet['CM1'] = 'product_card_attr_sender_account_info#Продукты запроса - Карта - Атрибуты продукта - Данные по счетам/банковским картам получателей/отправителей - Данные по счетам/банковским картам получателей/отправителей'  # 90
    sheet['CN1'] = 'product_card_attr_recipient_account_info#Продукты запроса - Карта - Атрибуты продукта - Данные по счетам/банковским картам получателей - Данные по счетам/банковским картам получателей'  # 91
    sheet['CO1'] = 'product_card_attr_recipient_account_operations_info#Продукты запроса - Карта - Атрибуты продукта - Данные по счетам/банковским картам получателя перевода и операции по ним - Данные по счетам/банковским картам получателя перевода и операции по ним'  # 92
    sheet['CP1'] = 'product_card_attr_file#Продукты запроса - Карта - Атрибуты продукта - Картотека - Картотека'  # 93
    sheet['CQ1'] = 'product_card_attr_exist_account_card_cert#Продукты запроса - Карта - Атрибуты продукта - Сведения о наличии счетов и банковских карт - Сведения о наличии счетов и банковских карт (справки)'  # 94
    sheet['CR1'] = 'product_card_attr_phone_to_account_info#Продукты запроса - Карта - Атрибуты продукта - Сведения о подключении телефона к счетам/картам клиента из запроса - Сведения о подключении телефона к счетам/картам клиента из запроса'  # 95
    sheet['CS1'] = 'product_card_attr_account_balance_cert#Продукты запроса - Карта - Атрибуты продукта - Справки (сведения) об остатках - Справки (сведения) об остатках'  # 96
    sheet['CT1'] = 'product_card_attr_photo_video_office#Продукты запроса - Карта - Атрибуты продукта - Фото/видео из отделений Банка - Фото/видео из отделений Банка'  # 97
    sheet['CU1'] = 'product_card_attr_photo_video_atm#Продукты запроса - Карта - Атрибуты продукта - Фото/видео из устройств самообслуживания - Фото/видео из устройств самообслуживания'  # 98

# Блок Продукты запроса Кредит
    sheet['CV1'] = 'product_credit_attr_bank_statement#Продукты запроса - Кредит - Атрибуты продукта - Выписки - Выписки'  # 99
    sheet['CW1'] = 'product_credit_attr_credit_contract_data#Продукты запроса - Кредит - Атрибуты продукта - Данные по кредитным договорам - Данные по кредитным договорам'  # 100
    sheet['CX1'] = 'product_credit_attr_credit_debt_balance#Продукты запроса - Кредит - Атрибуты продукта - Данные по остатку долга по кредиту - Данные по остатку долга по кредиту'  # 101
    sheet['CY1'] = 'product_credit_attr_credit_paid_out_percent#Продукты запроса - Кредит - Атрибуты продукта - Данные по уплаченным процентам по кредиту - Данные по уплаченным процентам по кредиту'  # 102
    sheet['CZ1'] = 'product_credit_attr_contract_file#Продукты запроса - Кредит - Атрибуты продукта - Досье по договорам (продуктовое досье) - Досье по договорам (продуктовое досье)'  # 103
    sheet['DA1'] = 'product_credit_attr_debt_credit_calc#Продукты запроса - Кредит - Атрибуты продукта - Расчет задолженности по Кредитным договорам - Расчет задолженности по Кредитным договорам'  # 104
    sheet['DB1'] = 'product_credit_attr_loan_debt_cert#Продукты запроса - Кредит - Атрибуты продукта - Справка о ссудной задолженности - Справка о ссудной задолженности'  # 105

# Блок Продукты запроса ОМС
    sheet['DC1'] = 'product_oms_attr_oms_account_data#Продукты запроса - ОМС - Атрибуты продукта - Данные по счетам ОМС (покупка драг.металов) - Данные по счетам ОМС (покупка драг.металов)'  # 106

# Блок Продукты запроса Сберегательный сертификат
    sheet['DD1'] = 'product_saving_certificate_attr_saving_certificate_data#Продукты запроса - Сберегательный сертификат - Атрибуты продукта - Сведения о сберегательных сертификатах - Сведения о сберегательных сертификатах'  # 107

# Блок Продукты запроса Страховка
    sheet['DE1'] = 'product_insurance_attr_contract_file#Продукты запроса - Страховка - Атрибуты продукта - Досье по договорам (продуктовое досье) - Досье по договорам (продуктовое досье)'  # 108

# Блок Продукты запроса Счет
    sheet['DF1'] = 'product_account_attr_client_pers_data#Продукты запроса - Счет - Атрибуты продукта - Анкетные данные владельца счета'  # 109
    sheet['DG1'] = 'product_account_attr_account_control_person#Продукты запроса - Счет - Атрибуты продукта - Данные о лицах имеющих право распоряжения по счетам'  # 110
    sheet['DH1'] = 'product_account_attr_opened_account_dept_data#Продукты запроса - Счет - Атрибуты продукта - Данные о наименовании/адресе ТП, где открыт счет'  # 111
    sheet['DI1'] = 'product_account_attr_account_file#Продукты запроса - Счет - Атрибуты продукта - Досье по счету/карте (копии документов)'  # 112
    sheet['DJ1'] = 'product_account_attr_recipient_pers_data#Продукты запроса - Счет - Атрибуты продукта - Анкетные данные получателя перевода - Анкетные данные получателя перевода'  # 113
    sheet['DK1'] = 'product_account_attr_bank_statement_balance#Продукты запроса - Счет - Атрибуты продукта - Выписка с текущим остатком - Выписка с текущим остатком'  # 114
    sheet['DL1'] = 'product_account_attr_bank_statement#Продукты запроса - Счет - Атрибуты продукта - Выписки - Выписки'  # 115
    sheet['DM1'] = 'product_account_attr_opened_account_pers_data#Продукты запроса - Счет - Атрибуты продукта - Данные о лицах, открывших счета - Данные о лицах, открывших счета'  # 116
    sheet['DN1'] = 'product_account_attr_kop_stamp_pers_data#Продукты запроса - Счет - Атрибуты продукта - Данные о лицах, указанных в КОП и оттисков печати - Данные о лицах, указанных в КОП и оттисков печати'  # 117
    sheet['DO1'] = 'product_account_attr_gis_gmp_payment_confirm#Продукты запроса - Счет - Атрибуты продукта - Данные по УИП (ГИС ГМП) - подтверждение платежа - Данные по УИП (ГИС ГМП) - подтверждение платежа'  # 118
    sheet['DP1'] = 'product_account_attr_blocking_data#Продукты запроса - Счет - Атрибуты продукта - Данные по арестам/блокировкам - Данные по арестам/блокировкам'  # 119
    sheet['DQ1'] = 'product_account_attr_atm_data#Продукты запроса - Счет - Атрибуты продукта - Данные по банкоматам/терминалам - Данные по банкоматам/терминалам'  # 120
    sheet['DR1'] = 'product_account_attr_internet_transfer_data#Продукты запроса - Счет - Атрибуты продукта - Данные по интернет переводам - Данные по интернет переводам'  # 121
    sheet['DS1'] = 'product_account_attr_phones_data#Продукты запроса - Счет - Атрибуты продукта - Данные по сотовым телефонам - Данные по сотовым телефонам'  # 122
    sheet['DT1'] = 'product_account_attr_sender_account_info#Продукты запроса - Счет - Атрибуты продукта - Данные по счетам/банковским картам отправителей - Данные по счетам/банковским картам отправителей'  # 123
    sheet['DU1'] = 'product_account_attr_recipient_account_info#Продукты запроса - Счет - Атрибуты продукта - Данные по счетам/банковским картам получателей - Данные по счетам/банковским картам получателей'  # 124
    sheet['DV1'] = 'product_account_attr_recipient_account_operations_info#Продукты запроса - Счет - Атрибуты продукта - Данные по счетам/банковским картам получателя перевода и операции по ним - Данные по счетам/банковским картам получателя перевода и операции по ним'  # 125
    sheet['DW1'] = 'product_account_attr_file#Продукты запроса - Счет - Атрибуты продукта - Картотека - Картотека'  # 126
    sheet['DX1'] = 'product_account_attr_power_of_attorney_copy#Продукты запроса - Счет - Атрибуты продукта - Копии доверенностей - Копии доверенностей'  # 127
    sheet['DY1'] = 'product_account_attr_payment_document#Продукты запроса - Счет - Атрибуты продукта - Платежные документы - Платежные документы'  # 128
    sheet['DZ1'] = 'product_account_attr_exist_account_card_cert#Продукты запроса - Счет - Атрибуты продукта - Сведения о наличии счетов и банковских карт - Сведения о наличии счетов и банковских карт (справки)'  # 129
    sheet['EA1'] = 'product_account_attr_phone_to_account_info#Продукты запроса - Счет - Атрибуты продукта - Сведения о подключении телефона к счетам/картам клиента из запроса - Сведения о подключении телефона к счетам/картам клиента из запроса'  # 130
    sheet['EB1'] = 'product_account_attr_transaction_data#Продукты запроса - Счет - Атрибуты продукта - Сведения по транзакциям - Сведения по транзакциям'  # 131
    sheet['EC1'] = 'product_account_attr_account_balance_cert#Продукты запроса - Счет - Атрибуты продукта - Справки (сведения) об остатках - Справки (сведения) об остатках'  # 132
    sheet['ED1'] = 'product_account_attr_photo_video_atm#Продукты запроса - Счет - Атрибуты продукта - Фото/видео из устройств самообслуживания - Фото/видео из устройств самообслуживания'  # 133

# Блок Продукты запроса Ценные бумаги
    sheet['EE1'] = 'product_securities_attr_securities_data#Продукты запроса - Ценные бумаги - Атрибуты продукта - Данные по ценным бумагам - Данные по ценным бумагам'  # 134

def write_row_to_exel(
        row_number: int,
        json_body: Dict[str, Any],
        sheet: openpyxl.worksheet.worksheet.Worksheet
) -> None:
    """Записывает строки в exel файл заготовку."""
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

# Блок Продукты запроса Аккредитив
    # product_credit_letter_attr_credit_letter_data
    write_productType(row_number, json_body, sheet, 'PROD_CREDIT_LETTER', 'ATTR_CREDIT_LETTER_DATA', 59)
# Блок Продукты запроса Банковские гарантии
    # product_bank_guarantees_attr_bank_guarantees_data
    write_productType(row_number, json_body, sheet, 'PROD_BANK_GUARANTEES', 'ATTR_BANK_GUARANTEES_DATA', 60)
    # product_bank_guarantees_attr_contract_file
    write_productType(row_number, json_body, sheet, 'PROD_BANK_GUARANTEES', 'ATTR_CONTRACT_FILE', 61)
# Блок Продукты запроса Вексель
    # product_promissory_note_attr_promissory_note_data
    write_productType(row_number, json_body, sheet, 'PROD_PROMISSORY_NOTE', 'ATTR_PROMISSORY_NOTE_DATA', 62)
# Блок Продукты запроса Вклад
    # product_investment_attr_client_pers_data
    write_productType(row_number, json_body, sheet, 'PROD_INVESTMENT', 'ATTR_CLIENT_PERS_DATA', 63)
    # product_investment_attr_opened_account_dept_data
    write_productType(row_number, json_body, sheet, 'PROD_INVESTMENT', 'ATTR_OPENED_ACCOUNT_DEPT_DATA', 64)
    # product_investment_attr_deposit_paid_out_percent
    write_productType(row_number, json_body, sheet, 'PROD_INVESTMENT', 'ATTR_DEPOSIT_PAID_OUT_PERCENT', 65)
    # product_investment_attr_account_file
    write_productType(row_number, json_body, sheet, 'PROD_INVESTMENT', 'ATTR_ACCOUNT_FILE', 66)
    # product_investment_attr_bank_statement
    write_productType(row_number, json_body, sheet, 'PROD_INVESTMENT', 'ATTR_BANK_STATEMENT', 67)
    # product_investment_attr_opened_account_pers_data
    write_productType(row_number, json_body, sheet, 'PROD_INVESTMENT', 'ATTR_OPENED_ACCOUNT_PERS_DATA', 68)
    # product_investment_attr_file
    write_productType(row_number, json_body, sheet, 'PROD_INVESTMENT', 'ATTR_FILE', 69)
    # product_investment_attr_power_of_attorney_copy
    write_productType(row_number, json_body, sheet, 'PROD_INVESTMENT', 'ATTR_POWER_OF_ATTORNEY_COPY', 70)
    # product_investment_attr_exist_account_card_cert
    write_productType(row_number, json_body, sheet, 'PROD_INVESTMENT', 'ATTR_EXIST_ACCOUNT_CARD_CERT', 71)
    # product_investment_attr_phone_to_account_info
    write_productType(row_number, json_body, sheet, 'PROD_INVESTMENT', 'ATTR_PHONE_TO_ACCOUNT_INFO', 72)
    # product_investment_attr_account_balance_cert
    write_productType(row_number, json_body, sheet, 'PROD_INVESTMENT', 'ATTR_ACCOUNT_BALANCE_CERT', 73)
    # product_investment_attr_photo_video_office
    write_productType(row_number, json_body, sheet, 'PROD_INVESTMENT', 'ATTR_PHOTO_VIDEO_OFFICE', 74)
# Блок Продукты запроса ДБО
    # product_dbo_attr_dbo_connection_data
    write_productType(row_number, json_body, sheet, 'PROD_DBO', 'ATTR_DBO_CONNECTION_DATA', 75)
    # product_dbo_attr_ip_address
    write_productType(row_number, json_body, sheet, 'PROD_DBO', 'ATTR_IP_ADDRESS', 76)
    # product_dbo_attr_used_ip_address_pers_data
    write_productType(row_number, json_body, sheet, 'PROD_DBO', 'ATTR_USED_IP_ADDRESS_PERS_DATA', 77)
    # product_dbo_attr_dbo_connected_pers_data
    write_productType(row_number, json_body, sheet, 'PROD_DBO', 'ATTR_DBO_CONNECTED_PERS_DATA', 78)
    # product_dbo_attr_dbo_document_copy
    write_productType(row_number, json_body, sheet, 'PROD_DBO', 'ATTR_DBO_DOCUMENT_COPY', 79)
# Блок Продукты запроса ИБС
    # product_ibs_attr_ibs_data
    write_productType(row_number, json_body, sheet, 'PROD_IBS', 'ATTR_IBS_DATA', 80)
    # product_ibs_attr_contract_file
    write_productType(row_number, json_body, sheet, 'PROD_IBS', 'ATTR_CONTRACT_FILE', 81)
# Блок Продукты запроса Инкассация
    # product_ibs_attr_ibs_data
    write_productType(row_number, json_body, sheet, 'PROD_ENCASHMENT', 'ATTR_ENCASHMENT_CONTRACT_COPY', 82)
# Блок Продукты запроса Карта
    # product_card_attr_client_pers_data
    write_productType(row_number, json_body, sheet, 'PROD_CARD', 'ATTR_CLIENT_PERS_DATA', 83)
    # product_card_attr_opened_account_dept_data
    write_productType(row_number, json_body, sheet, 'PROD_CARD', 'ATTR_OPENED_ACCOUNT_DEPT_DATA', 84)
    # product_card_attr_account_file
    write_productType(row_number, json_body, sheet, 'PROD_CARD', 'ATTR_ACCOUNT_FILE', 85)
    # product_card_attr_recipient_pers_data
    write_productType(row_number, json_body, sheet, 'PROD_CARD', 'ATTR_RECIPIENT_PERS_DATA', 86)
    # product_card_attr_bank_statement
    write_productType(row_number, json_body, sheet, 'PROD_CARD', 'ATTR_BANK_STATEMENT', 87)
    # product_card_attr_opened_account_pers_data
    write_productType(row_number, json_body, sheet, 'PROD_CARD', 'ATTR_OPENED_ACCOUNT_PERS_DATA', 88)
    # product_card_attr_atm_data
    write_productType(row_number, json_body, sheet, 'PROD_CARD', 'ATTR_ATM_DATA', 89)
    # product_card_attr_sender_account_info
    write_productType(row_number, json_body, sheet, 'PROD_CARD', 'ATTR_SENDER_ACCOUNT_INFO', 90)
    # product_card_attr_recipient_account_info
    write_productType(row_number, json_body, sheet, 'PROD_CARD', 'ATTR_RECIPIENT_ACCOUNT_INFO', 91)
    # product_card_attr_recipient_account_operations_info
    write_productType(row_number, json_body, sheet, 'PROD_CARD', 'ATTR_RECIPIENT_ACCOUNT_OPERATION_INFO', 92)
    # product_card_attr_file
    write_productType(row_number, json_body, sheet, 'PROD_CARD', 'ATTR_FILE', 93)
    # product_card_attr_exist_account_card_cert
    write_productType(row_number, json_body, sheet, 'PROD_CARD', 'ATTR_EXIST_ACCOUNT_CARD_CERT', 94)
    # product_card_attr_phone_to_account_info
    write_productType(row_number, json_body, sheet, 'PROD_CARD', 'ATTR_PHONE_TO_ACCOUNT_INFO', 95)
    # product_card_attr_account_balance_cert
    write_productType(row_number, json_body, sheet, 'PROD_CARD', 'ATTR_ACCOUNT_BALANCE_CERT', 96)
    # product_card_attr_photo_video_office
    write_productType(row_number, json_body, sheet, 'PROD_CARD', 'ATTR_PHOTO_VIDEO_OFFICE', 97)
    # product_card_attr_photo_video_atm
    write_productType(row_number, json_body, sheet, 'PROD_CARD', 'ATTR_PHOTO_VIDEO_ATM', 98)
# Блок Продукты запроса Кредит
    # product_credit_attr_bank_statement
    write_productType(row_number, json_body, sheet, 'PROD_CREDIT', 'ATTR_BANK_STATEMENT', 99)
    # product_credit_attr_credit_contract_data
    write_productType(row_number, json_body, sheet, 'PROD_CREDIT', 'ATTR_CREDIT_CONTRACT_DATA', 100)
    # product_credit_attr_credit_debt_balance
    write_productType(row_number, json_body, sheet, 'PROD_CREDIT', 'ATTR_CREDIT_DEBT_BALANCE', 101)
    # product_credit_attr_credit_paid_out_percent
    write_productType(row_number, json_body, sheet, 'PROD_CREDIT', 'ATTR_CREDIT_PAID_OUT_PERCENT', 102)
    # product_credit_attr_contract_file
    write_productType(row_number, json_body, sheet, 'PROD_CREDIT', 'ATTR_CONTRACT_FILE', 103)
    # product_credit_attr_debt_credit_calc
    write_productType(row_number, json_body, sheet, 'PROD_CREDIT', 'ATTR_DEBT_CREDIT_CALC', 104)
    # product_credit_attr_loan_debt_cert
    write_productType(row_number, json_body, sheet, 'PROD_CREDIT', 'ATTR_LOAN_DEBT_CERT', 105)
# Блок Продукты запроса ОМС
    # product_oms_attr_oms_account_data
    write_productType(row_number, json_body, sheet, 'PROD_OMS', 'ATTR_OMS_ACCOUNT_DATA', 106)
# Блок Продукты запроса Сберегательный сертификат
    # product_saving_certificate_attr_saving_certificate_data
    write_productType(row_number, json_body, sheet, 'PROD_SAVING_CERTIFICATE', 'ATTR_SAVING_CERTIFICATE_DATA', 107)
# Блок Продукты запроса Страховка
    # product_insurance_attr_contract_file
    write_productType(row_number, json_body, sheet, 'PROD_INSURANCE', 'ATTR_CONTRACT_FILE', 108)
# Блок Продукты запроса Счет
    # product_account_attr_client_pers_data
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_CLIENT_PERS_DATA', 109)
    # product_account_attr_account_control_person
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_ACCOUNT_CONTROL_PERSON', 110)
    # product_account_attr_opened_account_dept_data
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_OPENED_ACCOUNT_DEPT_DATA', 111)
    # product_account_attr_account_file
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_ACCOUNT_FILE', 112)
    # product_account_attr_recipient_pers_data
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_RECIPIENT_PERS_DATA', 113)
    # product_account_attr_bank_statement_balance
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_BANK_STATEMENT_BALANCE', 114)
    # product_account_attr_bank_statement
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_BANK_STATEMENT', 115)
    # product_account_attr_opened_account_pers_data
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_OPENED_ACCOUNT_PERS_DATA', 116)
    # product_account_attr_kop_stamp_pers_data
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_KOP_STAMP_PERS_DATA', 117)
    # product_account_attr_gis_gmp_payment_confirm
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_GIS_GMP_PAYMENT_CONFIRM', 118)
    # product_account_attr_blocking_data
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_BLOCKING_DATA', 119)
    # product_account_attr_atm_data
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_ATM_DATA', 120)
    # product_account_attr_internet_transfer_data
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_INTERNET_TRANSFER_DATA', 121)
    # product_account_attr_phones_data
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_PHONES_DATA', 122)
    # product_account_attr_sender_account_info
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_SENDER_ACCOUNT_INFO', 123)
    # product_account_attr_recipient_account_info
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_RECIPIENT_ACCOUNT_INFO', 124)
    # product_account_attr_recipient_account_operations_info
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_RECIPIENT_ACCOUNT_OPERATION_INFO', 125)
    # product_account_attr_file
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_FILE', 126)
    # product_account_attr_power_of_attorney_copy
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_POWER_OF_ATTORNEY_COPY', 127)
    # product_account_attr_payment_document
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_PAYMENT_DOCUMENT', 128)
    # product_account_attr_exist_account_card_cert
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_EXIST_ACCOUNT_CARD_CERT', 129)
    # product_account_attr_phone_to_account_info
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_PHONE_TO_ACCOUNT_INFO', 130)
    # product_account_attr_transaction_data
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_TRANSACTION_DATA', 131)
    # product_account_attr_account_balance_cert
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_ACCOUNT_BALANCE_CERT', 132)
    # product_account_attr_photo_video_atm
    write_productType(row_number, json_body, sheet, 'PROD_ACCOUNT', 'ATTR_PHOTO_VIDEO_ATM', 133)
# Блок Продукты запроса Ценные бумаги
    # product_securities_attr_securities_data
    write_productType(row_number, json_body, sheet, 'PROD_SECURITIES', 'ATTR_SECURITIES_DATA', 134)


def write_productType(
        row_number: int,
        json_body: Dict[str, Any],
        sheet: openpyxl.worksheet.worksheet.Worksheet,
        product_name: str,
        attr_name: str,
        header_number: int):
    """Записывает продукт запроса из json в exel файл"""
    if json_body.get('targets') is not None:
        for target in json_body['targets']:
            if target.get('products') is not None:
                for product in target['products']:
                    if product['productType']['codeName'] == product_name:
                        if product.get('attributes') is not None:
                            for attribute in product['attributes']:
                                if attribute['attributeType']['codeName'] == attr_name:
                                    sheet[row_number][header_number].value = 'true'
                                else:
                                    continue
                    else:
                        continue
