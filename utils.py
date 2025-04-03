import openpyxl


class CheckError(Exception):
    def __init__(self, reason):
        """Возвращает кастомную ошибку, если файл не прошел проверку"""
        message = f"Файл не прошел проверку, причина - {reason}."
        super().__init__(message)


def open_exel(filepath: str) -> openpyxl.worksheet.worksheet.Worksheet:
    """Открывает файл, создает рабочую книгу для работы с данными"""
    book = openpyxl.open(filepath)
    sheet = book.active
    return sheet


def return_attribute_dict() -> dict:
    """Возвращает список атрибутов необходимый для подсчета статистики (учитываются только атрибуты, которые имеют хотя
    бы одну регулярку).
    """
    attribute_dict = {
        # Блок Документ
        'A1': 'file_name#имя файла',
        'B1': 'document_anticorr#Документ - Антикоррупционный',
        #'C1': 'document_outgoing_date#Документ - Исх. дата',
        #'D1': 'document_outgoing_number#Документ - Исх. номер',
        'E1': 'document_repeatedly#Документ - Повторно',

        # Блок Кто запрашивает
        'F1': 'applicant_signer_position#Кто запрашивает - Подписант - Должность',
        'G1': 'applicant_signer_name#Кто запрашивает - Подписант - ФИО',
        'H1': 'applicant_foiv_address#Кто запрашивает - ФОИВ - Адрес',
        'I1': 'applicant_foiv_name#Кто запрашивает - ФОИВ - Наименование',
        #'J1': 'applicant_foiv_phone#Кто запрашивает - ФОИВ - Телефон',

        # Блок Объект запроса ФЛ
        'K1': 'target_individual_name#Объект запроса - ФЛ - ФИО',
        'L1': 'target_individual_reg_address#Объект запроса - ФЛ - Адрес регистрации',
        'M1': 'target_individual_date_of_birth#Объект запроса - ФЛ - Дата рождения',
        'N1': 'target_individual_death_date#Объект запроса - ФЛ - Дата смерти',
        'O1': 'target_individual_inn#Объект запроса - ФЛ - ИНН',
        'P1': 'target_individual_birth_place#Объект запроса - ФЛ - Место рождения',
        'Q1': 'target_individual_phones#Объект запроса - ФЛ - Телефонные номера',
        'R1': 'target_individual_fact_address#Объект запроса - ФЛ - Фактический адрес',

        # Блок Объект запроса ФЛ Идентиф.документ
        'S1': 'target_individual_dul_date#Объект запроса - ФЛ - Идентификационный документ - 1_Дата выдачи',
        'T1': 'target_individual_dul_issue_code#Объект запроса - ФЛ - Идентификационный документ - 2_Код подразделения',
        'U1': 'target_individual_identitydocument_series#Объект запроса - ФЛ - Идентификационный документ - 3_Серия',
        'V1': 'target_individual_identitydocument_number#Объект запроса - ФЛ - Идентификационный документ - 4_Номер',
        'W1': 'target_individual_dul_org#Объект запроса - ФЛ - Идентификационный документ - 5_Орган выдачи',
        'X1': 'target_individual_dul_type#Объект запроса - ФЛ - Идентификационный документ - 6_Тип',

        # Блок Объект запроса ИП
        'Y1': 'target_individual_entrepreneur_name#Объект запроса - ИП - ФИО',
        'Z1': 'target_individual_entrepreneur_regadress#Объект запроса - ИП - Адрес регистрации',
        'AA1': 'target_individual_entrepreneur_date_of_birth#Объект запроса - ИП - Дата рождения',
        'AB1': 'target_individual_entrepreneur_inn#Объект запроса - ИП - ИНН',
        'AC1': 'target_individual_entrepreneur_phones#Объект запроса - ИП - Телефонные номера',
        'AD1': 'target_individual_entrepreneur_factadress#Объект запроса - ИП - Фактический адрес',

        # Блок Объект запроса ЮЛ
        'AE1': 'target_company_name#Объект запроса - ЮЛ - Наименование',
        'AF1': 'target_company_regadress#Объект запроса - ЮЛ - Адрес регистрации',
        'AG1': 'target_company_taxpayer_number#Объект запроса - ЮЛ - ИНН',
        'AH1': 'target_company_kpp#Объект запроса - ЮЛ - КПП',
        'AI1': 'target_company_ogrn#Объект запроса - ЮЛ - ОГРН',
        'AJ1': 'target_company_phones#Объект запроса - ЮЛ - Телефонные номера',
        'AK1': 'target_company_factadress#Объект запроса - ЮЛ - Фактический адрес',

        # Блок Основание для запроса
        'AL1': 'basis_date_court_act#Основание для запроса - Дата основания запроса',
        'AM1': 'basis_custom_offence_value#Основание для запроса - Административное правонарушение - Номер дела об АП',
        'AN1': 'basis_arbitration_case_value#Основание для запроса - Арбитражное дело - Номер арбитражного дела',
        'AO1': 'basis_civil_case_value#Основание для запроса - Гражданское дело - Номер гражданского дела',
        'AP1': 'basis_position_value#Основание для запроса - Должность - Название должности',
        'AQ1': 'basis_inforcement_proceeding_value#Основание для запроса - Исполнительное производство - '
               'Номер исполнительного производства',
        'AR1': 'basis_preliminary_inquiry_value#Основание для запроса - Материал проверки - Номер материала проверки',
        'AS1': 'basis_inheritance_case_value#Основание для запроса - Наследственное дело - Номер наследственного дела',
        'AT1': 'basis_court_order_value#Основание для запроса - Постановление суда - Номер постановления суда',
        'AU1': 'basis_claim_value#Основание для запроса - Претензия/жалоба - Номер претензии/жалобы',
        'AV1': 'basis_criminal_case_value#Основание для запроса - Уголовное дело - Номер уголовного дела',
        'AW1': 'basis_legal_clause_lc_fz311#Основание для запроса - Статьи - 311-ФЗ',
        'AX1': 'basis_legal_clause_lc_a15_notary#Основание для запроса - Статьи - ст. 15 "О нотариате"',
        'AY1': 'basis_legal_clause_lc_a23_fz173#Основание для запроса - Статьи - ст. 23 (173-ФЗ)',

        # Блок Особые условия
        #'AZ1': 'conditions_give_out_on_purpose#Особые условия - Выдать нарочно',
        #'BA1': 'conditions_whom_send_response#Особые условия - Кому направить ответ',
        #'BB1': 'conditions_where_send_response#Особые условия - Куда направить ответ',
        #'BC1': 'conditions_in_digital_format#Особые условия - Предоставить в электронном виде',
        #'BD1': 'conditions_response_deadline#Особые условия - Предоставить до',
        #'BE1': 'conditions_in_excel_format#Особые условия - Формат сведений Excel',

        # Блок Продукты запроса Значения продукта
        'BF1': 'product_card_attr_number#Продукты запроса - Карта - Значения продукта - Номер карты',
        'BG1': 'product_account_number#Продукты запроса - Счет - Значения продукта - Номер счета',

        # Блок Продукты запроса Аккредитив
        'BH1': 'product_credit_letter_attr_credit_letter_data#Продукты запроса - Аккредитив - Атрибуты продукта - '
               'Данные по аккредитивам - Данные по аккредитивам',

        # Блок Продукты запроса Банковские гарантии
        'BI1': 'product_bank_guarantees_attr_bank_guarantees_data#Продукты запроса - Банковские гарантии - '
               'Атрибуты продукта - Данные по банковским гарантиям - Данные по банковским гарантиям',
        'BJ1': 'product_bank_guarantees_attr_contract_file#Продукты запроса - Банковские гарантии - '
               'Атрибуты продукта - Досье по договорам (продуктовое досье) - Досье по договорам (продуктовое досье)',

        # Блок Продукты запроса Вексель
        'BK1': 'product_promissory_note_attr_promissory_note_data#Продукты запроса - Вексель - '
               'Атрибуты продукта - Данные по векселям - Данные по векселям',

        # Блок Продукты запроса Вклад
        'BL1': 'product_investment_attr_client_pers_data#Продукты запроса - Вклад - '
               'Атрибуты продукта - Анкетные данные владельца счета',
        'BM1': 'product_investment_attr_opened_account_dept_data#Продукты запроса - Вклад - '
               'Атрибуты продукта - Данные о наименовании/адресе ТП, где открыт счет',
        'BN1': 'product_investment_attr_deposit_paid_out_percent#Продукты запроса - Вклад - '
               'Атрибуты продукта - Данные по уплаченным процентам по вкладу (депозиту)',
        'BO1': 'product_investment_attr_account_file#Продукты запроса - Вклад - '
               'Атрибуты продукта - Досье по счету/карте (копии документов)',
        'BP1': 'product_investment_attr_bank_statement#Продукты запроса - Вклад - '
               'Атрибуты продукта - Выписки - Выписки',
        'BQ1': 'product_investment_attr_opened_account_pers_data#Продукты запроса - Вклад - '
               'Атрибуты продукта - Данные о лицах, открывших счета - Данные о лицах, открывших счета',
        'BR1': 'product_investment_attr_file#Продукты запроса - Вклад - '
               'Атрибуты продукта - Картотека - Картотека',
        'BS1': 'product_investment_attr_power_of_attorney_copy#Продукты запроса - Вклад - '
               'Атрибуты продукта - Копии доверенностей - Копии доверенностей',
        'BT1': 'product_investment_attr_exist_account_card_cert#Продукты запроса - Вклад - '
               'Атрибуты продукта - Сведения о наличии счетов и банковских карт - '
               'Сведения о наличии счетов и банковских карт (справки)',
        'BU1': 'product_investment_attr_phone_to_account_info#Продукты запроса - Вклад - '
               'Атрибуты продукта - Сведения о подключении телефона к счетам/картам клиента из запроса - '
               'Сведения о подключении телефона к счетам/картам клиента из запроса',
        'BV1': 'product_investment_attr_account_balance_cert#Продукты запроса - Вклад - '
               'Атрибуты продукта - Справки (сведения) об остатках - Справки (сведения) об остатках',
        'BW1': 'product_investment_attr_photo_video_office#Продукты запроса - Вклад - '
               'Атрибуты продукта - Фото/видео из отделений Банка - Фото/видео из отделений Банка',

        # Блок Продукты запроса ДБО
        'BX1': 'product_dbo_attr_client_mobile_dbo#Продукты запроса - ДБО - '
               'Атрибуты продукта - Информация о подключении услуги Мобильный банк '
               '(к какому номеру подключена система "Клиет-банк")',
        'BY1': 'product_dbo_attr_ip_address#Продукты запроса - ДБО - '
               'Атрибуты продукта - IP адреса/Log файлы - IP адреса/Log файлы',
        'BZ1': 'product_dbo_attr_used_ip_address_pers_data#Продукты запроса - ДБО - '
               'Атрибуты продукта - Данные о лицах использовавших ip-адреса - Данные о лицах использовавших ip-адреса',
        'CA1': 'product_dbo_attr_dbo_connected_pers_data#Продукты запроса - ДБО - '
               'Атрибуты продукта - Данные о лицах подключивших ДБО - Данные о лицах, подключивших ДБО',
        'CB1': 'product_dbo_attr_dbo_document_copy#Продукты запроса - ДБО - '
               'Атрибуты продукта - Копии документов по ДБО (акты, сертификаты ключей) - '
               'Копии документов по ДБО (акты, сертификаты ключей)',

        # Блок Продукты запроса ИБС
        'CC1': 'product_ibs_attr_ibs_data#Продукты запроса - ИБС - '
               'Атрибуты продукта - Данные по ИБС (ячейкам) - Данные по ИБС (ячейкам)',
        'CD1': 'product_ibs_attr_contract_file#Продукты запроса - ИБС - '
               'Атрибуты продукта - Досье по договорам (продуктовое досье) - Досье по договорам (продуктовое досье)',

        # Блок Продукты запроса Инкассация
        'CE1': 'product_encashment_attr_encashment_contract_copy#Продукты запроса - Инкассация - '
               'Атрибуты продукта - Данные по договору инкассации (копии документов) - '
               'Данные по договору инкассации (копии документов)',

        # Блок Продукты запроса Карта
        'CF1': 'product_card_attr_client_pers_data#Продукты запроса - Карта - '
               'Атрибуты продукта - Анкетные данные владельца счета',
        'CG1': 'product_card_attr_opened_account_dept_data#Продукты запроса - Карта - '
               'Атрибуты продукта - Данные о наименовании/адресе ТП, где открыт счет',
        'CH1': 'product_card_attr_account_file#Продукты запроса - Карта - '
               'Атрибуты продукта - Досье по счету/карте (копии документов)',
        'CI1': 'product_card_attr_recipient_pers_data#Продукты запроса - Карта - '
               'Атрибуты продукта - Анкетные данные получателя перевода - Анкетные данные получателя перевода',
        'CJ1': 'product_card_attr_bank_statement#Продукты запроса - Карта - '
               'Атрибуты продукта - Выписки - Выписки',
        'CK1': 'product_card_attr_opened_account_pers_data#Продукты запроса - Карта - '
               'Атрибуты продукта - Данные о лицах, открывших счета - Данные о лицах, открывших счета',
        'CL1': 'product_card_attr_atm_data#Продукты запроса - Карта - '
               'Атрибуты продукта - Данные по банкоматам/терминалам - Данные по банкоматам/терминалам',
        'CM1': 'product_card_attr_sender_account_info#Продукты запроса - Карта - '
               'Атрибуты продукта - Данные по счетам/банковским картам отправителей - '
               'Данные по счетам/банковским картам получателей/отправителей',
        'CN1': 'product_card_attr_recipient_account_info#Продукты запроса - Карта - '
               'Атрибуты продукта - Данные по счетам/банковским картам получателей - '
               'Данные по счетам/банковским картам получателей',
        'CO1': 'product_card_attr_recipient_account_operations_info#Продукты запроса - Карта - '
               'Атрибуты продукта - Данные по счетам/банковским картам получателя перевода и операции по ним - '
               'Данные по счетам/банковским картам получателя перевода и операции по ним',
        'CP1': 'product_card_attr_file#Продукты запроса - Карта - '
               'Атрибуты продукта - Картотека - Картотека',
        'CQ1': 'product_card_attr_exist_account_card_cert#Продукты запроса - Карта - '
               'Атрибуты продукта - Сведения о наличии счетов и банковских карт - '
               'Сведения о наличии счетов и банковских карт (справки)',
        'CR1': 'product_card_attr_phone_to_account_info#Продукты запроса - Карта - '
               'Атрибуты продукта - Сведения о подключении телефона к счетам/картам клиента из запроса - '
               'Сведения о подключении телефона к счетам/картам клиента из запроса',
        'CS1': 'product_card_attr_account_balance_cert#Продукты запроса - Карта - '
               'Атрибуты продукта - Справки (сведения) об остатках - Справки (сведения) об остатках',
        'CT1': 'product_card_attr_photo_video_office#Продукты запроса - Карта - '
               'Атрибуты продукта - Фото/видео из отделений Банка - Фото/видео из отделений Банка',
        'CU1': 'product_card_attr_photo_video_atm#Продукты запроса - Карта - '
               'Атрибуты продукта - Фото/видео из устройств самообслуживания - '
               'Фото/видео из устройств самообслуживания',

        # Блок Продукты запроса Кредит
        'CV1': 'product_credit_attr_bank_statement#Продукты запроса - Кредит - '
               'Атрибуты продукта - Выписки - Выписки',
        'CW1': 'product_credit_attr_credit_contract_data#Продукты запроса - Кредит - '
               'Атрибуты продукта - Данные по кредитным договорам - Данные по кредитным договорам',
        'CX1': 'product_credit_attr_credit_debt_balance#Продукты запроса - Кредит - '
               'Атрибуты продукта - Данные по остатку долга по кредиту - Данные по остатку долга по кредиту',
        'CY1': 'product_credit_attr_credit_paid_out_percent#Продукты запроса - Кредит - '
               'Атрибуты продукта - Данные по уплаченным процентам по кредиту - '
               'Данные по уплаченным процентам по кредиту',
        'CZ1': 'product_credit_attr_contract_file#Продукты запроса - Кредит - '
               'Атрибуты продукта - Досье по договорам (продуктовое досье) - Досье по договорам (продуктовое досье)',
        'DA1': 'product_credit_attr_debt_credit_calc#Продукты запроса - Кредит - '
               'Атрибуты продукта - Расчет задолженности по Кредитным договорам - '
               'Расчет задолженности по Кредитным договорам',
        'DB1': 'product_credit_attr_loan_debt_cert#Продукты запроса - Кредит - '
               'Атрибуты продукта - Справка о ссудной задолженности - Справка о ссудной задолженности',

        # Блок Продукты запроса ОМС
        'DC1': 'product_oms_attr_oms_account_data#Продукты запроса - ОМС - '
               'Атрибуты продукта - Данные по счетам ОМС (покупка драг.металов) - '
               'Данные по счетам ОМС (покупка драг.металов)',

        # Блок Продукты запроса Сберегательный сертификат
        'DD1': 'product_saving_certificate_attr_saving_certificate_data#Продукты запроса - '
               'Сберегательный сертификат - Атрибуты продукта - '
               'Сведения о сберегательных сертификатах - Сведения о сберегательных сертификатах',

        # Блок Продукты запроса Страховка
        'DE1': 'product_insurance_attr_contract_file#Продукты запроса - Страховка - '
               'Атрибуты продукта - Досье по договорам (продуктовое досье) - Досье по договорам (продуктовое досье)',

        # Блок Продукты запроса Счет
        'DF1': 'product_account_attr_client_pers_data#Продукты запроса - Счет - '
               'Атрибуты продукта - Анкетные данные владельца счета',
        'DG1': 'product_account_attr_account_control_person#Продукты запроса - Счет - '
               'Атрибуты продукта - Данные о лицах имеющих право распоряжения по счетам',
        'DH1': 'product_account_attr_opened_account_dept_data#Продукты запроса - Счет - '
               'Атрибуты продукта - Данные о наименовании/адресе ТП, где открыт счет',
        'DI1': 'product_account_attr_account_file#Продукты запроса - Счет - '
               'Атрибуты продукта - Досье по счету/карте (копии документов)',
        'DJ1': 'product_account_attr_recipient_pers_data#Продукты запроса - Счет - '
               'Атрибуты продукта - Анкетные данные получателя перевода - Анкетные данные получателя перевода',
        'DK1': 'product_account_attr_bank_statement_balance#Продукты запроса - Счет - '
               'Атрибуты продукта - Выписка с текущим остатком - Выписка с текущим остатком',
        'DL1': 'product_account_attr_bank_statement#Продукты запроса - Счет - '
               'Атрибуты продукта - Выписки - Выписки',
        'DM1': 'product_account_attr_opened_account_pers_data#Продукты запроса - Счет - '
               'Атрибуты продукта - Данные о лицах, открывших счета - Данные о лицах, открывших счета',
        'DN1': 'product_account_attr_kop_stamp_pers_data#Продукты запроса - Счет - '
               'Атрибуты продукта - Данные о лицах, указанных в КОП и оттисков печати - '
               'Данные о лицах, указанных в КОП и оттисков печати',
        'DO1': 'product_account_attr_gis_gmp_payment_confirm#Продукты запроса - Счет - '
               'Атрибуты продукта - Данные по УИП (ГИС ГМП) - подтверждение платежа - '
               'Данные по УИП (ГИС ГМП) - подтверждение платежа',
        'DP1': 'product_account_attr_blocking_data#Продукты запроса - Счет - '
               'Атрибуты продукта - Данные по арестам/блокировкам - Данные по арестам/блокировкам',
        'DQ1': 'product_account_attr_atm_data#Продукты запроса - Счет - '
               'Атрибуты продукта - Данные по банкоматам/терминалам - Данные по банкоматам/терминалам',
        'DR1': 'product_account_attr_internet_transfer_data#Продукты запроса - Счет - '
               'Атрибуты продукта - Данные по интернет переводам - Данные по интернет переводам',
        'DS1': 'product_account_attr_phones_data#Продукты запроса - Счет - '
               'Атрибуты продукта - Данные по сотовым телефонам - Данные по сотовым телефонам',
        'DT1': 'product_account_attr_sender_account_info#Продукты запроса - Счет - '
               'Атрибуты продукта - Данные по счетам/банковским картам отправителей - '
               'Данные по счетам/банковским картам отправителей',
        'DU1': 'product_account_attr_recipient_account_info#Продукты запроса - Счет - '
               'Атрибуты продукта - Данные по счетам/банковским картам получателей - '
               'Данные по счетам/банковским картам получателей',
        'DV1': 'product_account_attr_recipient_account_operations_info#Продукты запроса - Счет - '
               'Атрибуты продукта - Данные по счетам/банковским картам получателя перевода и операции по ним - '
               'Данные по счетам/банковским картам получателя перевода и операции по ним',
        'DW1': 'product_account_attr_file#Продукты запроса - Счет - '
               'Атрибуты продукта - Картотека - Картотека',
        'DX1': 'product_account_attr_power_of_attorney_copy#Продукты запроса - Счет - '
               'Атрибуты продукта - Копии доверенностей - Копии доверенностей',
        'DY1': 'product_account_attr_payment_document#Продукты запроса - Счет - '
               'Атрибуты продукта - Платежные документы - Платежные документы',
        'DZ1': 'product_account_attr_exist_account_card_cert#Продукты запроса - Счет - '
               'Атрибуты продукта - Сведения о наличии счетов и банковских карт - '
               'Сведения о наличии счетов и банковских карт (справки)',
        'EA1': 'product_account_attr_phone_to_account_info#Продукты запроса - Счет - '
               'Атрибуты продукта - Сведения о подключении телефона к счетам/картам клиента из запроса - '
               'Сведения о подключении телефона к счетам/картам клиента из запроса',
        'EB1': 'product_account_attr_transaction_data#Продукты запроса - Счет - '
               'Атрибуты продукта - Сведения по транзакциям - Сведения по транзакциям',
        'EC1': 'product_account_attr_account_balance_cert#Продукты запроса - Счет - '
               'Атрибуты продукта - Справки (сведения) об остатках - Справки (сведения) об остатках',
        'ED1': 'product_account_attr_photo_video_atm#Продукты запроса - Счет - '
               'Атрибуты продукта - Фото/видео из устройств самообслуживания - '
               'Фото/видео из устройств самообслуживания',

        # Блок Продукты запроса Ценные бумаги
        'EE1': 'product_securities_attr_securities_data#Продукты запроса - Ценные бумаги - '
               'Атрибуты продукта - Данные по ценным бумагам - Данные по ценным бумагам'
    }
    return attribute_dict

def return_product_dict() -> dict:
    """Возвращает список продуктов запроса"""
    product_dict = {
        # Блок Продукты запроса Аккредитив
        'product_credit_letter_attr_credit_letter_data': ['PROD_CREDIT_LETTER', 'ATTR_CREDIT_LETTER_DATA', 59],

        # Блок Продукты запроса Банковские гарантии
        'product_bank_guarantees_attr_bank_guarantees_data':
            ['PROD_BANK_GUARANTEES', 'ATTR_BANK_GUARANTEES_DATA', 60],
        'product_bank_guarantees_attr_contract_file': ['PROD_BANK_GUARANTEES', 'ATTR_CONTRACT_FILE', 61],

        # Блок Продукты запроса Вексель
        'product_promissory_note_attr_promissory_note_data':
            ['PROD_PROMISSORY_NOTE', 'ATTR_PROMISSORY_NOTE_DATA', 62],

        # Блок Продукты запроса Вклад
        'product_investment_attr_client_pers_data': ['PROD_INVESTMENT', 'ATTR_CLIENT_PERS_DATA', 63],
        'product_investment_attr_opened_account_dept_data':
            ['PROD_INVESTMENT', 'ATTR_OPENED_ACCOUNT_DEPT_DATA', 64],
        'product_investment_attr_deposit_paid_out_percent':
            ['PROD_INVESTMENT', 'ATTR_DEPOSIT_PAID_OUT_PERCENT', 65],
        'product_investment_attr_account_file': ['PROD_INVESTMENT', 'ATTR_ACCOUNT_FILE', 66],
        'product_investment_attr_bank_statement': ['PROD_INVESTMENT', 'ATTR_BANK_STATEMENT', 67],
        'product_investment_attr_opened_account_pers_data':
            ['PROD_INVESTMENT', 'ATTR_OPENED_ACCOUNT_PERS_DATA', 68],
        'product_investment_attr_file': ['PROD_INVESTMENT', 'ATTR_FILE', 69],
        'product_investment_attr_power_of_attorney_copy': ['PROD_INVESTMENT', 'ATTR_POWER_OF_ATTORNEY_COPY', 70],
        'product_investment_attr_exist_account_card_cert': ['PROD_INVESTMENT', 'ATTR_EXIST_ACCOUNT_CARD_CERT', 71],
        'product_investment_attr_phone_to_account_info': ['PROD_INVESTMENT', 'ATTR_PHONE_TO_ACCOUNT_INFO', 72],
        'product_investment_attr_account_balance_cert': ['PROD_INVESTMENT', 'ATTR_ACCOUNT_BALANCE_CERT', 73],
        'product_investment_attr_photo_video_office': ['PROD_INVESTMENT', 'ATTR_PHOTO_VIDEO_OFFICE', 74],

        # Блок Продукты запроса ДБО
        'product_dbo_attr_dbo_connection_data': ['PROD_DBO', 'ATTR_DBO_CONNECTION_DATA', 75],
        'product_dbo_attr_ip_address': ['PROD_DBO', 'ATTR_IP_ADDRESS', 76],
        'product_dbo_attr_used_ip_address_pers_data': ['PROD_DBO', 'ATTR_USED_IP_ADDRESS_PERS_DATA', 77],
        'product_dbo_attr_dbo_connected_pers_data': ['PROD_DBO', 'ATTR_DBO_CONNECTED_PERS_DATA', 78],
        'product_dbo_attr_dbo_document_copy': ['PROD_DBO', 'ATTR_DBO_DOCUMENT_COPY', 79],

        # Блок Продукты запроса ИБС
        'product_ibs_attr_ibs_data': ['PROD_IBS', 'ATTR_IBS_DATA', 80],
        'product_ibs_attr_contract_file': ['PROD_IBS', 'ATTR_CONTRACT_FILE', 81],

        # Блок Продукты запроса Инкассация
        'product_encashment_attr_encashment_contract_copy':
            ['PROD_ENCASHMENT', 'ATTR_ENCASHMENT_CONTRACT_COPY', 82],

        # Блок Продукты запроса Карта
        'product_card_attr_client_pers_data': ['PROD_CARD', 'ATTR_CLIENT_PERS_DATA', 83],
        'product_card_attr_opened_account_dept_data': ['PROD_CARD', 'ATTR_OPENED_ACCOUNT_DEPT_DATA', 84],
        'product_card_attr_account_file': ['PROD_CARD', 'ATTR_ACCOUNT_FILE', 85],
        'product_card_attr_recipient_pers_data': ['PROD_CARD', 'ATTR_RECIPIENT_PERS_DATA', 86],
        'product_card_attr_bank_statement': ['PROD_CARD', 'ATTR_BANK_STATEMENT', 87],
        'product_card_attr_opened_account_pers_data': ['PROD_CARD', 'ATTR_OPENED_ACCOUNT_PERS_DATA', 88],
        'product_card_attr_atm_data': ['PROD_CARD', 'ATTR_ATM_DATA', 89],
        'product_card_attr_sender_account_info': ['PROD_CARD', 'ATTR_SENDER_ACCOUNT_INFO', 90],
        'product_card_attr_recipient_account_info': ['PROD_CARD', 'ATTR_RECIPIENT_ACCOUNT_INFO', 91],
        'product_card_attr_recipient_account_operations_info':
            ['PROD_CARD', 'ATTR_RECIPIENT_ACCOUNT_OPERATION_INFO', 92],
        'product_card_attr_file': ['PROD_CARD', 'ATTR_FILE', 93],
        'product_card_attr_exist_account_card_cert': ['PROD_CARD', 'ATTR_EXIST_ACCOUNT_CARD_CERT', 94],
        'product_card_attr_phone_to_account_info': ['PROD_CARD', 'ATTR_PHONE_TO_ACCOUNT_INFO', 95],
        'product_card_attr_account_balance_cert': ['PROD_CARD', 'ATTR_ACCOUNT_BALANCE_CERT', 96],
        'product_card_attr_photo_video_office': ['PROD_CARD', 'ATTR_PHOTO_VIDEO_OFFICE', 97],
        'product_card_attr_photo_video_atm': ['PROD_CARD', 'ATTR_PHOTO_VIDEO_ATM', 98],

        # Блок Продукты запроса Кредит
        'product_credit_attr_bank_statement': ['PROD_CREDIT', 'ATTR_BANK_STATEMENT', 99],
        'product_credit_attr_credit_contract_data': ['PROD_CREDIT', 'ATTR_CREDIT_CONTRACT_DATA', 100],
        'product_credit_attr_credit_debt_balance': ['PROD_CREDIT', 'ATTR_CREDIT_DEBT_BALANCE', 101],
        'product_credit_attr_credit_paid_out_percent': ['PROD_CREDIT', 'ATTR_CREDIT_PAID_OUT_PERCENT', 102],
        'product_credit_attr_contract_file': ['PROD_CREDIT', 'ATTR_CONTRACT_FILE', 103],
        'product_credit_attr_debt_credit_calc': ['PROD_CREDIT', 'ATTR_DEBT_CREDIT_CALC', 104],
        'product_credit_attr_loan_debt_cert': ['PROD_CREDIT', 'ATTR_LOAN_DEBT_CERT', 105],

        # Блок Продукты запроса ОМС
        'product_oms_attr_oms_account_data': ['PROD_OMS', 'ATTR_OMS_ACCOUNT_DATA', 106],

        # Блок Продукты запроса Сберегательный сертификат
        'product_saving_certificate_attr_saving_certificate_data':
            ['PROD_SAVING_CERTIFICATE', 'ATTR_SAVING_CERTIFICATE_DATA', 107],

        # Блок Продукты запроса Страховка
        'product_insurance_attr_contract_file': ['PROD_INSURANCE', 'ATTR_CONTRACT_FILE', 108],

        # Блок Продукты запроса Счет
        'product_account_attr_client_pers_data': ['PROD_ACCOUNT', 'ATTR_CLIENT_PERS_DATA', 109],
        'product_account_attr_account_control_person': ['PROD_ACCOUNT', 'ATTR_ACCOUNT_CONTROL_PERSON', 110],
        'product_account_attr_opened_account_dept_data': ['PROD_ACCOUNT', 'ATTR_OPENED_ACCOUNT_DEPT_DATA', 111],
        'product_account_attr_account_file': ['PROD_ACCOUNT', 'ATTR_ACCOUNT_FILE', 112],
        'product_account_attr_recipient_pers_data': ['PROD_ACCOUNT', 'ATTR_RECIPIENT_PERS_DATA', 113],
        'product_account_attr_bank_statement_balance': ['PROD_ACCOUNT', 'ATTR_BANK_STATEMENT_BALANCE', 114],
        'product_account_attr_bank_statement': ['PROD_ACCOUNT', 'ATTR_BANK_STATEMENT', 115],
        'product_account_attr_opened_account_pers_data': ['PROD_ACCOUNT', 'ATTR_OPENED_ACCOUNT_PERS_DATA', 116],
        'product_account_attr_kop_stamp_pers_data': ['PROD_ACCOUNT', 'ATTR_KOP_STAMP_PERS_DATA', 117],
        'product_account_attr_gis_gmp_payment_confirm': ['PROD_ACCOUNT', 'ATTR_GIS_GMP_PAYMENT_CONFIRM', 118],
        'product_account_attr_blocking_data': ['PROD_ACCOUNT', 'ATTR_BLOCKING_DATA', 119],
        'product_account_attr_atm_data': ['PROD_ACCOUNT', 'ATTR_ATM_DATA', 120],
        'product_account_attr_internet_transfer_data': ['PROD_ACCOUNT', 'ATTR_INTERNET_TRANSFER_DATA', 121],
        'product_account_attr_phones_data': ['PROD_ACCOUNT', 'ATTR_PHONES_DATA', 122],
        'product_account_attr_sender_account_info': ['PROD_ACCOUNT', 'ATTR_SENDER_ACCOUNT_INFO', 123],
        'product_account_attr_recipient_account_info': ['PROD_ACCOUNT', 'ATTR_RECIPIENT_ACCOUNT_INFO', 124],
        'product_account_attr_recipient_account_operations_info':
            ['PROD_ACCOUNT', 'ATTR_RECIPIENT_ACCOUNT_OPERATION_INFO', 125],
        'product_account_attr_file': ['PROD_ACCOUNT', 'ATTR_FILE', 126],
        'product_account_attr_power_of_attorney_copy': ['PROD_ACCOUNT', 'ATTR_POWER_OF_ATTORNEY_COPY', 127],
        'product_account_attr_payment_document': ['PROD_ACCOUNT', 'ATTR_PAYMENT_DOCUMENT', 128],
        'product_account_attr_exist_account_card_cert': ['PROD_ACCOUNT', 'ATTR_EXIST_ACCOUNT_CARD_CERT', 129],
        'product_account_attr_phone_to_account_info': ['PROD_ACCOUNT', 'ATTR_PHONE_TO_ACCOUNT_INFO', 130],
        'product_account_attr_transaction_data': ['PROD_ACCOUNT', 'ATTR_TRANSACTION_DATA', 131],
        'product_account_attr_account_balance_cert': ['PROD_ACCOUNT', 'ATTR_ACCOUNT_BALANCE_CERT', 132],
        'product_account_attr_photo_video_atm': ['PROD_ACCOUNT', 'ATTR_PHOTO_VIDEO_ATM', 133],

        # Блок Продукты запроса Ценные бумаги
        'product_securities_attr_securities_data': ['PROD_SECURITIES', 'ATTR_SECURITIES_DATA', 134],
    }
    return product_dict