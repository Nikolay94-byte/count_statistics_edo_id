import openpyxl


def open_excel(filepath: str) -> openpyxl.worksheet.worksheet.Worksheet:
    """Открывает файл, создает рабочую книгу для работы с данными"""
    book = openpyxl.open(filepath)
    sheet = book.active
    return sheet


def convert_file_attributes_to_dict(file_path) -> dict[str, str]:
    """Возвращает словарь атрибутов, необходимый для подсчета статистики."""
    result_dict = {}

    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                if not line or line.startswith('#'):
                    continue
                if ':' in line:
                    key_part, value_part = line.split(':', 1)
                    result_dict[key_part] = value_part

    except FileNotFoundError:
        raise FileNotFoundError(f"Ошибка: Файл '{file_path}' не найден.")
    except Exception as error:
        raise Exception(f"Произошла ошибка при чтении файла: {error}")

    return result_dict


def convert_file_products_to_dict(file_path) -> dict[str, list[str | int]]:
    """Возвращает словарь продуктов запроса"""
    result_dict = {}

    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                line = line.strip()
                if not line or line.startswith('#'):
                    continue

                key, value = line.split(':', 1)
                key = key.strip()
                value = value.strip()

                items = value.split(',')
                processed_items = []
                for item in items:
                    item = item.strip()
                    if item.isdigit():
                        processed_items.append(int(item))
                    else:
                        processed_items.append(item)

                result_dict[key] = processed_items

    except FileNotFoundError:
        raise FileNotFoundError(f"Ошибка: Файл '{file_path}' не найден.")
    except Exception as error:
        raise Exception(f"Произошла ошибка при чтении файла: {error}")

    return result_dict
