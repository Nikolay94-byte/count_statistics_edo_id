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
