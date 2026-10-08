import io
import openpyxl


def read_excel_from_bytes(content: bytes) -> list[dict]:
    """
    Считывает Excel-файл (.xlsx) из байтов и возвращает список словарей получателей.
    Первая строка считается заголовком.
    Поддерживает произвольные столбцы и регистронезависимые имена.
    """
    if not content:
        return []

    workbook = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    sheet = workbook.active
    if not sheet:
        return []

    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return []

    header_row = rows[0]
    if not header_row:
        return []

    # Нормализуем заголовки
    headers = []
    for idx, col in enumerate(header_row):
        if col is not None:
            col_str = str(col).strip()
            headers.append((idx, col_str, col_str.lower()))

    recipients = []
    for row in rows[1:]:
        # Проверяем, что строка не пустая
        if not any(cell is not None and str(cell).strip() != "" for cell in row):
            continue

        person_data = {}
        for idx, orig_header, lower_header in headers:
            if idx < len(row):
                val = row[idx]
                if val is not None:
                    # Чистим строковые значения от лишних пробелов
                    if isinstance(val, str):
                        val = val.strip()
                    elif isinstance(val, float) and val.is_integer():
                        val = int(val)
                    person_data[orig_header] = val
                    person_data[lower_header] = val
                else:
                    person_data[orig_header] = ""
                    person_data[lower_header] = ""

        # Проверяем наличие email
        email_val = person_data.get("email") or person_data.get("Email") or person_data.get("EMAIL")
        if email_val:
            recipients.append(person_data)

    return recipients
