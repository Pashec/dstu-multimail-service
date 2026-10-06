from io import BytesIO
import openpyxl

def read_excel_from_bytes(file_bytes: bytes) -> list[dict]:
    """Читает .xlsx напрямую из оперативной памяти."""
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    sheet = wb.active
    rows = list(sheet.iter_rows(values_only=True))

    if not rows:
        wb.close()
        return []

    headers = [str(cell).strip() for cell in rows[0] if cell is not None]
    recipients = []

    for row in rows[1:]:
        if not any(row):
            continue
        recipients.append(dict(zip(headers, row)))

    wb.close()
    return recipients