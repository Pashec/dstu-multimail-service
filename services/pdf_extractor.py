import io
import re
import html
from typing import Optional, List, Tuple

# Попытка импорта pypdfium2 для высококачественного визуального рендеринга страниц
try:
    import pypdfium2 as pdfium
except ImportError:
    pdfium = None

# Попытка импорта pypdf для извлечения текста
try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None


def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    """
    Извлекает чистый текст со всех страниц PDF документа.
    """
    if not pdf_bytes or not PdfReader:
        return ""
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        pages_text = []
        for idx, page in enumerate(reader.pages):
            try:
                page_text = page.extract_text()
                if page_text and page_text.strip():
                    pages_text.append(page_text.strip())
            except Exception as page_err:
                print(f"[PDF Extractor] Предупреждение: страница {idx + 1}: {page_err}")
                continue

        full_text = "\n\n".join(pages_text).strip()
        lines = [line.strip() for line in full_text.splitlines()]
        cleaned_text = "\n".join(lines)
        cleaned_text = re.sub(r'\n{3,}', '\n\n', cleaned_text)
        return cleaned_text
    except Exception as e:
        err_msg = f"[Ошибка распаковки текста PDF: {str(e)}]"
        print(f"[PDF Extractor] {err_msg}")
        return err_msg


def format_pdf_text_for_html(plain_text: str) -> str:
    """
    Форматирует извлеченный текст для вставки в HTML-письмо.
    """
    if not plain_text or not plain_text.strip():
        return ""

    escaped = html.escape(plain_text.strip())
    paragraphs = escaped.split("\n\n")

    html_parts = []
    for p in paragraphs:
        p_clean = p.strip()
        if not p_clean:
            continue
        p_html = p_clean.replace("\n", "<br>")
        html_parts.append(
            f'<div style="margin: 0 0 10px 0; line-height: 1.55; color: #1e293b;">{p_html}</div>'
        )

    return "\n".join(html_parts)


def render_pdf_to_images(pdf_bytes: bytes, scale: float = 2.0) -> List[Tuple[bytes, str]]:
    """
    Рендерит страницы PDF в высококачественные растровые изображения (JPEG).
    Позволяет на 100% сохранить чертежи, графики, таблицы, печати и шрифты
    при отображении документа прямо внутри тела email-письма.
    Возвращает список кортежей: [(image_bytes, filename), ...]
    """
    if not pdf_bytes:
        return []

    if pdfium is None:
        print("[PDF Extractor] Предупреждение: pypdfium2 не установлен. Визуальный рендеринг недоступен.")
        return []

    images = []
    try:
        pdf = pdfium.PdfDocument(pdf_bytes)
        for i, page in enumerate(pdf):
            # Рендерим с масштабом 2.0 (эквивалент ~150-200 DPI для четкости чертежей и печатей)
            pil_image = page.render(scale=scale).to_pil()
            buf = io.BytesIO()
            # Сохраняем в RGB JPEG с оптимизацией веса
            pil_image.convert("RGB").save(buf, format="JPEG", quality=90, optimize=True)
            img_bytes = buf.getvalue()
            images.append((img_bytes, f"doc_page_{i + 1}.jpg"))
        return images
    except Exception as e:
        print(f"[PDF Extractor] Ошибка рендеринга PDF в изображения: {e}")
        return []
