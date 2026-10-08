import re
from jinja2 import Template


def render_email_content(template_str: str, data_dict: dict, is_pure_html: bool = False, clean_cid: str = None) -> str:
    text = (template_str or "").strip()

    salutation_val = str(
        data_dict.get("обращение") or 
        data_dict.get("Обращение") or 
        data_dict.get("salutation") or 
        ""
    )
    first_name_val = str(data_dict.get("first_name") or data_dict.get("имя") or data_dict.get("Имя") or "")
    patronymic_val = str(data_dict.get("patronymic") or data_dict.get("отчество") or data_dict.get("Отчество") or "")
    last_name_val = str(data_dict.get("last_name") or data_dict.get("фамилия") or data_dict.get("Фамилия") or "")

    # Если поле текста абсолютно пустое — даем стандартное приветствие
    if not text and not is_pure_html:
        sal = salutation_val if salutation_val else "Уважаемый(-ая)"
        text = f"{sal} {first_name_val} {patronymic_val}!\n\nНаправляем вам документы для ознакомления."

    # 1. Прямая замена русских и системных тегов
    pdf_text = str(data_dict.get("содержимое_pdf") or data_dict.get("pdf_content") or "")

    replacements = {
        "{обращение}": salutation_val,
        "{Обращение}": salutation_val,
        "[обращение]": salutation_val,
        "[Обращение]": salutation_val,
        "{имя}": first_name_val,
        "{Имя}": first_name_val,
        "[имя]": first_name_val,
        "[Имя]": first_name_val,
        "{отчество}": patronymic_val,
        "{Отчество}": patronymic_val,
        "[отчество]": patronymic_val,
        "[Отчество]": patronymic_val,
        "{фамилия}": last_name_val,
        "{Фамилия}": last_name_val,
        "[фамилия]": last_name_val,
        "[Фамилия}": last_name_val,
        "{содержимое_pdf}": pdf_text,
        "{Содержимое_pdf}": pdf_text,
        "{содержимое pdf}": pdf_text,
        "{Содержимое pdf}": pdf_text,
        "{pdf_content}": pdf_text,
        "{PDF_CONTENT}": pdf_text,
    }
    for tag, val in replacements.items():
        text = text.replace(tag, val)

    # 2. Рендеринг через Jinja2 (для выражений вида {{ first_name }}, {{ salutation }}, {{ обращение }})
    try:
        jinja_tpl = Template(text)
        rendered_text = jinja_tpl.render(**data_dict)
    except Exception as e:
        rendered_text = text
        print(f"[ПРЕДУПРЕЖДЕНИЕ] Ошибка Jinja2: {e}")

    # 3. Подстановка любых кастомных колонок из Excel вида {колонка}
    for k, v in data_dict.items():
        val_str = str(v if v is not None else "")
        rendered_text = rendered_text.replace(f"{{{k}}}", val_str)
        rendered_text = rendered_text.replace(f"{{{k.capitalize()}}}", val_str)
        rendered_text = rendered_text.replace(f"{{{k.upper()}}}", val_str)

    if is_pure_html:
        return rendered_text

    # 4. Превращение ссылок в кликабельные HTML <a>
    markdown_pattern = r'\[([^\]]+)\]\((https?://[^\s\)]+)\)'
    rendered_text = re.sub(
        markdown_pattern, 
        r'<a href="\2" style="color: #1a73e8; text-decoration: underline;">\1</a>', 
        rendered_text
    )
    raw_url_pattern = r'(?<!href=")(https?://[^\s<>]+)'
    rendered_text = re.sub(
        raw_url_pattern, 
        r'<a href="\1" style="color: #1a73e8; text-decoration: underline;">\1</a>', 
        rendered_text
    )

    # 5. Переносы строк и баннер
    html_ready = rendered_text.replace("\n", "<br>")
    
    if clean_cid:
        src_attr = clean_cid if clean_cid.startswith("data:") else f"cid:{clean_cid.strip('<>')}"
        img_tag = f'<img src="{src_attr}" alt="Баннер" style="max-width:100%; height:auto; display:block; margin-bottom:15px;"><br>'
    else:
        img_tag = ""

    return f"""
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif, Arial; font-size: 14px; line-height: 1.6; color: #202124;">
        {img_tag}
        {html_ready}
    </div>
    """
