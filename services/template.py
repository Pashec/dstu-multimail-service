import re
from jinja2 import Template

def render_email_content(template_str: str, data_dict: dict, is_pure_html: bool = False, clean_cid: str = None) -> str:
    text = (template_str or "").strip()

    # Только если поле текста вообще оставили абсолютно пустым — даем стандартную строку
    if not text and not is_pure_html:
        salutation = data_dict.get("salutation", "Уважаемый(-ая)")
        first_name = data_dict.get("first_name", "")
        patronymic = data_dict.get("patronymic", "")
        text = f"{salutation} {first_name} {patronymic}!\n\nНаправляем вам документы для ознакомления."

    # 1. Прямая замена русских тегов (если они еще остались в фигурных скобках)
    replacements = {
        "{обращение}": str(data_dict.get("salutation") or ""),
        "{Обращение}": str(data_dict.get("salutation") or ""),
        "{имя}": str(data_dict.get("first_name") or ""),
        "{Имя}": str(data_dict.get("first_name") or ""),
        "{отчество}": str(data_dict.get("patronymic") or ""),
        "{Отчество}": str(data_dict.get("patronymic") or ""),
        "{фамилия}": str(data_dict.get("last_name") or data_dict.get("фамилия") or ""),
        "{Фамилия}": str(data_dict.get("last_name") or data_dict.get("фамилия") or ""),
    }
    for tag, val in replacements.items():
        text = text.replace(tag, val)

    # 2. Рендеринг через Jinja2 (для конструкций вида {{ first_name }}, {{ last_name }})
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

    if is_pure_html:
        return rendered_text

    # 4. Превращение ссылок в кликабельные HTML <a>
    markdown_pattern = r'\[([^\]]+)\]\((https?://[^\s\)]+)\)'
    rendered_text = re.sub(markdown_pattern, r'<a href="\2" style="color: #1a73e8; text-decoration: underline;">\1</a>', rendered_text)
    raw_url_pattern = r'(?<!href=")(https?://[^\s<>]+)'
    rendered_text = re.sub(raw_url_pattern, r'<a href="\1" style="color: #1a73e8; text-decoration: underline;">\1</a>', rendered_text)

    # 5. Переносы строк и баннер
    html_ready = rendered_text.replace("\n", "<br>")
    img_tag = f'<img src="cid:{clean_cid}" alt="Баннер" style="max-width:100%; height:auto; display:block; margin-bottom:15px;"><br>' if clean_cid else ""

    return f"""
    <div style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.5; color: #202124;">
        {img_tag}
        {html_ready}
    </div>
    """