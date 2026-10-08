import os
import json
import uuid
import datetime
from config import DATA_DIR

TEMPLATES_FILE = os.path.join(DATA_DIR, "templates_storage.json")

DEFAULT_TEMPLATES = [
    {
        "id": "tpl_official",
        "name": "Официальное уведомление кафедры",
        "subject": "Официальное уведомление ДГТУ",
        "letter_text": "{обращение} {имя} {отчество}!\n\nНаправляем вам официальный документ для ознакомления и руководства в работе.\n\n{содержимое_pdf}\n\nПо всем возникающим вопросам обращайтесь в деканат.\n\nС уважением,\nДонской государственный технический университет",
        "is_pure_html": False,
        "pdf_mode": "visual",
        "created_at": "2026-10-08 00:00:00"
    },
    {
        "id": "tpl_arrears",
        "name": "Академическая задолженность",
        "subject": "Уведомление о ликвидации задолженности",
        "letter_text": "{обращение} {имя} {отчество}!\n\nИнформируем вас о необходимости ликвидации академической задолженности по учебному плану в установленный срок.\n\n{содержимое_pdf}\n\nГрафик пересдач и ведомости доступны в личном кабинете обучающегося.\n\nС уважением,\nУчебный отдел ДГТУ",
        "is_pure_html": False,
        "pdf_mode": "visual",
        "created_at": "2026-10-08 00:00:00"
    },
    {
        "id": "tpl_info",
        "name": "Информационное сообщение",
        "subject": "Информационное сообщение для студентов и преподавателей",
        "letter_text": "{обращение} {имя} {отчество}!\n\nПросим вас ознакомиться с актуальной информацией университета.\n\n{содержимое_pdf}\n\nС уважением,\nАдминистрация ДГТУ",
        "is_pure_html": False,
        "pdf_mode": "visual",
        "created_at": "2026-10-08 00:00:00"
    }
]


def list_templates() -> list[dict]:
    if not os.path.exists(TEMPLATES_FILE):
        save_all_templates(DEFAULT_TEMPLATES)
        return DEFAULT_TEMPLATES

    try:
        with open(TEMPLATES_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, list) and len(data) > 0:
                return data
    except Exception as e:
        print(f"[TEMPLATES] Ошибка чтения шаблонов: {e}")

    save_all_templates(DEFAULT_TEMPLATES)
    return DEFAULT_TEMPLATES


def save_all_templates(templates: list[dict]):
    try:
        with open(TEMPLATES_FILE, "w", encoding="utf-8") as f:
            json.dump(templates, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[TEMPLATES] Ошибка записи шаблонов: {e}")


def save_template(name: str, subject: str, letter_text: str, is_pure_html: bool = False, pdf_mode: str = "visual") -> dict:
    templates = list_templates()
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    existing = next((t for t in templates if t["name"].strip().lower() == name.strip().lower()), None)
    if existing:
        existing["subject"] = subject
        existing["letter_text"] = letter_text
        existing["is_pure_html"] = is_pure_html
        existing["pdf_mode"] = pdf_mode
        existing["updated_at"] = now_str
        save_all_templates(templates)
        return existing

    new_tpl = {
        "id": f"tpl_{uuid.uuid4().hex[:8]}",
        "name": name.strip(),
        "subject": subject,
        "letter_text": letter_text,
        "is_pure_html": is_pure_html,
        "pdf_mode": pdf_mode,
        "created_at": now_str
    }
    templates.append(new_tpl)
    save_all_templates(templates)
    return new_tpl


def delete_template(template_id: str) -> bool:
    templates = list_templates()
    filtered = [t for t in templates if t["id"] != template_id]
    if len(filtered) != len(templates):
        save_all_templates(filtered)
        return True
    return False


def get_template(template_id: str) -> dict | None:
    templates = list_templates()
    return next((t for t in templates if t["id"] == template_id), None)
