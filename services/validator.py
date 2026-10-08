import re
from services.matching import detect_salutation, find_all_personal_docs

EMAIL_REGEX = re.compile(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$")

COMMON_DOMAIN_CORRECTIONS = {
    "gmai.com": "gmail.com",
    "gamil.com": "gmail.com",
    "gmial.com": "gmail.com",
    "yandx.ru": "yandex.ru",
    "yadnex.ru": "yandex.ru",
    "yande.ru": "yandex.ru",
    "donstu.com": "donstu.ru",
    "donstu.net": "donstu.ru",
    "mail.com": "mail.ru (если имелась в виду почта РФ)",
    "inbox.com": "inbox.ru",
}


def validate_email_address(email_str: str) -> tuple[bool, str]:
    if not email_str or not isinstance(email_str, str):
        return False, "Email не указан"

    cleaned = email_str.strip()
    if not EMAIL_REGEX.match(cleaned):
        return False, "Некорректный синтаксис адреса"

    parts = cleaned.split("@")
    if len(parts) == 2:
        domain = parts[1].lower()
        if domain in COMMON_DOMAIN_CORRECTIONS:
            return True, f"Возможная опечатка в домене: @{domain} ➔ рекомендуемый: @{COMMON_DOMAIN_CORRECTIONS[domain]}"

    return True, ""


def run_preflight_check(
    recipients: list[dict],
    personal_files_map: dict[str, tuple[bytes, str]] = None,
    letter_text: str = "",
    subject: str = ""
) -> dict:
    total = len(recipients)
    valid_emails = 0
    invalid_recipients = []
    typo_warnings = []
    missing_files_recipients = []
    matched_files_count = 0

    recognized_columns = set()
    for person in recipients:
        recognized_columns.update(k.lower() for k in person.keys())

    used_tags = re.findall(r"\{([^{}]+)\}", letter_text or "")
    missing_tag_warnings = []
    standard_tags = {
        "имя", "отчество", "фамилия", "обращение", 
        "содержимое_pdf", "содержимое pdf", "pdf_content",
        "email"
    }

    for tag in used_tags:
        tag_clean = tag.strip().lower()
        if tag_clean not in standard_tags and tag_clean not in recognized_columns:
            missing_tag_warnings.append(f"Тег «{{{tag}}}» отсутствует среди колонок таблицы адресатов.")

    for idx, person in enumerate(recipients, start=1):
        email_val = person.get("email") or person.get("Email") or person.get("EMAIL")
        fio = f"{person.get('фамилия', '')} {person.get('имя', '')} {person.get('отчество', '')}".strip()
        fio_display = fio if fio else f"Строка #{idx}"

        if not email_val:
            invalid_recipients.append({
                "row": idx,
                "name": fio_display,
                "email": "—",
                "reason": "Отсутствует адрес электронной почты"
            })
            continue

        is_valid, msg = validate_email_address(str(email_val))
        if not is_valid:
            invalid_recipients.append({
                "row": idx,
                "name": fio_display,
                "email": str(email_val),
                "reason": msg
            })
        else:
            valid_emails += 1
            if msg:
                typo_warnings.append({
                    "row": idx,
                    "name": fio_display,
                    "email": str(email_val),
                    "warning": msg
                })

        if personal_files_map is not None:
            personal_docs = find_all_personal_docs(person, personal_files_map)
            if personal_docs:
                matched_files_count += len(personal_docs)
            else:
                missing_files_recipients.append({
                    "row": idx,
                    "name": fio_display,
                    "email": str(email_val) if email_val else "—"
                })

    is_ready = (total > 0) and (valid_emails > 0) and (len(invalid_recipients) == 0)

    return {
        "total_recipients": total,
        "valid_emails_count": valid_emails,
        "invalid_count": len(invalid_recipients),
        "invalid_recipients": invalid_recipients,
        "typo_warnings": typo_warnings,
        "total_uploaded_files": len(personal_files_map) if personal_files_map else 0,
        "matched_files_count": matched_files_count,
        "missing_files_count": len(missing_files_recipients),
        "missing_files_recipients": missing_files_recipients,
        "missing_tag_warnings": missing_tag_warnings,
        "has_subject": bool(subject and subject.strip()),
        "ready_to_send": is_ready
    }
