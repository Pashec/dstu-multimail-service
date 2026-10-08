import base64
import os
import asyncio
import random
import ssl
import datetime
import email.utils
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders

try:
    from aiosmtplib import SMTP as AsyncSMTP
    from aiosmtplib.errors import SMTPResponseException, SMTPServerDisconnected, SMTPConnectError
except ImportError:
    AsyncSMTP = None
    SMTPResponseException = Exception
    SMTPServerDisconnected = Exception
    SMTPConnectError = Exception

from config import broadcast_status, append_log
from services.matching import detect_salutation, find_all_personal_docs
from services.template import render_email_content
from services.pdf_extractor import extract_text_from_pdf, format_pdf_text_for_html, render_pdf_to_images
from services.report import generate_excel_report


# Глобальное событие для управления паузой рассылки
pause_event = asyncio.Event()
pause_event.set()


def log_user_event(message: str, is_error: bool = False):
    """
    Добавляет лаконичное сообщение в терминал интерфейса.
    Строгий запрет на эмодзи в журнале событий.
    """
    now_str = datetime.datetime.now().strftime("%H:%M:%S")
    formatted = f"[{now_str}] {message}"
    print(formatted, flush=True)
    append_log(formatted, is_error=is_error)


def build_email_payload(
    email_address: str,
    email_to: str,
    subject: str,
    letter_text: str,
    is_pure_html: bool,
    person: dict,
    inline_img_bytes: bytes | None,
    inline_img_name: str | None,
    attachments_data: list[tuple[bytes, str]],
    personal_files_map: dict[str, tuple[bytes, str]] | None,
    pdf_mode: str = "visual",
    also_attach_pdf: bool = True,
    smtp_server: str = "mail.donstu.ru",
    for_preview: bool = False
) -> dict:
    """
    Собирает контекст, HTML-содержимое и структуру письма MIME для конкретного адресата.
    Используется как для реальной отправки, так и для интерактивного предпросмотра (Live Preview).
    """
    first_name = person.get("имя") or person.get("Имя") or person.get("name") or ""
    patronymic = person.get("отчество") or person.get("Отчество") or person.get("patronymic") or ""
    last_name = person.get("фамилия") or person.get("Фамилия") or person.get("last_name") or ""

    custom_salutation = (
        person.get("обращение") or 
        person.get("Обращение") or 
        person.get("salutation") or 
        person.get("Salutation")
    )
    if custom_salutation and str(custom_salutation).strip():
        salutation = str(custom_salutation).strip()
        gender = "custom"
    else:
        gender_data = detect_salutation(str(first_name), str(patronymic))
        salutation = gender_data["salutation"]
        gender = gender_data["gender"]

    personal_docs = find_all_personal_docs(person, personal_files_map) if personal_files_map else []

    extracted_pdf_html = ""
    rendered_pdf_images = []
    files_to_attach = []

    for p_bytes, p_name in personal_docs:
        p_name_str = str(p_name)
        is_pdf = p_name_str.lower().endswith(".pdf")

        if is_pdf:
            if pdf_mode in ("visual", "visual_and_attachment", "visual_only"):
                pages = render_pdf_to_images(p_bytes, scale=2.0)
                if pages:
                    img_html_tags = []
                    for idx, (p_img_bytes, p_img_name) in enumerate(pages, start=1):
                        if for_preview:
                            b64_img = base64.b64encode(p_img_bytes).decode("ascii")
                            src_url = f"data:image/jpeg;base64,{b64_img}"
                            img_html_tags.append(
                                f'<div style="margin: 18px 0; text-align: center;">'
                                f'<div style="font-size: 11px; color: #64748b; font-family: sans-serif; margin-bottom: 5px; font-weight: 600;">Страница {idx} из {len(pages)}</div>'
                                f'<img src="{src_url}" alt="Документ (Страница {idx})" '
                                f'style="max-width: 100%; height: auto; display: inline-block; margin: 0 auto; box-shadow: 0 4px 16px rgba(0,0,0,0.12); border-radius: 4px; border: 1px solid #e2e8f0;">'
                                f'</div>'
                            )
                        else:
                            page_cid = email.utils.make_msgid(domain="local")
                            clean_page_cid = page_cid.strip("<>")
                            rendered_pdf_images.append((p_img_bytes, page_cid, p_img_name))
                            img_html_tags.append(
                                f'<div style="margin: 15px 0 20px 0; text-align: left;">'
                                f'<img src="cid:{clean_page_cid}" alt="Документ (Страница {idx})" '
                                f'style="max-width: 100%; height: auto; display: block; margin: 0; padding: 0; border: none;">'
                                f'</div>'
                            )
                    extracted_pdf_html = "\n".join(img_html_tags)
                else:
                    raw_text = extract_text_from_pdf(p_bytes)
                    extracted_pdf_html = format_pdf_text_for_html(raw_text)

            elif pdf_mode in ("text", "body_and_attachment", "body_only"):
                raw_text = extract_text_from_pdf(p_bytes)
                if raw_text and not raw_text.startswith("[Ошибка"):
                    extracted_pdf_html = format_pdf_text_for_html(raw_text)

            if pdf_mode in ("visual_only", "body_only"):
                pass
            elif not also_attach_pdf and pdf_mode not in ("none", "attachment_only"):
                pass
            else:
                files_to_attach.append((p_bytes, p_name_str))
        else:
            files_to_attach.append((p_bytes, p_name_str))

    merged_context = {
        **person,
        "salutation": salutation,
        "gender": gender,
        "first_name": first_name,
        "patronymic": patronymic,
        "last_name": last_name,
        "фамилия": last_name,
        "имя": first_name,
        "отчество": patronymic,
        "обращение": salutation,
        "Обращение": salutation,
        "содержимое_pdf": extracted_pdf_html,
        "pdf_content": extracted_pdf_html
    }

    if for_preview and inline_img_bytes:
        subtype = inline_img_name.split(".")[-1] if inline_img_name and "." in inline_img_name else "png"
        b64_banner = base64.b64encode(inline_img_bytes).decode("ascii")
        clean_cid = f"data:image/{subtype};base64,{b64_banner}"
        image_cid = None
    else:
        image_cid = email.utils.make_msgid(domain="local") if inline_img_bytes else None
        clean_cid = image_cid.strip("<>") if image_cid else None

    has_pdf_tag = any(
        tag in (letter_text or "") 
        for tag in (
            "{содержимое_pdf}", "{содержимое pdf}", "{pdf_content}",
            "{{ содержимое_pdf }}", "{{ содержимое pdf }}", "{{ pdf_content }}"
        )
    )

    current_letter_text = letter_text or ""
    if extracted_pdf_html and not has_pdf_tag:
        current_letter_text = current_letter_text + "\n\n" + extracted_pdf_html

    html_body = render_email_content(
        template_str=current_letter_text,
        data_dict=merged_context,
        is_pure_html=is_pure_html,
        clean_cid=clean_cid
    )

    # Формирование MIME объекта
    msg_root = MIMEMultipart("mixed")
    msg_root["Subject"] = subject
    msg_root["From"] = email_address
    msg_root["To"] = email_to
    msg_root["Date"] = email.utils.formatdate(localtime=True)
    msg_root["Message-ID"] = email.utils.make_msgid(domain=smtp_server)

    msg_html_group = MIMEMultipart("related")
    msg_html_group.attach(MIMEText(html_body, "html", "utf-8"))

    if inline_img_bytes and image_cid:
        subtype = inline_img_name.split(".")[-1] if inline_img_name and "." in inline_img_name else "png"
        img_part = MIMEImage(inline_img_bytes, _subtype=subtype)
        img_part.add_header("Content-ID", image_cid)
        img_part.add_header("Content-Disposition", "inline", filename=inline_img_name)
        msg_html_group.attach(img_part)

    if rendered_pdf_images:
        for img_data, cid_val, fname in rendered_pdf_images:
            page_img_part = MIMEImage(img_data, _subtype="jpeg")
            page_img_part.add_header("Content-ID", cid_val)
            page_img_part.add_header("Content-Disposition", "inline", filename=fname)
            msg_html_group.attach(page_img_part)

    msg_root.attach(msg_html_group)

    def attach_file(file_data: bytes, file_title: str):
        part = MIMEBase("application", "octet-stream")
        part.set_payload(file_data)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", "attachment", filename=("utf-8", "", file_title))
        part.set_param("name", file_title)
        msg_root.attach(part)

    if attachments_data:
        for item in attachments_data:
            f_bytes = item[0] if isinstance(item[0], (bytes, bytearray)) else item[1]
            f_name = item[1] if isinstance(item[0], (bytes, bytearray)) else item[0]
            if f_bytes and len(f_bytes) > 0:
                attach_file(f_bytes, str(f_name))

    for f_bytes, f_name in files_to_attach:
        if f_bytes and len(f_bytes) > 0:
            attach_file(f_bytes, f_name)

    all_attached_names = [f[1] for f in files_to_attach] + [str(item[1] if isinstance(item[0], (bytes, bytearray)) else item[0]) for item in attachments_data]

    return {
        "msg_root": msg_root,
        "html_body": html_body,
        "fio": f"{last_name} {first_name} {patronymic}".strip(),
        "salutation": salutation,
        "personal_docs": personal_docs,
        "attached_filenames": all_attached_names,
        "rendered_images_count": len(rendered_pdf_images)
    }


class SMTPClientPool:
    """
    Менеджер SMTP-соединения с поддержкой Keep-Alive и автоматического восстановления связи.
    """
    def __init__(self, hostname: str, port: int, user: str, password: str):
        self.hostname = hostname
        self.port = port
        self.user = user
        self.password = password
        self.client: AsyncSMTP | None = None

    async def get_client(self) -> AsyncSMTP:
        if self.client is not None:
            try:
                # Проверка живости соединения (NOOP)
                await self.client.noop()
                return self.client
            except Exception:
                try:
                    await self.client.quit()
                except Exception:
                    pass
                self.client = None

        tls_context = ssl.create_default_context()
        tls_context.check_hostname = False
        tls_context.verify_mode = ssl.CERT_NONE

        client = AsyncSMTP(
            hostname=self.hostname,
            port=self.port,
            use_tls=(self.port == 465),
            tls_context=tls_context,
            timeout=30.0
        )
        await client.connect()
        if self.port in (25, 587):
            try:
                await client.starttls()
            except Exception:
                pass

        await client.login(self.user, self.password)
        self.client = client
        return self.client

    async def close(self):
        if self.client is not None:
            try:
                await self.client.quit()
            except Exception:
                pass
            self.client = None


async def run_bg_broadcast(
    smtp_server: str,
    smtp_port: int,
    email_address: str,
    email_password: str,
    subject: str,
    letter_text: str,
    is_pure_html: bool,
    recipients: list,
    inline_img_bytes: bytes | None,
    inline_img_name: str | None,
    attachments_data: list[tuple[bytes, str]],
    delay_min: float,
    delay_max: float,
    personal_files_map: dict[str, tuple[bytes, str]] = None,
    pdf_mode: str = "visual",
    also_attach_pdf: bool = True,
    auto_retry: bool = True
):
    broadcast_status.update({
        "is_running": True,
        "is_paused": False,
        "should_stop": False,
        "total": len(recipients),
        "sent": 0,
        "skipped": 0,
        "errors": [],
        "logs": [],
        "last_report_file": None,
        "last_report_name": None,
        "current_recipient": None,
        "progress_percent": 0
    })
    pause_event.set()

    if AsyncSMTP is None:
        log_user_event("[ОШИБКА] Библиотека aiosmtplib не установлена! Выполните: pip install -r requirements.txt", is_error=True)
        broadcast_status["is_running"] = False
        return

    mode_labels = {
        "visual": "Визуально 1 в 1",
        "text": "Текстовая распаковка",
        "none": "Только файл-вложение",
        "attachment_only": "Только файл-вложение"
    }
    mode_name = mode_labels.get(pdf_mode, pdf_mode)

    log_user_event(f"[СТАРТ] Запуск рассылки «ДГТУ Рассылка»: получателей — {len(recipients)}.")
    if personal_files_map:
        log_user_event(f"[ФАЙЛЫ] Персональных файлов в памяти: {len(personal_files_map)} шт. (Режим PDF: {mode_name})")

    pool = SMTPClientPool(smtp_server, smtp_port, email_address, email_password)
    report_records = []

    try:
        for index, person in enumerate(recipients, start=1):
            # Проверка остановки
            if broadcast_status["should_stop"]:
                log_user_event(f"[СТОП] Рассылка остановлена пользователем на шаге {index} из {len(recipients)}")
                break

            # Проверка паузы
            while broadcast_status.get("is_paused"):
                if broadcast_status.get("should_stop"):
                    break
                await asyncio.sleep(0.5)

            if broadcast_status["should_stop"]:
                break

            email_to = person.get("email") or person.get("Email") or person.get("EMAIL")
            if not email_to:
                log_user_event(f"[ПРОПУСК] Строка #{index}: отсутствует email-адрес")
                broadcast_status["skipped"] += 1
                report_records.append({
                    "index": index,
                    "email": "—",
                    "fio": f"{person.get('фамилия', '')} {person.get('имя', '')}".strip(),
                    "salutation": "—",
                    "status": "Пропущено",
                    "personal_files": "—",
                    "error_message": "Отсутствует email-адрес",
                    "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
                continue

            email_to = str(email_to).strip()
            broadcast_status["current_recipient"] = email_to

            payload = build_email_payload(
                email_address=email_address,
                email_to=email_to,
                subject=subject,
                letter_text=letter_text,
                is_pure_html=is_pure_html,
                person=person,
                inline_img_bytes=inline_img_bytes,
                inline_img_name=inline_img_name,
                attachments_data=attachments_data,
                personal_files_map=personal_files_map,
                pdf_mode=pdf_mode,
                also_attach_pdf=also_attach_pdf,
                smtp_server=smtp_server
            )

            # Отправка с механизмом автоматического Retry
            max_attempts = 3 if auto_retry else 1
            success = False
            last_err_msg = ""

            for attempt in range(1, max_attempts + 1):
                try:
                    smtp_client = await pool.get_client()
                    await smtp_client.send_message(payload["msg_root"])
                    success = True
                    break
                except Exception as send_err:
                    last_err_msg = str(send_err)
                    # Сброс пула при разрыве связи
                    await pool.close()

                    is_transient = any(
                        err_type in type(send_err).__name__ 
                        for err_type in ("SMTPServerDisconnected", "SMTPConnectError", "TimeoutError", "ConnectionResetError")
                    ) or "421" in last_err_msg or "450" in last_err_msg or "451" in last_err_msg

                    if attempt < max_attempts and is_transient:
                        backoff = 2.0 * attempt
                        log_user_event(f"[ПОВТОР] Строка #{index} ({email_to}): временный сбой соединения, повтор через {backoff:.1f}с (попытка {attempt}/{max_attempts})...")
                        await asyncio.sleep(backoff)
                    else:
                        break

            if success:
                broadcast_status["sent"] += 1
                log_user_event(f"[УСПЕХ] [{index}/{len(recipients)}] Письмо отправлено на {email_to}")
                report_records.append({
                    "index": index,
                    "email": email_to,
                    "fio": payload["fio"],
                    "salutation": payload["salutation"],
                    "status": "Успешно",
                    "personal_files": ", ".join(payload["attached_filenames"]) if payload["attached_filenames"] else "—",
                    "error_message": "",
                    "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
            else:
                err_text = f"[ОШИБКА] Строка #{index} ({email_to}): {last_err_msg}"
                log_user_event(err_text, is_error=True)
                report_records.append({
                    "index": index,
                    "email": email_to,
                    "fio": payload["fio"],
                    "salutation": payload["salutation"],
                    "status": "Ошибка",
                    "personal_files": ", ".join(payload["attached_filenames"]) if payload["attached_filenames"] else "—",
                    "error_message": last_err_msg,
                    "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })

            # Обновление процента
            if len(recipients) > 0:
                broadcast_status["progress_percent"] = int(((broadcast_status["sent"] + broadcast_status["skipped"] + len(broadcast_status["errors"])) / len(recipients)) * 100)

            if index < len(recipients) and not broadcast_status["should_stop"]:
                await asyncio.sleep(random.uniform(delay_min, delay_max))

    finally:
        await pool.close()

    # Формирование итогового отчета Excel
    summary_stats = {
        "total": len(recipients),
        "sent": broadcast_status["sent"],
        "errors": len(broadcast_status["errors"]),
        "skipped": broadcast_status["skipped"],
        "subject": subject
    }
    try:
        report_path = generate_excel_report(report_records, summary_stats)
        broadcast_status["last_report_file"] = report_path
        broadcast_status["last_report_name"] = os.path.basename(report_path)
        log_user_event(f"[ОТЧЕТ] Итоговый Excel-отчет сформирован: {os.path.basename(report_path)}")
    except Exception as rep_err:
        log_user_event(f"[ПРЕДУПРЕЖДЕНИЕ] Ошибка создания отчета: {rep_err}", is_error=True)

    broadcast_status["is_running"] = False
    broadcast_status["is_paused"] = False
    log_user_event(f"[ФИНИШ] Рассылка завершена. Успешно: {broadcast_status['sent']} из {broadcast_status['total']}")
