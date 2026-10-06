import asyncio
import random
import ssl
import email.utils
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
from aiosmtplib import SMTP as AsyncSMTP

from config import broadcast_status
from services.matching import detect_salutation, find_personal_doc
from services.template import render_email_content

def log_event(message: str, is_error: bool = False):
    """Печатает лог в консоль без задержек и отправляет в веб-интерфейс"""
    print(message, flush=True)
    if is_error:
        broadcast_status["errors"].append(message)
    broadcast_status["logs"].append(message)

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
):
    broadcast_status.update({
        "is_running": True,
        "should_stop": False,
        "total": len(recipients),
        "sent": 0,
        "skipped": 0,
        "errors": [],
        "logs": []
    })

    log_event(f"[СТАРТ] Запуск рассылки. Всего получателей: {len(recipients)}")
    if personal_files_map:
        log_event(f"[ФАЙЛЫ] Загружено персональных файлов: {list(personal_files_map.keys())}")
    else:
        log_event("[ВНИМАНИЕ] Список персональных файлов пуст!")

    for index, person in enumerate(recipients, start=1):
        if broadcast_status["should_stop"]:
            log_event(f"[СТОП] Процесс остановлен пользователем на шаге {index}!")
            break

        email_to = person.get("email") or person.get("Email") or person.get("EMAIL")
        if not email_to:
            log_event(f"[ПРОПУСК] Строка {index}: нет email адреса")
            broadcast_status["skipped"] += 1
            continue

        email_to = str(email_to).strip()
        first_name = person.get("имя") or person.get("Имя") or person.get("name") or ""
        patronymic = person.get("отчество") or person.get("Отчество") or person.get("patronymic") or ""
        last_name = person.get("фамилия") or person.get("Фамилия") or person.get("last_name") or ""

        gender_data = detect_salutation(str(first_name), str(patronymic))
        merged_context = {
            **person,
            "salutation": gender_data["salutation"],
            "gender": gender_data["gender"],
            "first_name": first_name,
            "patronymic": patronymic,
            "last_name": last_name,          
            "фамилия": last_name,             
            "имя": first_name,
            "отчество": patronymic,
            "обращение": gender_data["salutation"]
        }

        tls_context = ssl.create_default_context()
        tls_context.check_hostname = False
        tls_context.verify_mode = ssl.CERT_NONE

        smtp_client = AsyncSMTP(
            hostname=smtp_server,
            port=smtp_port,
            use_tls=(smtp_port == 465),
            tls_context=tls_context
        )

        try:
            await smtp_client.connect()
            if smtp_port in (25, 587):
                try:
                    await smtp_client.starttls()
                except Exception:
                    pass

            await smtp_client.login(email_address, email_password)

            image_cid = email.utils.make_msgid(domain="local") if inline_img_bytes else None
            clean_cid = image_cid.strip("<>") if image_cid else None

            html_body = render_email_content(
                template_str=letter_text,
                data_dict=merged_context,
                is_pure_html=is_pure_html,
                clean_cid=clean_cid
            )

            msg_root = MIMEMultipart("mixed")
            msg_root["Subject"] = subject
            msg_root["From"] = email_address
            msg_root["To"] = email_to
            msg_root["Date"] = email.utils.formatdate(localtime=True)
            msg_root["Message-ID"] = email.utils.make_msgid(domain=smtp_server)

            # 2. Контейнер HTML и встроенного баннера
            msg_html_group = MIMEMultipart("related")
            msg_html_group.attach(MIMEText(html_body, "html", "utf-8"))

            # Встроенный баннер (inline image с Content-ID)
            if inline_img_bytes and image_cid:
                subtype = inline_img_name.split(".")[-1] if inline_img_name and "." in inline_img_name else "png"
                img_part = MIMEImage(inline_img_bytes, _subtype=subtype)
                img_part.add_header("Content-ID", image_cid)
                img_part.add_header("Content-Disposition", "inline", filename=inline_img_name)
                msg_html_group.attach(img_part)

            # ПРИКРЕПЛЯЕМ HTML-ГРУППУ СТРОГО ОДИН РАЗ:
            msg_root.attach(msg_html_group)

            # 3. Вспомогательная функция для файлов вложений
            def attach_document(file_data: bytes, file_title: str):
                part = MIMEBase("application", "octet-stream")
                part.set_payload(file_data)
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition", 
                    "attachment", 
                    filename=("utf-8", "", file_title)
                )
                part.set_param("name", file_title)
                msg_root.attach(part)

            # 4. Прикрепление общих файлов
            if attachments_data:
                for item in attachments_data:
                    f_bytes = item[0] if isinstance(item[0], (bytes, bytearray)) else item[1]
                    f_name = item[1] if isinstance(item[0], (bytes, bytearray)) else item[0]
                    if f_bytes and len(f_bytes) > 0:
                        attach_document(f_bytes, str(f_name))

            # 5. Прикрепление персонального документа
            if personal_files_map:
                personal_doc = find_personal_doc(person, personal_files_map)
                if personal_doc:
                    doc_item1, doc_item2 = personal_doc
                    p_bytes = doc_item1 if isinstance(doc_item1, (bytes, bytearray)) else doc_item2
                    p_name = doc_item2 if isinstance(doc_item1, (bytes, bytearray)) else doc_item1

                    if p_bytes and len(p_bytes) > 0:
                        attach_document(p_bytes, str(p_name))
                        log_event(f"   [+] Документ '{p_name}' успешно прикреплен для {email_to}")
                    else:
                        log_event(f"   [!] Персональный файл '{p_name}' пуст (0 байт)!", is_error=True)
                else:
                    log_event(f"   [-] Персональный документ для адресата {email_to} НЕ найден среди файлов!")

            # 6. Отправка собранного письма
            await smtp_client.send_message(msg_root)
            broadcast_status["sent"] += 1
            log_event(f"[ГОТОВО] [{index}/{len(recipients)}] Отправлено на {email_to}")

        except Exception as err:
            err_msg = f"Ошибка в строке №{index} ({email_to}): {err}"
            log_event(err_msg, is_error=True)
        finally:
            try:
                await smtp_client.quit()
            except Exception:
                pass

        if index < len(recipients) and not broadcast_status["should_stop"]:
            await asyncio.sleep(random.uniform(delay_min, delay_max))

    broadcast_status["is_running"] = False
    log_event(f"[ФИНИШ] Рассылка завершена. Успешно: {broadcast_status['sent']} из {broadcast_status['total']}")