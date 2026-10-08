import sys
import subprocess
import os
import time
import threading
from fastapi import APIRouter, Form, UploadFile, File, Request, BackgroundTasks
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.templating import Jinja2Templates

import config
from config import (
    broadcast_status, 
    TEMPLATES_DIR, 
    save_window_state,
    load_saved_credentials,
    save_credentials_to_file,
    clear_saved_credentials_file,
    APP_NAME,
    APP_VERSION
)
from services.excel import read_excel_from_bytes
from services.mailer import run_bg_broadcast, build_email_payload, pause_event
from services.validator import run_preflight_check
from services.templates_manager import list_templates, save_template, delete_template, get_template

router = APIRouter()
templates = Jinja2Templates(directory=TEMPLATES_DIR)


def render_template(template_name: str, request: Request, context: dict = None, status_code: int = 200) -> HTMLResponse:
    """Универсальный рендерер шаблонов Jinja2, совместимый со всеми версиями FastAPI/Starlette."""
    ctx = dict(context or {})
    ctx["app_name"] = APP_NAME
    ctx["app_version"] = APP_VERSION
    try:
        return templates.TemplateResponse(request=request, name=template_name, context=ctx, status_code=status_code)
    except TypeError:
        ctx["request"] = request
        return templates.TemplateResponse(template_name, ctx, status_code=status_code)


@router.get("/", response_class=HTMLResponse)
async def index_page(request: Request):
    creds = load_saved_credentials()
    all_templates = list_templates()
    return render_template("index.html", request, {
        "message": None, 
        "creds": creds,
        "templates_list": all_templates
    })


@router.get("/api/status")
async def get_status():
    return broadcast_status


@router.get("/api/credentials")
async def get_credentials():
    return load_saved_credentials()


@router.post("/api/credentials")
async def save_credentials(request: Request):
    try:
        data = await request.json()
        if isinstance(data, dict):
            save_credentials_to_file(data)
            return {"status": "ok"}
    except Exception as e:
        return {"status": "error", "message": str(e)}
    return {"status": "ignored"}


@router.delete("/api/credentials")
async def delete_credentials():
    clear_saved_credentials_file()
    return {"status": "ok"}


@router.post("/api/pause")
async def pause_broadcast():
    """Приостанавливает текущую рассылку."""
    if broadcast_status.get("is_running") and not broadcast_status.get("is_paused"):
        broadcast_status["is_paused"] = True
        pause_event.clear()
        return {"status": "success", "message": "Рассылка приостановлена."}
    return {"status": "error", "message": "Рассылка не активна или уже на паузе."}


@router.post("/api/resume")
async def resume_broadcast():
    """Возобновляет рассылку после паузы."""
    if broadcast_status.get("is_running") and broadcast_status.get("is_paused"):
        broadcast_status["is_paused"] = False
        pause_event.set()
        return {"status": "success", "message": "Рассылка возобновлена."}
    return {"status": "error", "message": "Рассылка не находится на паузе."}


@router.post("/api/stop")
async def stop_broadcast():
    if broadcast_status["is_running"]:
        broadcast_status["should_stop"] = True
        broadcast_status["is_paused"] = False
        pause_event.set()
        return {"status": "success", "message": "Рассылка прервется на текущем шаге."}
    return {"status": "error", "message": "Рассылка не запущена."}


@router.get("/api/templates")
async def api_list_templates():
    return list_templates()


@router.post("/api/templates")
async def api_save_template(request: Request):
    try:
        data = await request.json()
        name = data.get("name")
        subject = data.get("subject", "")
        letter_text = data.get("letter_text", "")
        is_pure_html = data.get("is_pure_html", False)
        pdf_mode = data.get("pdf_mode", "visual")

        if not name or not name.strip():
            return JSONResponse({"status": "error", "message": "Имя шаблона не может быть пустым"}, status_code=400)

        saved = save_template(name=name, subject=subject, letter_text=letter_text, is_pure_html=is_pure_html, pdf_mode=pdf_mode)
        return {"status": "ok", "template": saved}
    except Exception as e:
        return JSONResponse({"status": "error", "message": str(e)}, status_code=500)


@router.delete("/api/templates/{template_id}")
async def api_delete_template(template_id: str):
    res = delete_template(template_id)
    return {"status": "ok" if res else "not_found"}


@router.get("/api/download-report")
async def download_report():
    """Отдает последний сформированный Excel-отчет рассылки."""
    report_file = broadcast_status.get("last_report_file")
    if report_file and os.path.exists(report_file):
        filename = broadcast_status.get("last_report_name") or os.path.basename(report_file)
        return FileResponse(
            path=report_file,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    return JSONResponse({"status": "error", "message": "Файл отчета еще не сформирован"}, status_code=404)



@router.post("/api/open-report")
async def api_open_report():
    """Открывает сформированный Excel-отчет напрямую в приложении по умолчанию (Excel)."""
    report_file = broadcast_status.get("last_report_file")
    if not report_file or not os.path.exists(report_file):
        return JSONResponse({"status": "error", "message": "Файл отчета еще не сформирован"}, status_code=404)

    abs_path = os.path.abspath(report_file)
    try:
        if sys.platform.startswith("win"):
            os.startfile(abs_path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", abs_path])
        else:
            subprocess.Popen(["xdg-open", abs_path])
        return {"status": "success", "message": f"Отчет открыт: {os.path.basename(report_file)}"}
    except Exception as e:
        return JSONResponse({"status": "error", "message": f"Не удалось открыть файл: {e}"}, status_code=500)


@router.post("/api/open-reports-folder")
async def api_open_reports_folder():
    """Открывает папку с отчетами в проводнике Windows / файловом менеджере с выделением файла."""
    report_file = broadcast_status.get("last_report_file")
    reports_dir = os.path.dirname(os.path.abspath(report_file)) if report_file and os.path.exists(report_file) else os.path.join(config.DATA_DIR, "reports")
    os.makedirs(reports_dir, exist_ok=True)

    try:
        if sys.platform.startswith("win"):
            if report_file and os.path.exists(report_file):
                subprocess.Popen(f'explorer /select,"{os.path.abspath(report_file)}"')
            else:
                subprocess.Popen(f'explorer "{os.path.abspath(reports_dir)}"')
        elif sys.platform == "darwin":
            subprocess.Popen(["open", os.path.abspath(reports_dir)])
        else:
            subprocess.Popen(["xdg-open", os.path.abspath(reports_dir)])
        return {"status": "success", "message": "Папка с отчетами открыта в проводнике"}
    except Exception as e:
        return JSONResponse({"status": "error", "message": f"Не удалось открыть папку: {e}"}, status_code=500)

@router.post("/api/save-window")
async def api_save_window(request: Request):
    try:
        data = await request.json()
        if isinstance(data, dict):
            save_window_state(data)
            return {"status": "ok"}
    except Exception:
        pass
    return {"status": "ignored"}


@router.post("/shutdown")
async def shutdown_server():
    threading.Thread(target=lambda: (time.sleep(0.3), os._exit(0)), daemon=True).start()
    return {"status": "success", "message": "Сервер остановлен"}


def normalize_template_text(raw_text: str) -> str:
    if not raw_text:
        return ""
    replacements = {
        "{имя}": "{{ first_name }}",
        "{Имя}": "{{ first_name }}",
        "[имя]": "{{ first_name }}",
        "[Имя]": "{{ first_name }}",
        "{отчество}": "{{ patronymic }}",
        "{Отчество}": "{{ patronymic }}",
        "[отчество]": "{{ patronymic }}",
        "[Отчество]": "{{ patronymic }}",
        "{фамилия}": "{{ last_name }}",
        "{Фамилия}": "{{ last_name }}",
        "[фамилия]": "{{ last_name }}",
        "[Фамилия}": "{{ last_name }}",
        "{обращение}": "{{ salutation }}",
        "{Обращение}": "{{ salutation }}",
        "[обращение]": "{{ salutation }}",
        "[Обращение}": "{{ salutation }}",
        "{содержимое_pdf}": "{{ содержимое_pdf }}",
        "{Содержимое_pdf}": "{{ содержимое_pdf }}",
        "{содержимое pdf}": "{{ содержимое_pdf }}",
        "{Содержимое pdf}": "{{ содержимое_pdf }}",
        "{pdf_content}": "{{ pdf_content }}",
        "{PDF_CONTENT}": "{{ pdf_content }}",
    }
    for simple_tag, jinja_tag in replacements.items():
        raw_text = raw_text.replace(simple_tag, jinja_tag)
    return raw_text


async def parse_request_files_and_recipients(
    send_mode: str,
    single_email: str | None,
    single_last_name: str | None,
    single_first_name: str | None,
    single_patronymic: str | None,
    excel_file: UploadFile | None,
    personal_docs: list[UploadFile]
) -> tuple[list[dict], dict[str, tuple[bytes, str]]]:
    """Вспомогательная функция для парсинга адресатов и персональных файлов."""
    recipients = []
    if send_mode == "single":
        if not single_email or not single_email.strip():
            raise ValueError("Не указан email получателя для одиночной отправки.")
        recipients = [{
            "email": single_email.strip(),
            "фамилия": (single_last_name or "").strip(),
            "имя": (single_first_name or "").strip(),
            "отчество": (single_patronymic or "").strip()
        }]
    else:
        if not excel_file:
            raise ValueError("Не прикреплен Excel-файл с таблицей получателей.")
        excel_content = await excel_file.read()
        recipients = read_excel_from_bytes(excel_content)
        if not recipients:
            raise ValueError("В Excel-файле не найдено ни одного корректного адресата.")

    personal_files_map = {}
    if personal_docs:
        for p_file in personal_docs:
            if p_file and p_file.filename:
                p_content = await p_file.read()
                if len(p_content) > 0:
                    personal_files_map[p_file.filename] = (p_content, p_file.filename)

    return recipients, personal_files_map


@router.post("/api/dry-run")
async def api_dry_run(
    subject: str = Form(""),
    letter_text: str | None = Form(default=None),
    send_mode: str = Form("mass"),
    single_email: str | None = Form(default=None),
    single_last_name: str | None = Form(default=None),
    single_first_name: str | None = Form(default=None),
    single_patronymic: str | None = Form(default=None),
    excel_file: UploadFile | None = File(default=None),
    personal_docs: list[UploadFile] = File(default=[])
):
    """
    Выполняет комплексный предварительный аудит (Dry Run) без реальной отправки.
    """
    try:
        recipients, personal_files_map = await parse_request_files_and_recipients(
            send_mode, single_email, single_last_name, single_first_name, single_patronymic,
            excel_file, personal_docs
        )
    except ValueError as ve:
        return JSONResponse({"status": "error", "message": str(ve)}, status_code=400)
    except Exception as e:
        return JSONResponse({"status": "error", "message": f"Ошибка обработки данных: {e}"}, status_code=500)

    check_results = run_preflight_check(
        recipients=recipients,
        personal_files_map=personal_files_map if personal_files_map else None,
        letter_text=letter_text or "",
        subject=subject or ""
    )
    return {"status": "ok", "report": check_results}


@router.post("/api/preview")
async def api_preview_email(
    subject: str = Form("Официальное письмо"),
    letter_text: str | None = Form(default=None),
    is_pure_html: bool = Form(False),
    send_mode: str = Form("mass"),
    single_email: str | None = Form(default=None),
    single_last_name: str | None = Form(default=None),
    single_first_name: str | None = Form(default=None),
    single_patronymic: str | None = Form(default=None),
    excel_file: UploadFile | None = File(default=None),
    personal_docs: list[UploadFile] = File(default=[]),
    attachments: list[UploadFile] = File(default=[]),
    inline_image: UploadFile | None = File(default=None),
    pdf_mode: str = Form("visual"),
    also_attach_pdf: bool = Form(True),
    recipient_index: int = Form(0)
):
    """
    Генерирует живой интерактивный предпросмотр письма для выбранного адресата.
    """
    try:
        recipients, personal_files_map = await parse_request_files_and_recipients(
            send_mode, single_email, single_last_name, single_first_name, single_patronymic,
            excel_file, personal_docs
        )
    except ValueError as ve:
        return JSONResponse({"status": "error", "message": str(ve)}, status_code=400)
    except Exception as e:
        return JSONResponse({"status": "error", "message": f"Ошибка обработки данных: {e}"}, status_code=500)

    if not recipients:
        return JSONResponse({"status": "error", "message": "Список получателей пуст"}, status_code=400)

    target_idx = max(0, min(recipient_index, len(recipients) - 1))
    person = recipients[target_idx]

    # Обработка общих файлов и баннера
    inline_img_bytes = None
    inline_img_name = None
    if inline_image and inline_image.filename:
        inline_img_bytes = await inline_image.read()
        inline_img_name = inline_image.filename

    attachments_data = []
    if attachments:
        for att in attachments:
            if att and att.filename:
                att_bytes = await att.read()
                if len(att_bytes) > 0:
                    attachments_data.append((att_bytes, att.filename))

    payload = build_email_payload(
        email_address="sender@donstu.ru",
        email_to=str(person.get("email") or "recipient@donstu.ru"),
        subject=subject or "Официальное письмо",
        letter_text=letter_text or "",
        is_pure_html=is_pure_html,
        person=person,
        inline_img_bytes=inline_img_bytes,
        inline_img_name=inline_img_name,
        attachments_data=attachments_data,
        personal_files_map=personal_files_map if personal_files_map else None,
        pdf_mode=pdf_mode,
        also_attach_pdf=also_attach_pdf,
        smtp_server="mail.donstu.ru",
        for_preview=True
    )

    return {
        "status": "ok",
        "recipient_index": target_idx,
        "total_recipients": len(recipients),
        "email_to": person.get("email"),
        "fio": payload["fio"],
        "salutation": payload["salutation"],
        "subject": subject,
        "html_body": payload["html_body"],
        "attached_filenames": payload["attached_filenames"],
        "personal_files_found": [f[1] for f in payload["personal_docs"]]
    }


@router.post("/start-broadcast")
async def start_broadcast(
    request: Request,
    background_tasks: BackgroundTasks,
    smtp_server: str = Form("mail.donstu.ru"),
    smtp_port: int = Form(25),
    email_address: str = Form(...),
    email_password: str = Form(...),
    delay_min: float = Form(1.0),
    delay_max: float = Form(2.5),
    subject: str = Form(...),
    letter_text: str | None = Form(default=None),
    send_mode: str = Form("mass"),
    single_email: str | None = Form(default=None),
    single_last_name: str | None = Form(default=None),
    single_first_name: str | None = Form(default=None),
    single_patronymic: str | None = Form(default=None),
    is_pure_html: bool = Form(False),
    excel_file: UploadFile | None = File(default=None),
    attachments: list[UploadFile] = File(default=[]),
    personal_docs: list[UploadFile] = File(default=[]),
    inline_image: UploadFile | None = File(default=None),
    pdf_mode: str = Form("visual"),
    also_attach_pdf: bool = Form(True),
    auto_retry: bool = Form(True)
):
    if broadcast_status["is_running"]:
        return JSONResponse({"status": "error", "message": "Рассылка уже запущена!"}, status_code=400)

    try:
        recipients, personal_files_map = await parse_request_files_and_recipients(
            send_mode, single_email, single_last_name, single_first_name, single_patronymic,
            excel_file, personal_docs
        )
    except ValueError as ve:
        return JSONResponse({"status": "error", "message": str(ve)}, status_code=400)
    except Exception as e:
        return JSONResponse({"status": "error", "message": f"Ошибка обработки файлов: {e}"}, status_code=500)

    inline_img_bytes = None
    inline_img_name = None
    if inline_image and inline_image.filename:
        inline_img_bytes = await inline_image.read()
        inline_img_name = inline_image.filename

    attachments_data = []
    if attachments:
        for att in attachments:
            if att and att.filename:
                att_bytes = await att.read()
                if len(att_bytes) > 0:
                    attachments_data.append((att_bytes, att.filename))

    cleaned_letter_text = letter_text or ""

    background_tasks.add_task(
        run_bg_broadcast,
        smtp_server=smtp_server,
        smtp_port=smtp_port,
        email_address=email_address,
        email_password=email_password,
        subject=subject,
        letter_text=cleaned_letter_text,
        is_pure_html=is_pure_html,
        recipients=recipients,
        inline_img_bytes=inline_img_bytes,
        inline_img_name=inline_img_name,
        attachments_data=attachments_data,
        delay_min=delay_min,
        delay_max=delay_max,
        personal_files_map=personal_files_map if personal_files_map else None,
        pdf_mode=pdf_mode,
        also_attach_pdf=also_attach_pdf,
        auto_retry=auto_retry
    )

    return {"status": "success", "message": f"Рассылка запущена для {len(recipients)} получателей."}
