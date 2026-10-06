import os
import time
import threading
from fastapi import APIRouter, Form, UploadFile, File, Request, BackgroundTasks
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from fastapi.responses import HTMLResponse, JSONResponse

import config
from config import broadcast_status, TEMPLATES_DIR
from services.excel import read_excel_from_bytes
from services.mailer import run_bg_broadcast

router = APIRouter()
templates = Jinja2Templates(directory=TEMPLATES_DIR)

@router.get("/", response_class=HTMLResponse)
async def index_page(request: Request):
    return templates.TemplateResponse(request=request, name="index.html", context={"message": None})

@router.get("/api/status")
async def get_status():
    return broadcast_status

@router.post("/api/stop")
async def stop_broadcast():
    if broadcast_status["is_running"]:
        broadcast_status["should_stop"] = True
        return {"status": "success", "message": "Рассылка прервется на текущем шаге."}
    return {"status": "error", "message": "Рассылка не запущена."}

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
        "{Отчество}": "{{ patronymic }}",
        "{фамилия}": "{{ last_name }}",
        "{Фамилия}": "{{ last_name }}",
        "[фамилия]": "{{ last_name }}",
        "[Фамилия}": "{{ last_name }}",
        "{обращение}": "{{ salutation }}",
        "{Обращение}": "{{ salutation }}",
        "[обращение]": "{{ salutation }}",
        "[Обращение}": "{{ salutation }}",
    }
    for simple_tag, jinja_tag in replacements.items():
        raw_text = raw_text.replace(simple_tag, jinja_tag)
    return raw_text


@router.post("/start-broadcast")
async def start_broadcast(
    request: Request,
    background_tasks: BackgroundTasks,
    smtp_server: str = Form("mail.donstu.ru"),
    smtp_port: int = Form(25),
    email_address: str = Form(...),
    email_password: str = Form(...),
    delay_min: float = Form(1.0),
    delay_max: float = Form(2.0),
    subject: str = Form(...),
    letter_text: str | None = Form(default=None),
    excel_file: UploadFile = File(...),
    html_template: UploadFile = File(None),
    inline_image: UploadFile = File(None),
    attachments: list[UploadFile] = File(default=[]),
    personal_docs: list[UploadFile] = File(default=[]),
):
    print(f"DEBUG: Сырой список personal_docs: {personal_docs}")
    print(f"DEBUG: Количество файлов: {len(personal_docs)}")
    for doc in personal_docs:
        print(f"DEBUG: Файл -> {doc.filename}")
    
    if broadcast_status["is_running"]:
        return templates.TemplateResponse(request=request, name="index.html", context={"message": "Рассылка уже активна!"})

    try:
        file_content = await excel_file.read()
        recipients = read_excel_from_bytes(file_content)
        if not recipients:
            return templates.TemplateResponse(request=request, name="index.html", context={"message": "Excel пуст или поврежден!"})

        is_pure_html = False
        final_letter_template = normalize_template_text(letter_text)
        if html_template and html_template.filename:
            raw_html = await html_template.read()
            if raw_html:
                final_letter_template = raw_html.decode("utf-8", errors="ignore")
                is_pure_html = True

        inline_img_bytes = await inline_image.read() if inline_image and inline_image.filename else None
        inline_img_name = inline_image.filename if inline_img_bytes else None

        # 1. Персональные файлы (Word / PDF)
        personal_files_map = {}
        if personal_docs:
            for pfile in personal_docs:
                if pfile and pfile.filename and pfile.filename.strip():
                    await pfile.seek(0)                  # 👈 Возвращаем каретку в начало!
                    content = await pfile.read()
                    if content:
                        # ВАЖНО: кортеж строго (байты, имя)
                        personal_files_map[pfile.filename] = (content, pfile.filename)
                        print(f"[DEBUG] Персональный файл '{pfile.filename}' загружен: {len(content)} байт")

        # 2. Общие файлы вложений
        attachments_data = []
        if attachments:
            for att in attachments:
                if att and att.filename and att.filename.strip():
                    await att.seek(0)                  # 👈 Возвращаем каретку в начало!
                    att_content = await att.read()
                    if att_content:
                        # ВАЖНО: кортеж строго (байты, имя)
                        attachments_data.append((att_content, att.filename))
                        print(f"[DEBUG] Общий файл '{att.filename}' загружен: {len(att_content)} байт")

        print(f"ИТОГОВЫЙ РАЗМЕР personal_files_map = {len(personal_files_map)}")

        background_tasks.add_task(
            run_bg_broadcast,
            smtp_server.strip(),
            smtp_port,
            email_address.strip(),
            email_password,
            subject,
            final_letter_template,
            is_pure_html,
            recipients,
            inline_img_bytes,
            inline_img_name,
            attachments_data,
            delay_min,            # 12-й: строго задержка min
            delay_max,            # 13-й: строго задержка max
            personal_files_map    # 14-й: строго словарь файлов!
        )
        msg = f"Рассылка успешно запущена для {len(recipients)} адресатов."
    except Exception as e:
        msg = f"Ошибка: {e}"
        return JSONResponse({"status": "error", "message": msg}, status_code=500)

    return JSONResponse({"status": "ok", "message": msg})