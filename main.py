import os
import sys
import asyncio
import random
import re
import signal
import ssl
from io import BytesIO
import email.utils
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders

from fastapi import FastAPI, Form, UploadFile, File, Request, BackgroundTasks
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
import openpyxl
from aiosmtplib import SMTP as AsyncSMTP

# Настройка базовой директории для корректной работы путей внутри .exe (PyInstaller)
if getattr(sys, 'frozen', False):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

templates_path = os.path.join(base_dir, "templates")
templates = Jinja2Templates(directory=templates_path)

app = FastAPI(title="Email Broadcast Service")

# Глобальный словарь для отслеживания состояния текущей рассылки в реальном времени
broadcast_status = {
    "is_running": False,    # Флаг активного процесса рассылки
    "should_stop": False,   # Сигнал для экстренной остановки пользователем
    "total": 0,             # Общее количество адресатов из Excel
    "sent": 0,              # Успешно отправленные письма
    "skipped": 0,           # Пропущенные строки (например, если нет email)
    "errors": []            # Лог ошибок отправки (строка + текст ошибки)
}


# =====================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =====================================================================

def raw_text_to_html_with_merge(text_template: str, data_dict: dict, clean_cid: str = None) -> str:
    """
    Преобразует обычный текст в HTML-формат, подставляет переменные из Excel,
    обрабатывает Markdown-ссылки, сырые URL и добавляет встроенное изображение.
    """
    text = text_template if text_template else ""

    # Подстановка значений из Excel вместо плейсхолдеров {name}, {company} и т.д.
    for key, value in data_dict.items():
        placeholder = f"{{{key}}}"
        if placeholder in text:
            text = text.replace(placeholder, str(value if value is not None else ""))

    # Преобразование Markdown ссылок [текст](ссылка) в HTML тег <a>
    markdown_pattern = r'\[([^\]]+)\]\((https?://[^\s\)]+)\)'
    text = re.sub(markdown_pattern, r'<a href="\2" style="color: #1a73e8; text-decoration: underline;">\1</a>', text)

    # Преобразование сырых ссылок в кликабельные HTML-ссылки
    raw_url_pattern = r'(?<!href=")(https?://[^\s<>]+)'
    text = re.sub(raw_url_pattern, r'<a href="\1" style="color: #1a73e8; text-decoration: underline;">\1</a>', text)

    # Замена переносов строк на HTML-тег <br>
    html_ready_text = text.replace("\n", "<br>")

    # Формирование тега картинки, если она передана для отображения внутри письма
    img_tag = ""
    if clean_cid:
        img_tag = f'<img src="cid:{clean_cid}" alt="Изображение" style="max-width:100%; height:auto; display:block; margin-bottom:20px;"><br>'

    return f"""
    <div style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.5; color: #333333;">
        {img_tag}
        {html_ready_text}
    </div>
    """


def read_excel_from_bytes(file_bytes: bytes) -> list:
    """
    Читает файл Excel напрямую из оперативной памяти без сохранения на жесткий диск.
    """
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    sheet = wb.active
    rows = list(sheet.iter_rows(values_only=True))

    if not rows:
        return []

    # Очищаем заголовки колонок от случайных пробелов на концах
    headers = [str(cell).strip() for cell in rows[0]]
    recipients = []

    # Собираем строки в словари на основе заголовков
    for row in rows[1:]:
        if not any(row):  # Полностью пустые строки в Excel просто пропускаем
            continue
        recipients.append(dict(zip(headers, row)))

    wb.close()
    return recipients


# =====================================================================
# ФОНОВАЯ ЗАДАЧА ОТПРАВКИ
# =====================================================================

async def run_bg_broadcast(
    smtp_server: str,
    smtp_port: int,
    email_address: str,
    email_password: str,
    subject: str,
    letter_text: str,
    recipients: list,
    inline_img_bytes: bytes | None,
    inline_img_name: str | None,
    attachments_data: list[tuple[bytes, str]],
    delay_min: float,
    delay_max: float
):
    """
    Фоновый асинхронный процесс рассылки писем.
    Для каждого адресата создается изолированное подключение к SMTP,
    чтобы избежать разрыва длинных сессий по таймауту со стороны почтовых провайдеров.
    """
    global broadcast_status

    # Инициализируем и сбрасываем статус перед стартом новой сессии
    broadcast_status["is_running"] = True
    broadcast_status["should_stop"] = False
    broadcast_status["total"] = len(recipients)
    broadcast_status["sent"] = 0
    broadcast_status["skipped"] = 0
    broadcast_status["errors"] = []

    print(f"[РАССЫЛКА] Процесс запущен. Всего адресатов: {len(recipients)}")

    for index, person in enumerate(recipients, start=1):
        # Проверяем флаг экстренной остановки на каждом шаге цикла
        if broadcast_status["should_stop"]:
            print(f"[РАССЫЛКА] Процесс принудительно остановлен пользователем на шаге {index}!")
            break

        email_to = person.get("email")
        if not email_to:
            print(f"[ПРОПУСК] У строки №{index} в Excel не заполнен email.")
            broadcast_status["skipped"] += 1
            continue

        # Настройка типа подключения (SSL/TLS для порта 465, остальные — голые/STARTTLS)
        use_tls = (smtp_port == 465)
        
        
        tls_context = ssl.create_default_context()
        tls_context.check_hostname = False
        tls_context.verify_mode = ssl.CERT_NONE
        
        smtp_client = AsyncSMTP(
            hostname=smtp_server, 
            port=smtp_port, 
            use_tls=use_tls,
            tls_context=tls_context
        )

        try:
            await smtp_client.connect()

            # Если используются порты 25 или 587, мягко поднимаем шифрование через STARTTLS
            if smtp_port in (25, 587):
                try:
                    await smtp_client.starttls()
                except Exception as tls_err:
                    print(f"[ИНФО] Работаем без STARTTLS на порту {smtp_port}: {tls_err}")

            await smtp_client.login(email_address, email_password)

            # Генерация Content-ID для встроенного изображения
            image_cid = email.utils.make_msgid(domain="local") if inline_img_bytes else None
            clean_cid = image_cid.strip("<>") if image_cid else None

            # Подготовка HTML контента письма
            html_body = raw_text_to_html_with_merge(letter_text, person, clean_cid=clean_cid)

            # Создание структуры многокомпонентного письма
            msg_root = MIMEMultipart("mixed")
            msg_root["Subject"] = subject
            msg_root["From"] = email_address
            msg_root["To"] = email_to

            # Контейнер для HTML-текста и встроенных картинок
            msg_html_group = MIMEMultipart("related")
            msg_html_part = MIMEText(html_body, "html", "utf-8")
            msg_html_group.attach(msg_html_part)

            # Если картинка передана, прикрепляем её с уникальным Content-ID для отображения внутри HTML
            if inline_img_bytes and image_cid:
                subtype = inline_img_name.split(".")[-1] if inline_img_name and "." in inline_img_name else "jpeg"
                img_part = MIMEImage(inline_img_bytes, _subtype=subtype)
                img_part.add_header("Content-ID", image_cid)
                img_part.add_header("Content-Disposition", "inline", filename=inline_img_name)
                msg_html_group.attach(img_part)

            msg_root.attach(msg_html_group)

            # Добавление обычных файлов-вложений (документы, архивы и т.д.)
            for file_bytes, filename in attachments_data:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(file_bytes)
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", "attachment", filename=filename)
                msg_root.attach(part)

            # Отправка сформированного письма
            await smtp_client.send_message(msg_root)
            broadcast_status["sent"] += 1
            print(f"[ОТПРАВЛЕНО] [{index}/{len(recipients)}] Успешно ушло на {email_to}")

        except Exception as err:
            # Фиксируем ошибку, чтобы фронтенд мог её считать, и спокойно продолжаем цикл
            error_msg = f"Строка №{index} ({email_to}): {err}"
            broadcast_status["errors"].append(error_msg)
            print(f"[ОШИБКА] {error_msg}")

        finally:
            # Вежливо закрываем сессию с сервером
            try:
                await smtp_client.quit()
            except Exception:
                pass

        # Случайная пауза между отправками (не делаем паузу после самого последнего письма)
        if index < len(recipients) and not broadcast_status["should_stop"]:
            current_delay = random.uniform(delay_min, delay_max)
            await asyncio.sleep(current_delay)

    # Фоновый процесс завершил свою работу
    broadcast_status["is_running"] = False
    print(f"\n[ИТОГ] Рассылка полностью завершена. Отправлено: {broadcast_status['sent']}")


# =====================================================================
# ЭНДПОИНТЫ FASTAPI
# =====================================================================

@app.get("/", response_class=HTMLResponse)
async def index_page(request: Request):
    """Отображает главную страницу веб-интерфейса рассыльщика."""
    return templates.TemplateResponse(request=request, name="index.html", context={"message": None})


@app.post("/start-broadcast", response_class=HTMLResponse)
async def start_broadcast(
    request: Request,
    background_tasks: BackgroundTasks,
    smtp_server: str = Form('mail.donstu.ru'),
    smtp_port: int = Form(25),
    email_address: str = Form(...),
    email_password: str = Form(...),
    delay_min: float = Form(1.0),
    delay_max: float = Form(2.0),
    subject: str = Form(...),
    letter_text: str | None = Form(default=None),
    excel_file: UploadFile = File(...),
    inline_image: UploadFile = File(None),
    attachments: list[UploadFile] = File(None)
):
    """
    Принимает настройки и файлы из формы, проводит валидацию данных
    и запускает процесс отправки писем в фоновом режиме.
    """
    # Защита от случайного повторного запуска, пока идет текущая рассылка
    if broadcast_status["is_running"]:
        return templates.TemplateResponse(
            request=request,
            name="index.html",
            context={"message": "Ошибка: Рассылка уже запущена и выполняется прямо сейчас!"}
        )

    try:
        # Чтение и парсинг Excel-файла с базой адресатов
        file_content = await excel_file.read()
        recipients = read_excel_from_bytes(file_content)

        if not recipients:
            return templates.TemplateResponse(
                request=request,
                name="index.html",
                context={"message": "Ошибка: Загруженный файл Excel пустой или некорректный!"}
            )

        # Чтение встроенной в тело письма картинки (если есть)
        inline_img_bytes = None
        inline_img_name = None
        if inline_image and inline_image.filename:
            inline_img_bytes = await inline_image.read()
            inline_img_name = inline_image.filename

        # Чтение обычных файлов-вложений
        attachments_data = []
        if attachments:
            for attach in attachments:
                if attach.filename:
                    content = await attach.read()
                    attachments_data.append((content, attach.filename))

        # Передаем задачу в фоновый поток FastAPI, чтобы не блокировать UI веб-страницы
        background_tasks.add_task(
            run_bg_broadcast,
            smtp_server=smtp_server.strip(),
            smtp_port=smtp_port,
            email_address=email_address.strip(),
            email_password=email_password,
            subject=subject,
            letter_text=letter_text,
            recipients=recipients,
            inline_img_bytes=inline_img_bytes,
            inline_img_name=inline_img_name,
            attachments_data=attachments_data,
            delay_min=delay_min,
            delay_max=delay_max
        )

        status_msg = f"Рассылка успешно запущена для {len(recipients)} адресатов."

    except Exception as e:
        print(f"[ОШИБКА СЕРВЕРА] {e}")
        status_msg = f"Произошла критическая ошибка при обработке данных: {e}"

    return templates.TemplateResponse(request=request, name="index.html", context={"message": status_msg})


@app.get("/api/status")
async def get_status():
    """Возвращает текущий прогресс рассылки в формате JSON для AJAX-запросов с фронтенда."""
    return broadcast_status


@app.post("/api/stop")
async def stop_broadcast():
    """Выставляет флаг принудительной остановки активной рассылки."""
    if broadcast_status["is_running"]:
        broadcast_status["should_stop"] = True
        return {"status": "success", "message": "Сигнал остановки отправлен. Рассылка прервется на текущем шаге."}
    return {"status": "error", "message": "Рассылка не запущена в данный момент."}


@app.post("/shutdown")
async def shutdown_server():
    """Экстренно выключает процесс веб-сервера (полезно при работе внутри автономного .exe файла)."""
    print("[INFO] Получен сигнал на выключение. Завершаю работу сервера...")
    os.kill(os.getpid(), signal.SIGINT)
    return {"message": "Сервер успешно выключен"}


if __name__ == "__main__":
    import uvicorn
    import webbrowser

    # Автоматически открываем страницу приложения в браузере при старте
    webbrowser.open("http://127.0.0.1:8000")

    # Отключаем ANSI-цвета в логах Uvicorn для чистого вывода в консоль Windows (.exe)
    log_config = uvicorn.config.LOGGING_CONFIG
    if "formatters" in log_config:
        if "default" in log_config["formatters"]:
            log_config["formatters"]["default"]["use_colors"] = False
        if "access" in log_config["formatters"]:
            log_config["formatters"]["access"]["use_colors"] = False

    uvicorn.run(app, host="127.0.0.1", port=8000, log_config=log_config)