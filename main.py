import os
import sys
import uvicorn
import io
from email.header import Header
import asyncio
import pymorphy3
from jinja2 import Template
import random
import re
import webbrowser
import ssl
from io import BytesIO
import email.utils
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
import threading
import time

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

# Инициализируем анализатор один раз при запуске сервера
morph = pymorphy3.MorphAnalyzer()

import re

# Проверь, чтобы в начале main.py был объявлен морфологический анализатор:
# morph = pymorphy3.MorphAnalyzer()

def extract_person_key(text: str) -> str:
    """
    Превращает любую строку ('Аветисяну А.И.', 'Аветисян Арутюн') 
    в единый ключ вида 'аветисян_а'
    """
    if not text:
        return ""
    # Очищаем от расширения файла и лишних знаков
    clean = re.sub(r'\.(docx|doc|pdf)$', '', str(text), flags=re.IGNORECASE)
    clean = clean.replace('.', ' ').replace('_', ' ').replace('-', ' ')
    parts = clean.split()
    
    if not parts:
        return ""
    
    # 1. Фамилию приводим к начальной форме через pymorphy3 (Аветисяну -> аветисян)
    raw_surname = parts[0]
    parsed = morph.parse(raw_surname)
    norm_surname = parsed[0].normal_form.lower() if parsed else raw_surname.lower()
    
    # 2. Первую букву имени берем как инициал ('а' для 'А.И.' или 'а' для 'Арутюн')
    first_initial = parts[1][0].lower() if len(parts) > 1 else ""
    
    return f"{norm_surname}_{first_initial}".rstrip("_")


def find_personal_doc(person: dict, personal_files: dict[str, tuple[bytes, str]]) -> tuple[bytes, str] | None:
    """Мгновенно находит файл по сформированному ключу"""
    if not personal_files:
        return None

    # Достаем строку с именем из Excel (ФИО или связка Фамилия + Имя)
    excel_fio = (
        person.get("ФИО") or person.get("фио") or 
        f"{person.get('Фамилия', '')} {person.get('Имя', '')}"
    ).strip()

    target_key = extract_person_key(excel_fio)
    if not target_key:
        return None

    # Ищем в загруженных файлах тот, чей ключ совпал
    for orig_name, file_tuple in personal_files.items():
        if extract_person_key(orig_name) == target_key:
            return file_tuple

    return None

async def watchdog_task():
    """Фоновый сторож: убивает процесс, если вкладка закрыта и рассылка не идет"""
    global last_heartbeat_time
    while True:
        await asyncio.sleep(3)
        now = time.time()
        
        # Если вкладка молчит больше 8 секунд и рассылка не активна:
        if (now - last_heartbeat_time > 8.0) and not broadcast_status.get("is_running", False):
            print("\n[АВТОСТОП] Вкладка браузера закрыта. Завершение работы процесса...")
            os._exit(0)



import unicodedata

def normalize_string(val: str) -> str:
    """Удаляет лишние символы, расширения и пробелы для надежного сравнения"""
    if not val:
        return ""
    # Убираем расширения .docx / .doc
    val = re.sub(r'\.(docx|doc|pdf)$', '', val, flags=re.IGNORECASE)
    # Заменяем подчёркивания и дефисы на пробелы
    val = val.replace('_', ' ').replace('-', ' ')
    # Очищаем от лишних пробелов и переводим в нижний регистр
    val = " ".join(val.split()).lower()
    return val

def find_personal_attachment(person_dict: dict, personal_files: dict[str, tuple[bytes, str]]) -> tuple[bytes, str] | None:
    """
    Ищет персональный файл по ФИО, Фамилии+Имени или отдельным колонкам.
    personal_files — словарь: {normalized_filename: (file_bytes, original_filename)}
    """
    # Собираем возможные варианты имени человека из колонок таблицы
    fio = person_dict.get("ФИО") or person_dict.get("фио") or ""
    first = person_dict.get("Имя") or person_dict.get("имя") or ""
    patronymic = person_dict.get("Отчество") or person_dict.get("отчество") or ""
    last = person_dict.get("Фамилия") or person_dict.get("фамилия") or ""

    candidates = []
    if fio:
        candidates.append(normalize_string(fio))
    if last and first and patronymic:
        candidates.append(normalize_string(f"{last} {first} {patronymic}"))
        candidates.append(normalize_string(f"{first} {patronymic} {last}"))
    if last and first:
        candidates.append(normalize_string(f"{last} {first}"))
        candidates.append(normalize_string(f"{first} {last}"))
    if first:
        candidates.append(normalize_string(first))

    # 1. Точное совпадение
    for cand in candidates:
        if cand in personal_files:
            return personal_files[cand]

    # 2. Нестрогое совпадение (если имя кандидата содержится в названии файла)
    for cand in candidates:
        if len(cand) >= 4:  # Чтобы не матчить слишком короткие слова
            for norm_filename, file_tuple in personal_files.items():
                if cand in norm_filename:
                    return file_tuple

    return None

def detect_salutation(name: str, patronymic: str = "") -> dict:
    """Определяет пол и правильное обращение по Имени и Отчеству"""
    name_clean = str(name).strip() if name and str(name).lower() != 'none' else ""
    patronymic_clean = str(patronymic).strip() if patronymic and str(patronymic).lower() != 'none' else ""

    # Сначала проверяем отчество
    if patronymic_clean:
        p_lower = patronymic_clean.lower()
        if p_lower.endswith(("вна", "чна", "ична", "ычна")):
            return {"salutation": "Уважаемая", "gender": "female"}
        if p_lower.endswith(("вич", "ич", "ыч")):
            return {"salutation": "Уважаемый", "gender": "male"}

    # Если отчества нет, смотрим на имя
    if name_clean:
        parsed = morph.parse(name_clean)[0]
        if parsed.tag.gender == "femn":
            return {"salutation": "Уважаемая", "gender": "female"}
        if parsed.tag.gender == "masc":
            return {"salutation": "Уважаемый", "gender": "male"}

    # Резервный вариант на случай нейтральных или неопознанных имен
    return {"salutation": "Уважаемый(-ая)", "gender": "unknown"}


# =====================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =====================================================================

def render_email_content(template_str: str, data_dict: dict, is_pure_html: bool = False, clean_cid: str = None) -> str:
    """
    Рендерит тело письма через Jinja2, автоматически добавляет приветствие сверху,
    если оно не прописано вручную.
    """
    text = template_str if template_str else ""

    # 1. Автоматическая генерация вступительной строки с обращением
    # Проверяем, не написал ли пользователь обращение вручную
    has_manual_greeting = any(word in text.lower() for word in ["уважаем", "{{ salutation }}", "{salutation}"])
    
    if not has_manual_greeting and not is_pure_html:
        salutation = data_dict.get("salutation", "Уважаемый(-ая)")
        first_name = data_dict.get("first_name", "")
        patronymic = data_dict.get("patronymic", "")
        last_name = data_dict.get("Фамилия") or data_dict.get("фамилия") or data_dict.get("last_name") or ""

        # Собираем обращение в зависимости от того, что заполнено в Excel:
        # Приоритет 1: Имя + Отчество
        if first_name and patronymic:
            greeting_line = f"{salutation} {first_name} {patronymic}!\n\n"
        # Приоритет 2: Имя + Фамилия
        elif first_name and last_name:
            greeting_line = f"{salutation} {first_name} {last_name}!\n\n"
        # Приоритет 3: Только Имя
        elif first_name:
            greeting_line = f"{salutation} {first_name}!\n\n"
        else:
            greeting_line = "Здравствуйте!\n\n"

        # Приклеиваем приветствие перед основным текстом
        text = greeting_line + text

    try:
        jinja_tpl = Template(text)
        rendered_text = jinja_tpl.render(data_dict)
    except Exception as e:
        rendered_text = text
        print(f"[ПРЕДУПРЕЖДЕНИЕ] Ошибка Jinja2-рендеринга: {e}")

    # 3. Обратная совместимость для скобок {Имя}
    for key, value in data_dict.items():
        placeholder = f"{{{key}}}"
        if placeholder in rendered_text:
            rendered_text = rendered_text.replace(placeholder, str(value if value is not None else ""))

    if is_pure_html:
        return rendered_text

    # 4. Преобразование ссылок и переносов
    markdown_pattern = r'\[([^\]]+)\]\((https?://[^\s\)]+)\)'
    rendered_text = re.sub(markdown_pattern, r'<a href="\2" style="color: #1a73e8; text-decoration: underline;">\1</a>', rendered_text)

    raw_url_pattern = r'(?<!href=")(https?://[^\s<>]+)'
    rendered_text = re.sub(raw_url_pattern, r'<a href="\1" style="color: #1a73e8; text-decoration: underline;">\1</a>', rendered_text)

    html_ready_text = rendered_text.replace("\n", "<br>")

    # 5. Картинка-баннер
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
    Читает файл Excel напрямую из памяти без сохранения на диск.
    """
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    sheet = wb.active
    rows = list(sheet.iter_rows(values_only=True))

    if not rows:
        return []

    headers = [str(cell).strip() for cell in rows[0]]
    recipients = []

    for row in rows[1:]:
        if not any(row):
            continue
        recipients.append(dict(zip(headers, row)))

    wb.close()
    return recipients


# =====================================================================
# ФОНОВАЯ ЗАДАЧА ОТПРАВКИ
# =====================================================================

def normalize_name(text: str) -> str:
    """Очищает строку от расширений, спецсимволов и лишних пробелов для точного сравнения"""
    if not text:
        return ""
    # Убираем расширения .docx, .doc, .pdf
    text = re.sub(r'\.(docx|doc|pdf)$', '', str(text), flags=re.IGNORECASE)
    # Заменяем подчёркивания и дефисы на пробелы
    text = text.replace('_', ' ').replace('-', ' ')
    # Убираем повторяющиеся пробелы и переводим в нижний регистр
    return " ".join(text.split()).lower()

def find_personal_doc(person: dict, personal_files: dict[str, tuple[bytes, str]]) -> tuple[bytes, str] | None:
    """Ищет личный файл человека среди загруженных документов"""
    if not personal_files:
        return None

    fio = person.get("фио") or person.get("ФИО") or ""
    first = person.get("имя") or person.get("Имя") or person.get("name") or person.get("Name") or ""
    patronymic = person.get("отчество") or person.get("Отчество") or person.get("patronymic") or ""
    last = person.get("фамилия") or person.get("Фамилия") or person.get("last_name") or ""

    # Генерируем возможные варианты названия файла для этого человека
    candidates = []
    if fio:
        candidates.append(normalize_name(fio))
    if last and first and patronymic:
        candidates.append(normalize_name(f"{last} {first} {patronymic}"))
        candidates.append(normalize_name(f"{first} {patronymic} {last}"))
    if last and first:
        candidates.append(normalize_name(f"{last} {first}"))
        candidates.append(normalize_name(f"{first} {last}"))
    if first and patronymic:
        candidates.append(normalize_name(f"{first} {patronymic}"))
    if first:
        candidates.append(normalize_name(first))

    # 1. Точное совпадение
    for cand in candidates:
        if cand in personal_files:
            return personal_files[cand]

    # 2. Нестрогое вхождение (если имя есть внутри длинного названия файла)
    for cand in candidates:
        if len(cand) >= 4:
            for norm_filename, file_tuple in personal_files.items():
                if cand in norm_filename:
                    return file_tuple

    return None


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
    personal_files_map: dict[str, tuple[bytes, str]] = None  # <--- Принимает словарь персональных файлов
):
    global broadcast_status

    broadcast_status["is_running"] = True
    broadcast_status["should_stop"] = False
    broadcast_status["total"] = len(recipients)
    broadcast_status["sent"] = 0
    broadcast_status["skipped"] = 0
    broadcast_status["errors"] = []

    print(f"[РАССЫЛКА] Процесс запущен. Всего адресатов: {len(recipients)}")
    if personal_files_map:
        print(f"[ИНФО] Доступно персональных файлов для подстановки: {len(personal_files_map)}")

    for index, person in enumerate(recipients, start=1): 
        if broadcast_status["should_stop"]:
            print(f"[РАССЫЛКА] Процесс принудительно остановлен пользователем на шаге {index}!")
            break

        # Поиск email в разных возможных регистрах написания колонок
        email_to = person.get("email") or person.get("Email") or person.get("EMAIL")
        if not email_to:
            print(f"[ПРОПУСК] У строки №{index} в Excel не заполнен email.")
            broadcast_status["skipped"] += 1
            continue

        email_to = str(email_to).strip()

        # Извлечение Имени и Отчества для определения обращения
        first_name = person.get("имя") or person.get("Имя") or person.get("name") or person.get("Name") or ""
        patronymic = person.get("отчество") or person.get("Отчество") or person.get("patronymic") or ""

        # Вычисляем обращение и род
        gender_data = detect_salutation(str(first_name), str(patronymic))

        # Обогащаем словарь контакта вычисленными переменными
        merged_context = {**person}
        merged_context.update({
            "salutation": gender_data["salutation"],
            "gender": gender_data["gender"],
            "first_name": first_name,
            "patronymic": patronymic,
        })

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

            if smtp_port in (25, 587):
                try:
                    await smtp_client.starttls()
                except Exception as tls_err:
                    print(f"[ИНФО] Работаем без STARTTLS на порту {smtp_port}: {tls_err}")

            await smtp_client.login(email_address, email_password)

            # Генерация Content-ID для встроенного изображения
            image_cid = email.utils.make_msgid(domain="local") if inline_img_bytes else None
            clean_cid = image_cid.strip("<>") if image_cid else None

            # Подготовка динамического тела письма с авто-обращением
            html_body = render_email_content(
                template_str=letter_text,
                data_dict=merged_context,
                is_pure_html=is_pure_html,
                clean_cid=clean_cid
            )

            # Создание структуры письма
            msg_root = MIMEMultipart("mixed")
            msg_root["Subject"] = subject
            msg_root["From"] = email_address
            msg_root["To"] = email_to
            msg_root["Date"] = email.utils.formatdate(localtime=True)
            msg_root["Message-ID"] = email.utils.make_msgid(domain=smtp_server)

            msg_html_group = MIMEMultipart("related")
            msg_html_part = MIMEText(html_body, "html", "utf-8")
            msg_html_group.attach(msg_html_part)

            if inline_img_bytes and image_cid:
                subtype = inline_img_name.split(".")[-1] if inline_img_name and "." in inline_img_name else "jpeg"
                img_part = MIMEImage(inline_img_bytes, _subtype=subtype)
                img_part.add_header("Content-ID", image_cid)
                img_part.add_header("Content-Disposition", "inline", filename=inline_img_name)
                msg_html_group.attach(img_part)

            msg_root.attach(msg_html_group)

            # 1. Прикрепление ОБЩИХ файлов (для всех адресатов)
            for file_bytes, filename in attachments_data:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(file_bytes)
                encoders.encode_base64(part)
                encoded_fn = Header(filename, 'utf-8').encode()
                part.add_header("Content-Disposition", "attachment", filename=encoded_fn)
                msg_root.attach(part)

            # 2. ПОИСК И ПРИКРЕПЛЕНИЕ ПЕРСОНАЛЬНОГО WORD-ФАЙЛА 📄👤
            if personal_files_map:
                personal_doc = find_personal_doc(person, personal_files_map)
                if personal_doc:
                    p_bytes, p_name = personal_doc
                    p_part = MIMEBase("application", "vnd.openxmlformats-officedocument.wordprocessingml.document")
                    p_part.set_payload(p_bytes)
                    encoders.encode_base64(p_part)
                    # Корректно экранируем русские символы в названии файла
                    encoded_p_name = Header(p_name, 'utf-8').encode()
                    p_part.add_header("Content-Disposition", "attachment", filename=encoded_p_name)
                    msg_root.attach(p_part)
                    print(f"   [ФАЙЛ] Прикреплен персональный документ: {p_name}")
                else:
                    print(f"   [ВНИМАНИЕ] Личный файл для {email_to} не найден среди загруженных!")

            # Отправка письма
            await smtp_client.send_message(msg_root)
            broadcast_status["sent"] += 1
            print(f"[ОТПРАВЛЕНО] [{index}/{len(recipients)}] {gender_data['salutation']} -> {email_to}")

        except Exception as err:
            error_msg = f"Строка №{index} ({email_to}): {err}"
            broadcast_status["errors"].append(error_msg)
            print(f"[ОШИБКА] {error_msg}")

        finally:
            try:
                await smtp_client.quit()
            except Exception:
                pass

        if index < len(recipients) and not broadcast_status["should_stop"]:
            current_delay = random.uniform(delay_min, delay_max)
            await asyncio.sleep(current_delay)

    broadcast_status["is_running"] = False
    print(f"\n[ИТОГ] Рассылка полностью завершена. Отправлено: {broadcast_status['sent']}")

# =====================================================================
# ЭНДПОИНТЫ FASTAPI
# =====================================================================

@app.get("/", response_class=HTMLResponse)
async def index_page(request: Request):
    return templates.TemplateResponse(request=request, name="index.html", context={"message": None})

@app.on_event("startup")
async def start_watchdog():
    # Запускаем сторожа вместе с сервером
    asyncio.create_task(watchdog_task())

@app.post("/api/heartbeat")
async def receive_heartbeat():
    """Эндпоинт, куда браузер отстукивает пульс"""
    global last_heartbeat_time
    last_heartbeat_time = time.time()
    return {"status": "alive"}


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
    html_template: UploadFile = File(None),
    inline_image: UploadFile = File(None),
    attachments: list[UploadFile] = File(None),
    personal_docs: list[UploadFile] = File(None)
):
    if broadcast_status["is_running"]:
        return templates.TemplateResponse(
            request=request,
            name="index.html",
            context={"message": "Ошибка: Рассылка уже запущена и выполняется прямо сейчас!"}
        )

    try:
        file_content = await excel_file.read()
        recipients = read_excel_from_bytes(file_content)

        if not recipients:
            return templates.TemplateResponse(
                request=request,
                name="index.html",
                context={"message": "Ошибка: Загруженный файл Excel пустой или некорректный!"}
            )

        # Выбираем источник тела письма: загруженный файл .html или поле textarea
        is_pure_html = False
        final_letter_template = letter_text or ""
        
        if html_template and html_template.filename:
            raw_html_bytes = await html_template.read()
            if raw_html_bytes:
                final_letter_template = raw_html_bytes.decode("utf-8", errors="ignore")
                is_pure_html = True

        inline_img_bytes = None
        inline_img_name = None
        if inline_image and inline_image.filename:
            inline_img_bytes = await inline_image.read()
            inline_img_name = inline_image.filename

        attachments_data = []
        if attachments:
            for attach in attachments:
                if attach.filename:
                    content = await attach.read()
                    attachments_data.append((content, attach.filename))

        personal_files_map = {}
        if personal_docs:
            for pfile in personal_docs:
                if pfile.filename:
                    p_content = await pfile.read()
                    personal_files_map[pfile.filename] = (p_content, pfile.filename)

        background_tasks.add_task(
            run_bg_broadcast,
            smtp_server=smtp_server.strip(),
            smtp_port=smtp_port,
            email_address=email_address.strip(),
            email_password=email_password,
            subject=subject,
            letter_text=final_letter_template,
            is_pure_html=is_pure_html,
            recipients=recipients,
            inline_img_bytes=inline_img_bytes,
            inline_img_name=inline_img_name,
            attachments_data=attachments_data,
            personal_files_map=personal_files_map,
            delay_min=delay_min,
            delay_max=delay_max,
        )

        status_msg = f"Рассылка успешно запущена для {len(recipients)} адресатов."

    except Exception as e:
        print(f"[ОШИБКА СЕРВЕРА] {e}")
        status_msg = f"Произошла критическая ошибка при обработке данных: {e}"

    return templates.TemplateResponse(request=request, name="index.html", context={"message": status_msg})


@app.get("/api/status")
async def get_status():
    return broadcast_status


@app.post("/api/stop")
async def stop_broadcast():
    if broadcast_status["is_running"]:
        broadcast_status["should_stop"] = True
        return {"status": "success", "message": "Сигнал остановки отправлен. Рассылка прервется на текущем шаге."}
    return {"status": "error", "message": "Рассылка не запущена в данный момент."}


@app.post("/shutdown")
async def shutdown_server():
    def kill_process():
        time.sleep(0.3)
        os._exit(0)

    threading.Thread(target=kill_process, daemon=True).start()
    return {"status": "success", "message": "Сервер остановлен"}


def open_browser():
    time.sleep(1.5)
    webbrowser.open("http://127.0.0.1:8000")


if __name__ == "__main__":
    if sys.stdout is None:
        sys.stdout = open(os.devnull, 'w')
    if sys.stderr is None:
        sys.stderr = open(os.devnull, 'w')

    threading.Thread(target=open_browser, daemon=True).start()

    log_config = uvicorn.config.LOGGING_CONFIG
    if "formatters" in log_config:
        if "default" in log_config["formatters"]:
            log_config["formatters"]["default"]["use_colors"] = False
        if "access" in log_config["formatters"]:
            log_config["formatters"]["access"]["use_colors"] = False

    uvicorn.run(app, host="127.0.0.1", port=8000, log_config=log_config)