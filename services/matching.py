import re
import pymorphy3

morph = pymorphy3.MorphAnalyzer()

def normalize_word(word: str) -> str:
    """Приводит слово (в т.ч. фамилию в косвенном падеже) к нормальной форме."""
    if not word:
        return ""
    parsed = morph.parse(word.strip())
    return parsed[0].normal_form.lower() if parsed else word.strip().lower()

def extract_person_key(text: str) -> str:
    """
    Превращает любую строку:
    - 'Иванову И.docx' -> 'иванов_и'
    - 'Иванов Иван'     -> 'иванов_и'
    - 'Иван Иванов'     -> 'иванов_и'
    - 'Иванову И.И.'   -> 'иванов_и'
    """
    if not text:
        return ""

    # Срезаем расширения файлов
    clean = re.sub(r"\.(docx|doc|pdf)$", "", str(text), flags=re.IGNORECASE)
    # Заменяем точки, дефисы, подчеркивания и запятые на пробелы
    clean = re.sub(r"[._\-,]", " ", clean)
    parts = clean.split()

    if not parts:
        return ""

    # Ищем, кто из слов фамилия, а кто имя/инициал
    surname = ""
    initial = ""

    # Если передан формат 'Фамилия Инициалы' или 'Имя Фамилия'
    if len(parts) == 1:
        return normalize_word(parts[0])

    # Проверяем токены: обычно слово длиной 1 символ — это инициал
    words_norm = []
    for p in parts:
        parsed = morph.parse(p)[0]
        words_norm.append((p, parsed))

    # Пытаемся эвристически определить фамилию (обычно Surn или первая при стандартной записи)
    found_surn = None
    for raw, p in words_norm:
        if "Surn" in p.tag:
            found_surn = p.normal_form.lower()
            # Инициалом берем первую букву следующего или предыдущего слова
            other_parts = [r for r, _ in words_norm if r.lower() != raw.lower()]
            if other_parts:
                initial = other_parts[0][0].lower()
            break

    if found_surn:
        surname = found_surn
    else:
        # Стандартное предположение: Первое слово — фамилия, второе — имя/инициал
        surname = normalize_word(parts[0])
        initial = parts[1][0].lower()

    return f"{surname}_{initial}".rstrip("_")


def find_personal_doc(person: dict, personal_files_map: dict[str, tuple[bytes, str]]) -> tuple[bytes, str] | None:
    """
    Находит персональный файл в словаре personal_files_map:
    ключ словаря - оригинальное имя файла, значение - (bytes, filename)
    """
    if not personal_files_map:
        return None

    # Достаем имя из всех возможных вариантов колонок в Excel
    last_name = person.get("фамилия") or person.get("Фамилия") or person.get("last_name") or ""
    first_name = person.get("имя") or person.get("Имя") or person.get("name") or person.get("Name") or ""
    patronymic = person.get("отчество") or person.get("Отчество") or person.get("patronymic") or ""

    raw_name = (
        person.get("ФИО") or person.get("фио") or
        f"{last_name} {first_name} {patronymic}".strip() or
        f"{last_name} {first_name}".strip() or
        ""
    )

    target_key = extract_person_key(raw_name)
    if not target_key:
        return None

    # 1. Поиск по точному совпадению нормализованного ключа (иванов_и == иванов_и)
    for orig_filename, file_tuple in personal_files_map.items():
        if extract_person_key(orig_filename) == target_key:
            return file_tuple

    # 2. Фолбэк: если в файле только фамилия (например, "Иванову.docx")
    surname_only = target_key.split("_")[0]
    if len(surname_only) >= 4:
        for orig_filename, file_tuple in personal_files_map.items():
            file_key = extract_person_key(orig_filename)
            if file_key.split("_")[0] == surname_only:
                return file_tuple

    return None


def detect_salutation(name: str, patronymic: str = "") -> dict:
    """Определяет пол и вежливое обращение по имени и отчеству."""
    name_clean = str(name).strip() if name and str(name).lower() != "none" else ""
    patronymic_clean = str(patronymic).strip() if patronymic and str(patronymic).lower() != "none" else ""

    if patronymic_clean:
        p_lower = patronymic_clean.lower()
        if p_lower.endswith(("вна", "чна", "ична", "ычна")):
            return {"salutation": "Уважаемая", "gender": "female"}
        if p_lower.endswith(("вич", "ич", "ыч")):
            return {"salutation": "Уважаемый", "gender": "male"}

    if name_clean:
        parsed = morph.parse(name_clean)[0]
        if parsed.tag.gender == "femn":
            return {"salutation": "Уважаемая", "gender": "female"}
        if parsed.tag.gender == "masc":
            return {"salutation": "Уважаемый", "gender": "male"}

    return {"salutation": "Уважаемый(-ая)", "gender": "unknown"}