import re
from typing import Optional, List, Tuple

# Попытка инициализации морфологического анализатора pymorphy3
try:
    import pymorphy3
    morph = pymorphy3.MorphAnalyzer()
except Exception:
    morph = None


MALE_PATRONYMIC_ENDINGS = ("вич", "ьич", "тич", "мич", "гич")
FEMALE_PATRONYMIC_ENDINGS = ("вна", "чна", "шна", "инична")

MALE_NAMES_WITH_A = {
    "илья", "никита", "данила", "данил", "лёва", "лева", "фома", 
    "кузьма", "лука", "савва", "миша", "саша", "паша", "дима", 
    "юра", "вова", "ваня", "коля", "толя", "серёжа", "сережа"
}


def detect_salutation(first_name: str, patronymic: str) -> dict:
    """
    Определяет пол адресата и возвращает вежливое обращение.
    1. Приоритет анализу отчества по стандартным окончаниям.
    2. Анализ имени через pymorphy3 или грамматические правила.
    """
    first_name_clean = (first_name or "").strip()
    patronymic_clean = (patronymic or "").strip()

    pat_lower = patronymic_clean.lower()
    first_lower = first_name_clean.lower()

    # 1. Проверка по отчеству
    if pat_lower:
        if any(pat_lower.endswith(end) for end in MALE_PATRONYMIC_ENDINGS):
            return {"salutation": "Уважаемый", "gender": "masc"}
        if any(pat_lower.endswith(end) for end in FEMALE_PATRONYMIC_ENDINGS):
            return {"salutation": "Уважаемая", "gender": "fem"}

    # 2. Проверка через pymorphy3
    if morph and first_lower:
        parsed = morph.parse(first_lower)
        for p in parsed:
            if 'Name' in p.tag:
                if p.tag.gender == 'masc':
                    return {"salutation": "Уважаемый", "gender": "masc"}
                elif p.tag.gender == 'femn':
                    return {"salutation": "Уважаемая", "gender": "fem"}

    # 3. Эвристический анализ по имени
    if first_lower:
        if first_lower in MALE_NAMES_WITH_A:
            return {"salutation": "Уважаемый", "gender": "masc"}
        if first_lower.endswith(("а", "я", "ия")) and not first_lower.endswith(("илья", "никита")):
            return {"salutation": "Уважаемая", "gender": "fem"}
        if first_lower.endswith(("й", "л", "н", "р", "м", "в", "д", "г", "к", "п", "т", "б", "с")):
            return {"salutation": "Уважаемый", "gender": "masc"}

    return {"salutation": "Уважаемый(-ая)", "gender": "unknown"}


def normalize_word(word: str) -> str:
    """Приводит слово к начальной форме (лемме)."""
    clean = re.sub(r'[^а-яёa-z]', '', word.lower())
    if not clean:
        return ""
    if morph:
        try:
            return morph.parse(clean)[0].normal_form
        except Exception:
            pass
    return clean


def find_all_personal_docs(person: dict, personal_files_map: dict[str, tuple[bytes, str]]) -> List[Tuple[bytes, str]]:
    """
    Находит ВСЕ персональные файлы адресата (например, Word .docx И PDF .pdf одновременно).
    Сопоставляет по фамилии (включая склонения: «Купчуку», «Купчук») и первой букве имени.
    """
    if not personal_files_map:
        return []

    last_name = str(person.get("фамилия") or person.get("Фамилия") or person.get("last_name") or "").strip()
    first_name = str(person.get("имя") or person.get("Имя") or person.get("name") or "").strip()

    if not last_name:
        return []

    norm_last_name = normalize_word(last_name)
    first_initial = first_name[0].lower() if first_name else ""
    lower_last_name = last_name.lower()

    matched = []

    for filename, file_tuple in personal_files_map.items():
        base_name = filename.rsplit('.', 1)[0] if '.' in filename else filename
        tokens = re.findall(r'[а-яёa-z]+', base_name.lower())

        is_match = False
        for token in tokens:
            token_norm = normalize_word(token)
            if norm_last_name and (norm_last_name == token_norm or norm_last_name == token):
                if first_initial:
                    initials_in_file = [t for t in tokens if len(t) == 1]
                    if initials_in_file:
                        if first_initial in initials_in_file:
                            is_match = True
                            break
                    else:
                        is_match = True
                        break
                else:
                    is_match = True
                    break

        # Резервный поиск по прямому вхождению подстроки фамилии
        if not is_match and lower_last_name in filename.lower():
            is_match = True

        if is_match and file_tuple not in matched:
            matched.append(file_tuple)

    return matched


def find_personal_doc(person: dict, personal_files_map: dict[str, tuple[bytes, str]]) -> Optional[Tuple[bytes, str]]:
    """Для обратной совместимости: возвращает первый найденный персональный документ."""
    docs = find_all_personal_docs(person, personal_files_map)
    return docs[0] if docs else None
