import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import io
import os
import pytesseract
from PIL import Image
import re
import cv2
import numpy as np

# =========================================================
# НАСТРОЙКИ
# =========================================================

# ПРИ ПУБЛИКАЦИИ НА GITHUB ЭТА СТРОКА ДОЛЖНА БЫТЬ ЗАКОММЕНТИРОВАНА!
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

TEMPLATE_NAME = "образец отчета.docx"

st.set_page_config(page_title="Генератор Отчетов - Гарант Оценка", layout="wide")

# ── Инициализация session state ────────────────────────────────────────────────
keys = ['auto_vin', 'auto_reg', 'auto_year', 'auto_model', 'auto_fio',
        'auto_engine', 'auto_color', 'auto_address', 'auto_tech_passport', 'auto_body_type']
for k in keys:
    if k not in st.session_state:
        st.session_state[k] = ""
if 'auto_steering' not in st.session_state:
    st.session_state.auto_steering = "Левый руль"

# =========================================================
# БАЗЫ ДАННЫХ
# =========================================================

CAR_BRANDS = [
    "TOYOTA", "HONDA", "HYUNDAI", "KIA", "LEXUS", "BMW", "MERCEDES-BENZ", "MERCEDES", "NISSAN",
    "VOLKSWAGEN", "AUDI", "SUBARU", "CHEVROLET", "FORD", "MAZDA", "MITSUBISHI",
    "RENAULT", "SKODA", "LADA", "PORSCHE", "LAND ROVER", "DAEWOO",
    "GEELY", "BYD", "ZEEKR", "LIXIANG", "LI", "CHANGAN", "CHERY", "HAVAL",
    "EXEED", "OMODA", "JETOUR", "TANK", "HONGQI", "FAW", "JAC", "VOYAH", "DONGFENG"
]

CAR_MODELS = [
    "ALPHARD", "CAMRY", "COROLLA", "RAV4", "LAND CRUISER", "PRADO", "HIGHLANDER", "PRIUS",
    "FIT", "ACCORD", "CR-V", "CIVIC", "ODYSSEY", "STEPWGN", "SONATA", "ELANTRA", "SANTA FE",
    "TUCSON", "RIO", "SPORTAGE", "SORENTO", "MORNING", "SPRINTER", "FOCUS", "TRANSIT", "GOLF",
    "PASSAT", "JETTA", "TIGUAN", "POLO", "FORESTER", "OUTBACK", "IMPREZA", "LEGACY", "CRUZE",
    "MALIBU", "COBALT", "NEXIA", "MATIZ", "SPARK", "ESTIMA", "WISH", "NOAH", "VOXY", "HARRIER",
    "AVANTE", "PALISADE", "MONJARO", "COOLRAY", "TUGELLA", "OKAVANGO", "EMGRAND", "SONG", "HAN",
    "TANG", "YUAN", "CHAZOR", "SEAGULL", "001", "009", "TIGGO", "ARRIZO", "JOLION", "DARGO",
    "UNI-K", "UNI-T", "UNI-V", "TXL", "DASHING", "FREE", "DREAM", "K5"
]

KNOWN_COLORS_RU = ["БЕЛЫЙ", "СЕРЕБРИСТЫЙ", "СЕРЫЙ", "ЧЁРНЫЙ", "ЧЕРНЫЙ", "СИНИЙ",
                    "КРАСНЫЙ", "ЗЕЛЕНЫЙ", "ЖЁЛТЫЙ", "ЖЕЛТЫЙ", "ОРАНЖЕВЫЙ", "КОРИЧНЕВЫЙ",
                    "БОРДОВЫЙ", "ГОЛУБОЙ", "БЕЖЕВЫЙ", "ЗОЛОТИСТЫЙ", "ФИОЛЕТОВЫЙ"]
KNOWN_BODIES_RU = ["СЕДАН", "ХЭТЧБЕК", "УНИВЕРСАЛ", "ВНЕДОРОЖНИК", "КРОССОВЕР",
                    "КУПЕ", "МИНИВЭН", "ПИКАП", "ФУРГОН", "ЛЕГКОВОЙ", "МИКРОАВТОБУС"]

# =========================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =========================================================

def cyr_to_lat(text: str) -> str:
    """Транслитерация кириллицы → латиница для поиска марок/моделей."""
    mapping = {
        'А':'A','В':'B','С':'C','Е':'E','Н':'H','К':'K','М':'M',
        'Р':'P','Т':'T','Х':'X','У':'Y','И':'I','Л':'L','Д':'D',
        'Ф':'F','Г':'G','З':'Z','Б':'B','П':'P','Ь':'','Ъ':'',
        'Э':'E','Ч':'CH','Я':'YA','Ж':'J','О':'O'
    }
    return "".join(mapping.get(c, c) for c in text.upper())


def enhance_image_for_ocr(uploaded_file):
    """Улучшение изображения для OCR."""
    try:
        img = Image.open(uploaded_file).convert('RGB')
        img_array = np.array(img)
        gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
        # Увеличение масштаба для лучшего OCR
        gray = cv2.resize(gray, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC)
        # Нормализация
        norm_img = np.zeros_like(gray)
        gray = cv2.normalize(gray, norm_img, 0, 255, cv2.NORM_MINMAX)
        # Адаптивная бинаризация — лучше для физических документов с неравномерным освещением
        gray = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                      cv2.THRESH_BINARY, 31, 10)
        return Image.fromarray(gray)
    except Exception:
        return Image.open(uploaded_file)


def prepare_image_for_word(doc, uploaded_file, width_mm=70):
    """Подготовка фото для вставки в Word."""
    try:
        img = Image.open(uploaded_file)
        if img.mode != 'RGB':
            img = img.convert('RGB')
        temp_buffer = io.BytesIO()
        img.save(temp_buffer, format='JPEG')
        temp_buffer.seek(0)
        return InlineImage(doc, temp_buffer, width=Mm(width_mm))
    except Exception as e:
        st.error(f"Ошибка при обработке фото: {e}")
        return ""


def clean_field(text: str) -> str:
    """Убирает мусорные символы OCR из значения поля."""
    text = re.sub(r'[|}{\\@#%^&*_+=\[\]~`<>]', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

# =========================================================
# ГЛАВНЫЙ ПАРСЕР: ДВУХРЕЖИМНЫЙ
# =========================================================

def parse_documents(photos) -> dict:
    """
    Парсит один или несколько фото документов (физических или электронных).
    Возвращает словарь с найденными данными.
    """
    result = {k: "" for k in keys}
    result['auto_steering'] = "Левый руль"

    all_raw_texts = []

    for photo in photos:
        # Три прохода с разными настройками — берём лучший результат
        photo.seek(0)
        raw_img = Image.open(photo)
        photo.seek(0)
        enh_img = enhance_image_for_ocr(photo)

        text_raw  = pytesseract.image_to_string(raw_img, lang='rus+eng',
                                                 config='--psm 6 --oem 3')
        text_enh  = pytesseract.image_to_string(enh_img, lang='rus+eng',
                                                 config='--psm 6 --oem 3')
        # Режим для плотных горизонтальных таблиц (физический документ)
        text_psm4 = pytesseract.image_to_string(enh_img, lang='rus+eng',
                                                 config='--psm 4 --oem 3')

        # Выбираем наиболее полный текст для данного фото
        best = max([text_raw, text_enh, text_psm4], key=len)
        all_raw_texts.append(text_raw + "\n" + text_enh + "\n" + text_psm4)

    full_text = "\n".join(all_raw_texts)
    lines = [l.strip() for l in full_text.split('\n') if l.strip()]

    # ── 1. ОПРЕДЕЛЯЕМ ТИП ДОКУМЕНТА ───────────────────────────────────────────
    # Электронный СТС из Tunduk: данные и заголовки чередуются строками
    # Физический СТС: данные написаны СПРАВА от заголовка в одной строке
    is_electronic = any("Ээси / Владелец" in l or "Владелец / Owner" in l or
                         "Каттоо номери" in l.lower() for l in lines)
    is_physical   = any("МАРКАСЫ, МОДЕЛИ" in l.upper() or
                         "МАРКА, МОДЕЛЬ / MAKE" in l.upper() for l in lines)

    if is_electronic:
        result = _parse_electronic(lines, result, full_text)
    elif is_physical:
        result = _parse_physical(lines, result, full_text)
    else:
        # Универсальный режим — пробуем оба
        result = _parse_electronic(lines, result, full_text)
        result = _parse_physical(lines, result, full_text)

    # ── 2. ГЛОБАЛЬНЫЙ ПОИСК (работает для обоих типов) ────────────────────────
    clean_num = full_text.upper().replace(" ", "").replace("\n", "").replace("O", "0")
    clean_up  = full_text.upper()

    # Госномер КГ
    if not result['auto_reg']:
        m = re.search(r'\b\d{2}KG\d{3}[A-Z]{2,3}\b', clean_num)
        if m:
            result['auto_reg'] = m.group(0)

    # VIN (17 символов)
    if not result['auto_vin']:
        for m in re.finditer(r'\b[A-HJ-NPR-Z0-9]{17}\b', clean_up.replace("O", "0")):
            cand = m.group(0)
            if sum(c.isdigit() for c in cand) >= 4:
                if not re.search(r'(KATT|YEAR|KYRG|PERS|NUMB|MAKE|MOD|COL|VOL|ENG|OWN)', cand):
                    result['auto_vin'] = cand
                    break

    # Шасси (например MNH15-0044258)
    if not result['auto_vin']:
        for m in re.finditer(r'\b[A-Z]{2,5}\d{1,4}[-]?\d{4,8}\b', clean_up.replace("O", "0")):
            cand = m.group(0)
            if not cand.startswith("KG") and not re.search(r'(KATT|YEAR|KYRG)', cand):
                result['auto_vin'] = cand
                break

    # Марка + модель (глобально по всему тексту)
    if not result['auto_model']:
        lat = cyr_to_lat(clean_up)
        found_brand, found_model = "", ""
        for b in CAR_BRANDS:
            if re.search(rf'\b{re.escape(b)}\b', lat):
                found_brand = b
                break
        for m_name in CAR_MODELS:
            if re.search(rf'\b{re.escape(m_name)}\b', lat):
                found_model = m_name
                break
        if found_brand or found_model:
            result['auto_model'] = f"{found_brand} {found_model}".strip()

    # Положение руля
    if "ПРАВ" in clean_up and ("РУЛЬ" in clean_up or "WHEEL" in clean_up):
        result['auto_steering'] = "Правый руль"
    elif "ЛЕВ" in clean_up and ("РУЛЬ" in clean_up or "WHEEL" in clean_up):
        result['auto_steering'] = "Левый руль"

    return result


def _parse_physical(lines: list, result: dict, full_text: str) -> dict:
    """
    Парсер для ФИЗИЧЕСКОГО документа (пластиковая карточка СТС).
    Структура: «ЗАГОЛОВОК  ЗНАЧЕНИЕ» в одной строке.
    Заголовок — слева (на кыргызском/русском), значение — справа (обычно крупнее/жирнее).
    """

    def right_of(line: str, *keywords) -> str:
        """Берёт текст правее ключевого слова в строке."""
        ul = line.upper()
        for kw in keywords:
            if kw in ul:
                idx = ul.index(kw) + len(kw)
                return clean_field(line[idx:])
        return ""

    # Физический документ: данные справа от ключевого слова
    ANCHORS = {
        'auto_model':        ["МАРКА, МОДЕЛЬ", "MAKE, MODEL", "МАРКАСЫ, МОДЕЛИ"],
        'auto_year':         ["ГОД ВЫПУСКА", "YEAR OF MANUFACTURE", "ЧЫГАРЫЛГАН ЖЫЛЫ"],
        'auto_color':        ["ЦВЕТ", "COLOUR", "ТҮСҮ", "TYCY"],
        'auto_vin':          ["ИДЕНТИФИКАЦИОННЫЙ НОМЕР (VIN)", "VEHICLE IDENTIFICATION NUMBER",
                              "ИДЕНТИФИКАЦИЯЛЫК НОМЕРИ"],
        'auto_body_type':    ["ВИД КУЗОВА", "CAR BODY TYPE", "КУЗОВУНУН ТYPY", "КУЗОВУНУН ТУРУ"],
        'auto_engine':       ["РАБОЧИЙ ОБЪЕМ ДВИГАТЕЛЯ", "ENGINE CAPACITY", "КЫЙ МЫЛДАТКЫЧЫНЫН ИШ",
                              "КЫЙМЫЛДАТКЫЧЫНЫН ИШ КӨЛӨМҮ", "КЫЙМЫЛДАТКЫЧЫНЫН ИШ"],
    }

    for i, line in enumerate(lines):
        lu = line.upper()

        for field, kws in ANCHORS.items():
            if result[field]:
                continue
            for kw in kws:
                if kw in lu:
                    val = right_of(line, kw)
                    if not val and i + 1 < len(lines):
                        # значение может быть на следующей строке
                        val = clean_field(lines[i + 1])
                    if val and len(val) > 1:
                        # Для года — только 4-значное число
                        if field == 'auto_year':
                            ym = re.search(r'\b(199\d|200\d|201\d|202\d)\b', val)
                            result[field] = ym.group(0) if ym else ""
                        # Для объёма — только число
                        elif field == 'auto_engine':
                            vm = re.search(r'\b([8-9]\d{2}|[1-9]\d{3})\b', val)
                            if vm:
                                result[field] = vm.group(0)
                            elif i + 1 < len(lines):
                                vm2 = re.search(r'\b([8-9]\d{2}|[1-9]\d{3})\b', lines[i+1])
                                result[field] = vm2.group(0) if vm2 else ""
                        else:
                            result[field] = val.title() if field in ('auto_color', 'auto_body_type') else val
                    break

        # Положение руля в физическом документе (в строке с категорией)
        if "КАТЕГОРИЯ" in lu or "CATEGORY" in lu or "ОРУНДАРДЫН" in lu:
            if "ПРАВ" in lu:
                result['auto_steering'] = "Правый руль"
            elif "ЛЕВ" in lu:
                result['auto_steering'] = "Левый руль"

    # Тех.паспорт — в физическом документе НЕТ, он на обратной стороне карточки СТС
    # Ищем «AB XXXXXXX» глобально, игнорируя строки с оборудованием
    if not result['auto_tech_passport']:
        for line in lines:
            lu = line.upper()
            if any(bad in lu for bad in ["ОБОРУДОВАНИЕ", "ГАЗОБАЛЛОННОЕ", "ВЗАМЕН",
                                          "ОТМЕТКИ", "SPECIAL", "ГБО", "GAZI"]):
                continue
            m = re.search(r'\b([A-ZА-Я]{2})\s?(\d{6,7})\b', lu)
            if m:
                prefix = m.group(1)
                if prefix not in ("ОТ", "ДО", "ПИН", "KG", "VI"):
                    result['auto_tech_passport'] = f"{prefix} {m.group(2)}"
                    break

    return result


def _parse_electronic(lines: list, result: dict, full_text: str) -> dict:
    """
    Парсер для ЭЛЕКТРОННОГО документа (скриншот из приложения Tunduk).
    Структура: заголовок — одна строка, значение — следующая строка.
    """

    # Карта: поле → список ключевых слов-заголовков
    ANCHORS = {
        'auto_fio':          ["ЭЭСИ", "ВЛАДЕЛЕЦ", "OWNER", "СОБСТВЕННИК"],
        'auto_address':      ["ЭЭСИНИН ДАРЕГИ", "АДРЕС СОБСТВЕННИКА", "OWNER'S ADDRESS",
                              "ДАРЕГИ", "АДРЕС"],
        'auto_reg':          ["КАТТОО НОМЕРИ", "РЕГИСТРАЦИОННЫЙ НОМЕР", "REGISTRATION NUMBER",
                              "КАТТОО НОМЕРИ"],
        'auto_vin':          ["ИДЕНТИФИКАЦИЯЛЫК НОМЕРИ", "ИДЕНТИФИКАЦИОННЫЙ НОМЕР",
                              "IDENTIFICATION NUMBER", "КУЗОВУН №", "ШАССИСИНИН №",
                              "CAR BODY", "CHASSIS"],
        'auto_year':         ["ГОД ВЫПУСКА", "YEAR OF MANUFACTURE", "ЧЫГАРЫЛГАН ЖЫЛЫ"],
        'auto_model':        ["МАРКАСЫ", "МАРКА", "МОДЕЛИ", "MAKE, MODEL",
                              "VEHICLE BRAND", "BRAND", "MODEL"],
        'auto_color':        ["ТҮСҮ", "ЦВЕТ", "COLOUR", "COLOR"],
        'auto_body_type':    ["ВИД КУЗОВА", "КУЗОВТУН ТУРУ", "CAR BODY TYPE", "BODY TYPE",
                              "VEHICLE TYPE", "ТК ТУРУ", "ТИП ТС"],
        'auto_engine':       ["ИШ КӨЛӨМҮ", "ОБЪЕМ ДВИГАТЕЛЯ", "ENGINE CAPACITY",
                              "WORKING VOLUME", "РАБОЧИЙ ОБЪЕМ"],
        'auto_tech_passport': ["СЕРИЯ И НОМЕР", "SERIES AND NUMBER", "СЕРИЯСЫ ЖАНА НОМЕРИ",
                               "СЕРИЯСЫ"],
    }

    def get_value_after_anchor(i: int, kw_found: str) -> str:
        """Берёт значение: сначала справа от якоря, потом — следующая строка."""
        line = lines[i]
        # Справа от якоря в той же строке
        parts = re.split(re.escape(kw_found), line.upper(), maxsplit=1)
        right = clean_field(line[len(parts[0]) + len(kw_found):]) if len(parts) > 1 else ""

        if right and len(right) > 1:
            return right

        # Следующие строки (до 3х), пропуская слишком короткие
        for j in range(1, 4):
            if i + j >= len(lines):
                break
            cand = clean_field(lines[i + j])
            cand_up = cand.upper()
            # Пропускаем строки-заголовки и переводы
            if any(bad in cand_up for bad in ["/ OWNER", "/ АДРЕС", "REGISTRATION", "PERSONAL",
                                               "VEHICLE ID", "MANUFACTURE", "CAPACITY",
                                               "IDENTIFICATION", "ISSUING", "SERIES"]):
                continue
            if len(cand) > 2:
                return cand
        return ""

    for i, line in enumerate(lines):
        lu = line.upper()

        for field, kws in ANCHORS.items():
            if result[field]:
                continue
            matched_kw = None
            for kw in kws:
                if kw in lu:
                    matched_kw = kw
                    break
            if not matched_kw:
                continue

            val = get_value_after_anchor(i, matched_kw)
            if not val:
                continue

            val_up = val.upper()

            # ── Пост-обработка по типу поля ──────────────────────────────────
            if field == 'auto_fio':
                # ФИО: только кириллица, ≥2 слова, без цифр, без "РЕСПУБЛИКА"
                cyr = re.sub(r'[^А-ЯЁа-яё\s\-]', '', val).strip()
                if (len(cyr.split()) >= 2 and not re.search(r'\d', cyr)
                        and "РЕСПУБЛИКА" not in cyr.upper()
                        and "КЫРГЫЗ" not in cyr.upper()):
                    result[field] = cyr.title()

            elif field == 'auto_address':
                # Адрес: не должен быть числом, пин-кодом или фио
                if (len(val) > 10
                        and not re.fullmatch(r'[\d\s]+', val)
                        and "БЕРГЕН" not in val_up
                        and "ОРГАН" not in val_up):
                    result[field] = val

            elif field == 'auto_reg':
                m = re.search(r'\b\d{2}KG\d{3}[A-Z]{2,3}\b', val.replace(" ", "").upper())
                if m:
                    result[field] = m.group(0)

            elif field == 'auto_vin':
                # Шасси / VIN
                m17 = re.search(r'\b[A-HJ-NPR-Z0-9]{17}\b', val.replace("O", "0"))
                if m17:
                    result[field] = m17.group(0)
                else:
                    m_chassis = re.search(r'\b[A-Z]{2,5}\d{1,4}[-]?\d{4,8}\b', val)
                    if m_chassis:
                        result[field] = m_chassis.group(0)

            elif field == 'auto_year':
                ym = re.search(r'\b(199\d|200\d|201\d|202\d)\b', val)
                if ym:
                    result[field] = ym.group(0)

            elif field == 'auto_model':
                # Марка может прийти отдельно от модели — склеиваем
                lat = cyr_to_lat(val_up)
                brand, model_name = "", ""
                for b in CAR_BRANDS:
                    if re.search(rf'\b{re.escape(b)}\b', lat):
                        brand = b
                        break
                for mn in CAR_MODELS:
                    if re.search(rf'\b{re.escape(mn)}\b', lat):
                        model_name = mn
                        break
                # Если следующая строка — модель (а эта — марка)
                if brand and not model_name and i + 1 < len(lines):
                    next_lat = cyr_to_lat(lines[i+1].upper())
                    for mn in CAR_MODELS:
                        if re.search(rf'\b{re.escape(mn)}\b', next_lat):
                            model_name = mn
                            break
                if brand or model_name:
                    result[field] = f"{brand} {model_name}".strip()
                elif val and len(val) > 1:
                    result[field] = val

            elif field == 'auto_color':
                for color in KNOWN_COLORS_RU:
                    if color in val_up:
                        result[field] = color.capitalize()
                        break
                if not result[field] and len(val) > 2:
                    result[field] = val.capitalize()

            elif field == 'auto_body_type':
                for bt in KNOWN_BODIES_RU:
                    if bt in val_up:
                        result[field] = val.title()
                        break
                if not result[field] and len(val) > 2:
                    result[field] = val.title()

            elif field == 'auto_engine':
                vm = re.search(r'\b([8-9]\d{2}|[1-9]\d{3})\b', val)
                if vm:
                    result[field] = vm.group(0)

            elif field == 'auto_tech_passport':
                m = re.search(r'\b([A-ZА-Я]{2})\s?(\d{6,7})\b', val.upper())
                if m and m.group(1) not in ("ОТ", "ДО", "KG", "VI", "ПИН"):
                    result[field] = f"{m.group(1)} {m.group(2)}"

        # Положение руля
        if any(kw in lu for kw in ["РАСПОЛОЖЕНИЕ РУЛЯ", "STEERING WHEEL", "ПОЛОЖЕНИЕ РУЛЯ",
                                     "РУЛЬ", "WHEEL POSITION"]):
            val = get_value_after_anchor(i, [kw for kw in
                                             ["РАСПОЛОЖЕНИЕ РУЛЯ","STEERING","РУЛЬ"] if kw in lu][0])
            if "ПРАВ" in val.upper() or "RIGHT" in val.upper():
                result['auto_steering'] = "Правый руль"
            elif "ЛЕВ" in val.upper() or "LEFT" in val.upper():
                result['auto_steering'] = "Левый руль"

    # Тех.паспорт — fallback глобальный
    if not result['auto_tech_passport']:
        for line in lines:
            lu = line.upper()
            if any(bad in lu for bad in ["ОБОРУДОВАНИЕ", "ГАЗОБАЛЛОННОЕ", "ВЗАМЕН",
                                          "ОТМЕТКИ", "SPECIAL", "ГБО"]):
                continue
            m = re.search(r'\b([A-ZА-Я]{2})\s?(\d{6,7})\b', lu)
            if m and m.group(1) not in ("ОТ", "ДО", "KG", "VI", "ПИН", "НА", "ВИ"):
                result['auto_tech_passport'] = f"{m.group(1)} {m.group(2)}"
                break

    return result


# =========================================================
# ИНТЕРФЕЙС
# =========================================================

st.title("🚗 Рабочее место оценщика — Гарант Оценка")

# Шаблон
if os.path.exists(TEMPLATE_NAME):
    st.success(f"✅ Шаблон `{TEMPLATE_NAME}` подключён автоматически.")
    template_source = TEMPLATE_NAME
else:
    st.warning(f"⚠️ Файл `{TEMPLATE_NAME}` не найден. Загрузите его вручную:")
    template_source = st.file_uploader(f"Загрузите шаблон отчёта ({TEMPLATE_NAME})", type="docx")

st.divider()

# ── Блок сканера ───────────────────────────────────────────────────────────────
st.header("📸 1. Автозаполнение по фото документов")
st.caption("Поддерживаются фото физических документов (пластиковая карточка СТС) "
           "и скриншоты из приложения (Tunduk/Goskey). Можно загрузить сразу оба.")

sts_photos = st.file_uploader(
    "Загрузите фото СТС / техпаспорта (можно выбрать несколько)",
    accept_multiple_files=True,
    type=["jpg", "jpeg", "png", "bmp", "webp"]
)

if sts_photos:
    if st.button("🔍 Распознать документы", type="primary"):
        for k in keys:
            st.session_state[k] = ""
        st.session_state.auto_steering = "Левый руль"

        with st.spinner("Анализируем документы..."):
            try:
                parsed = parse_documents(sts_photos)

                # Переносим в session_state
                for k in keys:
                    if parsed.get(k):
                        st.session_state[k] = parsed[k]
                st.session_state.auto_steering = parsed.get('auto_steering', "Левый руль")

                # Показываем что нашли / не нашли
                found_fields = {k: v for k, v in parsed.items() if v and k != 'auto_steering'}
                empty_fields = [k for k in keys if not parsed.get(k)]

                st.success(f"✅ Распознано полей: {len(found_fields)} из {len(keys)}")

                if empty_fields:
                    labels = {
                        'auto_vin': 'VIN/Шасси', 'auto_reg': 'Гос. номер',
                        'auto_year': 'Год выпуска', 'auto_model': 'Марка/Модель',
                        'auto_fio': 'ФИО', 'auto_engine': 'Объём ДВС',
                        'auto_color': 'Цвет', 'auto_address': 'Адрес',
                        'auto_tech_passport': 'Тех. паспорт', 'auto_body_type': 'Тип кузова'
                    }
                    missing = ", ".join(labels.get(f, f) for f in empty_fields)
                    st.warning(f"⚠️ Не удалось распознать: {missing} — заполните вручную.")

                with st.expander("📄 Показать весь распознанный текст"):
                    all_texts = []
                    for photo in sts_photos:
                        photo.seek(0)
                        t = pytesseract.image_to_string(Image.open(photo), lang='rus+eng')
                        all_texts.append(t)
                    st.text("\n\n---\n\n".join(all_texts))

            except Exception as e:
                st.error(f"Ошибка распознавания: {e}")

st.divider()

# ── Форма ввода данных ─────────────────────────────────────────────────────────
st.header("📝 2. Данные для отчёта")
col1, col2 = st.columns(2)

with col1:
    st.subheader("Общие данные")
    report_num   = st.text_input("Номер отчёта:")
    contract_num = st.text_input("Номер договора:")
    date         = st.text_input("Дата оценки:")
    customer     = st.text_input("ФИО Заказчика:", key="auto_fio")
    address      = st.text_input("Адрес регистрации:", key="auto_address")
    sum_num      = st.text_input("Сумма ущерба (цифрами):")
    sum_words    = st.text_input("Сумма ущерба (прописью):")

with col2:
    st.subheader("Данные автомобиля")
    car_model    = st.text_input("Марка, модель:", key="auto_model")
    reg_num      = st.text_input("Гос. номер:", key="auto_reg")
    vin          = st.text_input("VIN / Шасси №:", key="auto_vin")
    year         = st.text_input("Год выпуска:", key="auto_year")
    tech_passport = st.text_input("Тех. паспорт №:", key="auto_tech_passport")
    engine_vol   = st.text_input("Объём ДВС (куб.см):", key="auto_engine")
    color        = st.text_input("Цвет кузова:", key="auto_color")

    col_inner1, col_inner2 = st.columns(2)
    with col_inner1:
        body_type = st.text_input("Тип кузова:", key="auto_body_type")
    with col_inner2:
        steering = st.selectbox("Положение руля:", ["Левый руль", "Правый руль"],
                                 key="auto_steering")

st.subheader("Описание повреждений")
damage_desc = st.text_area("При осмотре установлено (повреждения):", height=80)
repair_desc = st.text_area("Для восстановления требуется:", height=80)

st.divider()

# ── Фотографии повреждений ─────────────────────────────────────────────────────
st.header("🖼️ 3. Фотографии автомобиля для отчёта")
uploaded_photos = st.file_uploader(
    "Загрузите фотографии повреждений",
    accept_multiple_files=True,
    type=["jpg", "jpeg", "png"]
)

photo_data = []
if uploaded_photos:
    st.markdown("**Подпишите загруженные фотографии:**")
    col_l1, col_l2 = st.columns(2)
    defaults = ["Вид спереди", "Вид слева", "Вид сзади", "Вид справа"]
    for i, photo in enumerate(uploaded_photos):
        with col_l1 if i % 2 == 0 else col_l2:
            default_label = defaults[i] if i < len(defaults) else f"Дефект {i-1}"
            label = st.text_input(f"Подпись для '{photo.name}':",
                                   value=default_label, key=f"photo_lab_{i}")
            photo_data.append({"file": photo, "label": label})

st.divider()

# ── Генерация DOCX ─────────────────────────────────────────────────────────────
if template_source is not None:
    if st.button("📄 СГЕНЕРИРОВАТЬ ОТЧЁТ", type="primary", use_container_width=True):
        try:
            doc = DocxTemplate(template_source)

            photos_rows = []
            for i in range(0, len(photo_data), 2):
                item1 = photo_data[i]
                item1["file"].seek(0)
                img1 = prepare_image_for_word(doc, item1["file"])
                lab1 = item1["label"]

                img2, lab2 = "", ""
                if i + 1 < len(photo_data):
                    item2 = photo_data[i + 1]
                    item2["file"].seek(0)
                    img2 = prepare_image_for_word(doc, item2["file"])
                    lab2 = item2["label"]

                photos_rows.append({"img1": img1, "lab1": lab1, "img2": img2, "lab2": lab2})

            context = {
                "REPORT_NUM":      report_num,
                "CONTRACT_NUM":    contract_num,
                "DATE":            date,
                "CUSTOMER_NAME":   customer,
                "ADDRESS":         address,
                "CAR_MODEL":       car_model,
                "REG_NUM":         reg_num,
                "VIN":             vin,
                "TECH_PASSPORT":   tech_passport,
                "YEAR":            year,
                "ENGINE_VOL":      engine_vol,
                "COLOR":           color,
                "BODY_TYPE":       body_type,
                "STEERING":        steering,
                "TOTAL_SUM_NUM":   sum_num,
                "TOTAL_SUM_WORDS": sum_words,
                "DAMAGE_DESC":     damage_desc,
                "REPAIR_DESC":     repair_desc,
                "photos":          photos_rows,
            }

            doc.render(context)
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            st.success("🎉 Отчёт успешно сгенерирован!")
            file_name = f"Отчет_{customer or 'Новый_клиент'}_{reg_num or ''}.docx"
            st.download_button(
                label="📥 СКАЧАТЬ ГОТОВЫЙ ОТЧЁТ",
                data=buffer,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"Ошибка при генерации файла: {e}")
