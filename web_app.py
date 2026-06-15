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
# НАСТРОЙКИ СИСТЕМЫ И ПАМЯТЬ
# =========================================================

# ПРИ ПУБЛИКАЦИИ НА GITHUB ЭТА СТРОКА ДОЛЖНА БЫТЬ ЗАКОММЕНТИРОВАНА!
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

TEMPLATE_NAME = "образец отчета.docx"

st.set_page_config(page_title="Генератор Отчетов - Гарант Оценка", layout="wide")

keys = ['auto_vin', 'auto_reg', 'auto_year', 'auto_model', 'auto_fio', 
        'auto_engine', 'auto_color', 'auto_address', 'auto_tech_passport', 'auto_body_type']
for k in keys:
    if k not in st.session_state:
        st.session_state[k] = ""

if 'auto_steering' not in st.session_state:
    st.session_state.auto_steering = "Левый руль"

# =========================================================
# ФУНКЦИИ ОБРАБОТКИ ФОТОГРАФИЙ
# =========================================================

def prepare_image_for_word(doc, uploaded_file, width_mm=70):
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

def enhance_image_for_ocr(uploaded_file):
    try:
        img = Image.open(uploaded_file).convert('RGB')
        img_array = np.array(img)
        gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
        gray = cv2.resize(gray, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_CUBIC)
        norm_img = np.zeros((gray.shape[0], gray.shape[1]))
        gray = cv2.normalize(gray, norm_img, 0, 255, cv2.NORM_MINMAX)
        blur = cv2.GaussianBlur(gray, (3, 3), 0)
        return Image.fromarray(blur)
    except Exception:
        return Image.open(uploaded_file)

# =========================================================
# БАЗА ДАННЫХ АВТО 
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

KNOWN_COLORS = ["БЕЛЫЙ", "СЕРЕБРИСТЫЙ", "СЕРЫЙ", "ЧЕРНЫЙ", "СИНИЙ", "КРАСНЫЙ", "ЗЕЛЕНЫЙ", "ЖЕЛТЫЙ", "ОРАНЖЕВЫЙ", "КОРИЧНЕВЫЙ", "БОРДОВЫЙ", "ГОЛУБОЙ", "БЕЖЕВЫЙ", "ЗОЛОТИСТЫЙ", "ФИОЛЕТОВЫЙ"]
KNOWN_BODIES = ["СЕДАН", "ХЭТЧБЕК", "УНИВЕРСАЛ", "ВНЕДОРОЖНИК", "КРОССОВЕР", "КУПЕ", "МИНИВЭН", "ПИКАП", "ФУРГОН", "ЛЕГКОВОЙ"]

# =========================================================
# ИНТЕРФЕЙС ПРИЛОЖЕНИЯ
# =========================================================

st.title("🚗 Рабочее место оценщика")

if os.path.exists(TEMPLATE_NAME):
    st.success(f"✅ Базовый шаблон отчета (`{TEMPLATE_NAME}`) успешно подключен автоматически.")
    template_source = TEMPLATE_NAME
else:
    st.warning(f"⚠️ Файл `{TEMPLATE_NAME}` не найден. Загрузите его вручную:")
    template_source = st.file_uploader(f"Загрузите шаблон отчета ({TEMPLATE_NAME})", type="docx")

st.divider()

# --- БЛОК СКАНЕРА СТС ---
st.header("📸 1. Автозаполнение по фото")

sts_photos = st.file_uploader(
    "Загрузите фото документов (СТС, техпаспорт) - можно выбрать сразу несколько", 
    accept_multiple_files=True
)

if sts_photos:
    if st.button("Распознать документы", type="primary"):
        for k in keys: st.session_state[k] = ""
        st.session_state.auto_steering = "Левый руль"
        
        with st.spinner("Снайперский двунаправленный алгоритм читает документы..."):
            try:
                all_extracted_text = ""
                for photo in sts_photos:
                    raw_img = Image.open(photo)
                    text_raw = pytesseract.image_to_string(raw_img, lang='rus+eng')
                    
                    enhanced_img = enhance_image_for_ocr(photo)
                    text_enh = pytesseract.image_to_string(enhanced_img, lang='rus+eng')
                    
                    all_extracted_text += text_raw + "\n\n" + text_enh + "\n\n"
                
                lines = [line.strip() for line in all_extracted_text.split('\n') if line.strip()]
                clean_text_numbers = all_extracted_text.upper().replace(" ", "").replace("\n", "").replace("O", "0")
                clean_text_words = all_extracted_text.upper()

                # ==============================================================
                # 1. ГЛОБАЛЬНЫЙ ПОИСК (Госномер, Марка, Модель, VIN)
                # ==============================================================

                reg_match_kg = re.search(r'\d{2}KG\d{3}[A-Z]{2,3}', clean_text_numbers)
                if reg_match_kg: st.session_state.auto_reg = reg_match_kg.group(0)

                # VIN
                vin_matches = re.finditer(r'\b[A-HJ-NPR-Z0-9]{17}\b', clean_text_words.replace("O", "0"))
                for match in vin_matches:
                    vin_cand = match.group(0)
                    if sum(c.isdigit() for c in vin_cand) >= 4 and not re.search(r'(KATT|YEAR|KYRG|PERS|NUMB|MAKE|MOD|COL|VOL|ENG|OWN)', vin_cand):
                        st.session_state.auto_vin = vin_cand
                        break
                
                # Номер Шасси
                if not st.session_state.auto_vin:
                    chassis_matches = re.finditer(r'\b[A-Z]{2,5}\d{1,4}[-]?\d{4,8}\b', clean_text_words.replace("O", "0"))
                    for match in chassis_matches:
                        chass = match.group(0)
                        if not chass.startswith("KG") and not re.search(r'(KATT|YEAR|KYRG)', chass):
                            st.session_state.auto_vin = chass
                            break

                # Марка и Модель
                mapping = {
                    'А': 'A', 'В': 'B', 'С': 'C', 'Е': 'E', 'Н': 'H', 'К': 'K', 'М': 'M', 
                    'Р': 'P', 'Т': 'T', 'Х': 'X', 'У': 'Y', 'И': 'I', 'Л': 'L', 'Д': 'D', 'Ф': 'F', 
                    'Г': 'G', 'З': 'Z', 'Б': 'B', 'П': 'P', 'Ь': '', 'Ъ': '', 'Э': 'E', 'Ч': 'CH', 
                    'Я': 'YA', 'Ж': 'J', 'О': 'O'
                }
                lat_global_text = "".join(mapping.get(c, c) for c in clean_text_words)
                
                found_brand, found_model = "", ""
                for b in CAR_BRANDS:
                    if re.search(rf'\b{b}\b', lat_global_text): found_brand = b; break
                for m in CAR_MODELS:
                    if re.search(rf'\b{m}\b', lat_global_text): found_model = m; break
                
                if found_brand or found_model:
                    st.session_state.auto_model = f"{found_brand} {found_model}".strip()

                if "ПРАВ" in clean_text_words: st.session_state.auto_steering = "Правый руль"
                elif "ЛЕВ" in clean_text_words: st.session_state.auto_steering = "Левый руль"

                # ==============================================================
                # 2. УМНЫЙ ПОИСК "СПРАВА И СНИЗУ" (Якорный метод)
                # ==============================================================
                
                for i, line in enumerate(lines):
                    line_up = line.upper()

                    # --- ПОИСК ФИО ---
                    if not st.session_state.auto_fio and any(kw in line_up for kw in ["ВЛАДЕЛЕЦ", "ЭЭСИ", "OWNER"]):
                        right = re.split(r'ВЛАДЕЛЕЦ|ЭЭСИ|OWNER', line_up)[-1]
                        right_cyr = re.sub(r'[A-Z!@#\$%\^&\*\(\)_+=\{\}\[\];\'<>\?\|\\~`]', '', right).replace(":", "").strip()
                        if len(right_cyr.split()) >= 2 and not re.search(r'\d', right_cyr):
                            st.session_state.auto_fio = right_cyr.title()
                        else:
                            for j in range(1, 4):
                                if i + j < len(lines):
                                    cand = lines[i+j]
                                    cand_cyr = re.sub(r'[A-Z!@#\$%\^&\*\(\)_+=\{\}\[\];\'<>\?\|\\~`]', '', cand.upper()).replace(":", "").strip()
                                    if len(cand_cyr.split()) >= 2 and not re.search(r'\d', cand_cyr) and "АДРЕС" not in cand_cyr:
                                        st.session_state.auto_fio = cand_cyr.title()
                                        break

                    # --- ПОИСК АДРЕСА ---
                    if not st.session_state.auto_address and any(kw in line_up for kw in ["ДАРЕГИ", "АДРЕС", "ADDRESS"]):
                        right = re.split(r'ДАРЕГИ|АДРЕС СОБСТВЕННИКА|ADDRESS', line_up)[-1]
                        right_cyr = re.sub(r'[A-Z]', '', right).replace(":", "").replace("'", "").strip()
                        if len(right_cyr) > 10 and "БЕРГЕН" not in right_cyr:
                            st.session_state.auto_address = right_cyr
                        else:
                            addr_parts = []
                            for j in range(1, 5):
                                if i + j < len(lines):
                                    cand_up = lines[i+j].upper()
                                    if any(kw in cand_up for kw in ["БЕРГЕН", "ОРГАН", "МАРКА", "МОДЕЛЬ", "КАТТАЛГАН"]): break
                                    cand_cyr = re.sub(r'[A-Z]', '', cand_up).strip()
                                    if len(cand_cyr) > 3 and not re.fullmatch(r'[\d\s\.\,]+', cand_cyr):
                                        addr_parts.append(cand_cyr)
                            if addr_parts:
                                st.session_state.auto_address = " ".join(addr_parts).strip()

                    # --- ПОИСК ОБЪЕМА ДВС ---
                    if not st.session_state.auto_engine and any(kw in line_up for kw in ["ОБЪЕМ", "КӨЛӨМҮ", "CAPACITY"]):
                        right = re.split(r'ОБЪЕМ|КӨЛӨМҮ|CAPACITY|СМ3', line_up)[-1]
                        vols = re.findall(r'\b([8-9]\d{2}|[1-7]\d{3})\b', right)
                        if vols:
                            st.session_state.auto_engine = vols[-1]
                        else:
                            for j in range(1, 4):
                                if i + j < len(lines):
                                    vols_below = re.findall(r'\b([8-9]\d{2}|[1-7]\d{3})\b', lines[i+j])
                                    if vols_below:
                                        st.session_state.auto_engine = vols_below[-1]
                                        break

                    # --- ПОИСК ГОДА ---
                    if not st.session_state.auto_year and any(kw in line_up for kw in ["ГОД", "ЖЫЛЫ", "YEAR"]):
                        right = re.split(r'ГОД|ЖЫЛЫ|MANUFACTURE', line_up)[-1]
                        y_match = re.search(r'\b(199\d|200\d|201\d|202\d)\b', right)
                        if y_match:
                            st.session_state.auto_year = y_match.group(0)
                        else:
                            for j in range(1, 3):
                                if i + j < len(lines):
                                    ym = re.search(r'\b(199\d|200\d|201\d|202\d)\b', lines[i+j])
                                    if ym:
                                        st.session_state.auto_year = ym.group(0)
                                        break

                    # --- ЦВЕТ И ТИП КУЗОВА ---
                    if not st.session_state.auto_color:
                        for c in KNOWN_COLORS:
                            if c in line_up and "ЦВЕТ" not in line_up:
                                st.session_state.auto_color = line.title(); break
                    if not st.session_state.auto_body_type:
                        for bt in KNOWN_BODIES:
                            if bt in line_up:
                                st.session_state.auto_body_type = line.title(); break
                                
                    # --- ПОИСК ТЕХПАСПОРТА (ЯКОРЬ) ---
                    if not st.session_state.auto_tech_passport and any(kw in line_up for kw in ["СЕРИЯ", "SERIES"]):
                        right = re.split(r'НОМЕР|СЕРИЯСЫ|NUMBER', line_up)[-1]
                        tp_match = re.search(r'\b[A-ZА-Я]{2}\s?\d{6,7}\b', right)
                        if tp_match and not tp_match.group(0).startswith("ОТ") and not tp_match.group(0).startswith("ДО"):
                            st.session_state.auto_tech_passport = tp_match.group(0)
                        else:
                            for j in range(1, 4):
                                if i + j < len(lines):
                                    cand = lines[i+j].upper()
                                    tp_match_below = re.search(r'\b[A-ZА-Я]{2}\s?\d{6,7}\b', cand)
                                    if tp_match_below and not tp_match_below.group(0).startswith("ОТ") and not tp_match_below.group(0).startswith("ДО"):
                                        st.session_state.auto_tech_passport = tp_match_below.group(0)
                                        break

                # --- ФОЛБЭК ДЛЯ ТЕХПАСПОРТА (ГЛОБАЛЬНЫЙ ПОИСК С ЗАЩИТОЙ ОТ ГБО) ---
                if not st.session_state.auto_tech_passport:
                    for line in lines:
                        line_up = line.upper()
                        # Жестко игнорируем строки, в которых есть упоминания об оборудовании и заменах!
                        if any(bad in line_up for bad in ["ОБОРУДОВАНИЕ", "ГАЗОБАЛЛОННОЕ", "ВЗАМЕН", "ОТМЕТКИ", "SPECIAL"]):
                            continue
                        tp_match = re.search(r'\b[A-ZА-Я]{2}\s?\d{6,7}\b', line_up)
                        if tp_match and not tp_match.group(0).startswith("ОТ") and not tp_match.group(0).startswith("ДО"):
                            st.session_state.auto_tech_passport = tp_match.group(0)
                            break

                st.success("Документы проанализированы! Данные перенесены в форму ниже.")
                
                with st.expander("Показать весь распознанный текст (системный)"):
                    st.text(all_extracted_text)
                    
            except Exception as e:
                st.error(f"Ошибка распознавания: {e}")

st.divider()

# --- РУЧНОЙ ВВОД ДАННЫХ ---
st.header("📝 2. Данные для отчета")
col1, col2 = st.columns(2)

with col1:
    st.subheader("Общие данные")
    report_num = st.text_input("Номер отчета:")
    contract_num = st.text_input("Номер договора:")
    date = st.text_input("Дата оценки:")
    
    customer = st.text_input("ФИО Заказчика:", key="auto_fio")
    address = st.text_input("Адрес регистрации:", key="auto_address")
    
    sum_num = st.text_input("Сумма ущерба (цифрами):")
    sum_words = st.text_input("Сумма ущерба (прописью):")

with col2:
    st.subheader("Данные автомобиля")
    car_model = st.text_input("Марка, модель:", key="auto_model")
    reg_num = st.text_input("Гос. номер:", key="auto_reg")
    vin = st.text_input("VIN код (Шасси):", key="auto_vin")
    year = st.text_input("Год выпуска:", key="auto_year")
    
    tech_passport = st.text_input("Тех. паспорт №:", key="auto_tech_passport")
    engine_vol = st.text_input("Объем ДВС:", key="auto_engine")
    color = st.text_input("Цвет кузова:", key="auto_color")
    
    col_inner1, col_inner2 = st.columns(2)
    with col_inner1:
        body_type = st.text_input("Тип кузова:", key="auto_body_type")
    with col_inner2:
        steering = st.selectbox("Положение руля:", ["Левый руль", "Правый руль"], key="auto_steering")

st.subheader("Описание повреждений")
damage_desc = st.text_area("Характеристика повреждений (при осмотре установлено):", height=80)
repair_desc = st.text_area("Требуемый ремонт (для восстановления требуется):", height=80)

st.divider()

# --- БЛОК ФОТОГРАФИЙ ПОВРЕЖДЕНИЙ ---
st.header("🖼️ 3. Фотографии автомобиля")

uploaded_photos = st.file_uploader(
    "Загрузите фотографии для таблицы отчета", 
    accept_multiple_files=True
)

photo_data = []

if uploaded_photos:
    st.markdown("**Подпишите загруженные фотографии:**")
    col_labels1, col_labels2 = st.columns(2)
    
    for i, photo in enumerate(uploaded_photos):
        with col_labels1 if i % 2 == 0 else col_labels2:
            default_label = f"Вид спереди" if i == 0 else f"Вид слева" if i == 1 else f"Дефект {i-1}"
            label = st.text_input(f"Подпись для '{photo.name}':", value=default_label, key=f"photo_lab_{i}")
            photo_data.append({"file": photo, "label": label})

st.divider()

# =========================================================
# ЛОГИКА ГЕНЕРАЦИИ WORD ДОКУМЕНТА
# =========================================================

if template_source is not None:
    if st.button("СГЕНЕРИРОВАТЬ ОТЧЕТ", type="primary", use_container_width=True):
        try:
            doc = DocxTemplate(template_source)
            
            photos_rows = []
            for i in range(0, len(photo_data), 2):
                item1 = photo_data[i]
                img1 = prepare_image_for_word(doc, item1["file"])
                lab1 = item1["label"]
                
                img2 = ""
                lab2 = ""
                if i + 1 < len(photo_data):
                    item2 = photo_data[i+1]
                    img2 = prepare_image_for_word(doc, item2["file"])
                    lab2 = item2["label"]
                
                photos_rows.append({
                    "img1": img1, "lab1": lab1,
                    "img2": img2, "lab2": lab2
                })

            context = {
                "REPORT_NUM": report_num,
                "CONTRACT_NUM": contract_num,
                "DATE": date,
                "CUSTOMER_NAME": customer,
                "ADDRESS": address,
                "CAR_MODEL": car_model,
                "REG_NUM": reg_num,
                "VIN": vin,
                "TECH_PASSPORT": tech_passport,
                "YEAR": year,
                "ENGINE_VOL": engine_vol,
                "COLOR": color,
                "BODY_TYPE": body_type,
                "STEERING": steering,
                "TOTAL_SUM_NUM": sum_num,
                "TOTAL_SUM_WORDS": sum_words,
                "DAMAGE_DESC": damage_desc,
                "REPAIR_DESC": repair_desc,
                "photos": photos_rows
            }
            
            doc.render(context)
            
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.success("🎉 Отчет успешно сгенерирован и готов к скачиванию!")
            
            file_name = f"Отчет_Оценка_{customer if customer else 'Новый_клиент'}.docx"
            
            st.download_button(
                label="📥 СКАЧАТЬ ГОТОВЫЙ ОТЧЕТ",
                data=buffer,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"Произошла ошибка при генерации файла: {e}")
