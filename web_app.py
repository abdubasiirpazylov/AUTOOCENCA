import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import io
import os
import pytesseract
from PIL import Image
import re
import difflib
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
    """OpenCV фильтр: очищает документ от теней и бликов"""
    try:
        # Обязательно конвертируем в RGB, чтобы избежать ошибок с альфа-каналом (RGBA)
        img = Image.open(uploaded_file).convert('RGB')
        img_array = np.array(img)

        gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
        gray = cv2.resize(gray, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_CUBIC)

        norm_img = np.zeros((gray.shape[0], gray.shape[1]))
        gray = cv2.normalize(gray, norm_img, 0, 255, cv2.NORM_MINMAX)

        blur = cv2.GaussianBlur(gray, (3, 3), 0)

        return Image.fromarray(blur)
    except Exception as e:
        return Image.open(uploaded_file)

# =========================================================
# БАЗА ДАННЫХ АВТО (ОЧИЩЕНА ОТ 2-БУКВЕННЫХ ЛОЖНЫХ МОДЕЛЕЙ)
# =========================================================

CAR_BRANDS = [
    "TOYOTA", "HONDA", "HYUNDAI", "KIA", "LEXUS", "BMW", "MERCEDES-BENZ", "MERCEDES", "NISSAN", 
    "VOLKSWAGEN", "AUDI", "SUBARU", "CHEVROLET", "FORD", "MAZDA", "MITSUBISHI", 
    "RENAULT", "SKODA", "LADA", "PORSCHE", "LAND ROVER", "DAEWOO", 
    "GEELY", "BYD", "ZEEKR", "LIXIANG", "LI", "CHANGAN", "CHERY", "HAVAL", 
    "EXEED", "OMODA", "JETOUR", "TANK", "HONGQI", "FAW", "JAC", "VOYAH", "DONGFENG"
]

# Убраны: IS, X, NX, RX, LX, ES, GX, GS, L7, L8, L9 (Они искали английские слова и цифры в тексте)
CAR_MODELS = [
    "ALPHARD", "CAMRY", "COROLLA", "RAV4", "LAND CRUISER", "PRADO", "HIGHLANDER", "PRIUS", 
    "FIT", "ACCORD", "CR-V", "CIVIC", "ODYSSEY", "STEPWGN", "SONATA", "ELANTRA", "SANTA FE", 
    "TUCSON", "RIO", "SPORTAGE", "SORENTO", "MORNING", "SPRINTER", "FOCUS", "TRANSIT", "GOLF", 
    "PASSAT", "JETTA", "TIGUAN", "POLO", "FORESTER", "OUTBACK", "IMPREZA", "LEGACY", "CRUZE", 
    "MALIBU", "COBALT", "NEXIA", "MATIZ", "SPARK", "ESTIMA", "WISH", "NOAH", "VOXY", "HARRIER", 
    "AVANTE", "PALISADE", "MONJARO", "COOLRAY", "TUGELLA", "OKAVANGO", "EMGRAND", "SONG", "HAN", 
    "TANG", "YUAN", "CHAZOR", "SEAGULL", "001", "009", "TIGGO", "ARRIZO", "JOLION", "DARGO", 
    "UNI-K", "UNI-T", "UNI-V", "TXL", "DASHING", "FREE", "DREAM"
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
    template_source = st.file_uploader("Загрузите шаблон отчета (template.docx)", type="docx")

st.divider()

# --- БЛОК СКАНЕРА СТС ---
st.header("📸 1. Автозаполнение по фото")

sts_photos = st.file_uploader(
    "Загрузите фото документов (СТС, техпаспорт) - можно выбрать сразу несколько", 
    accept_multiple_files=True
)

if sts_photos:
    if st.button("Распознать документы", type="primary"):
        with st.spinner("Двойное сканирование (Оригинал + OpenCV)..."):
            try:
                all_extracted_text = ""
                for photo in sts_photos:
                    # ДВОЙНОЕ ЧТЕНИЕ: Читаем и сырую картинку (спасает от бликов), и улучшенную
                    raw_img = Image.open(photo)
                    text_raw = pytesseract.image_to_string(raw_img, lang='rus+eng')
                    
                    enhanced_img = enhance_image_for_ocr(photo)
                    text_enh = pytesseract.image_to_string(enhanced_img, lang='rus+eng')
                    
                    all_extracted_text += text_raw + "\n\n" + text_enh + "\n\n"
                
                lines = [line.strip() for line in all_extracted_text.split('\n') if line.strip()]
                upper_lines = [line.upper() for line in lines]

                clean_text_no_spaces = all_extracted_text.replace(" ", "").replace("\n", "").upper().replace("O", "0")
                clean_text_upper = all_extracted_text.upper().replace("O", "0")

                # ==============================================================
                # ГЛОБАЛЬНЫЙ МАТЕМАТИЧЕСКИЙ ПОИСК
                # ==============================================================

                # 1. Госномер (Смягчили правило, чтобы ловить даже если последняя буква смазана)
                reg_match_kg = re.search(r'\d{2}KG\d{3}[A-Z]{2,3}', clean_text_no_spaces)
                if reg_match_kg:
                    st.session_state.auto_reg = reg_match_kg.group(0)
                else:
                    reg_match_old = re.search(r'\b[A-ZА-Я]\d{4}[A-ZА-Я]{1,2}\b', clean_text_upper)
                    if reg_match_old: st.session_state.auto_reg = reg_match_old.group(0).replace(" ", "")

                # 2. VIN или Номер Шасси (ЗАЩИТА ОТ "YEAROFMANUFACTURE")
                vin_matches = re.finditer(r'[A-HJ-NPR-Z0-9]{17}', clean_text_no_spaces)
                for match in vin_matches:
                    vin_candidate = match.group(0)
                    # В VIN коде ОБЯЗАТЕЛЬНО должна быть хотя бы пара цифр!
                    if sum(c.isdigit() for c in vin_candidate) >= 3:
                        st.session_state.auto_vin = vin_candidate
                        break
                
                if not st.session_state.auto_vin:
                    chassis_match = re.search(r'[A-Z0-9]{3,6}\-[0-9]{5,7}', clean_text_no_spaces)
                    if chassis_match: st.session_state.auto_vin = chassis_match.group(0)

                # 3. Номер техпаспорта (AB 1602330)
                tp_match = re.search(r'\b[A-ZА-Я]{2}\s?\d{6,7}\b', clean_text_upper)
                if tp_match: st.session_state.auto_tech_passport = tp_match.group(0)

                # 4. Год выпуска
                year_match = re.search(r'\b(199\d|200\d|201\d|202\d)\b', clean_text_upper)
                if year_match: st.session_state.auto_year = year_match.group(0)

                # 5. Объем двигателя (ЗАЩИТА: только числа от 800 до 7999, спасает от квартир и домов)
                vols = re.findall(r'\b([8-9]\d{2}|[1-7]\d{3})\b', clean_text_upper)
                if vols:
                    for v in reversed(vols):
                        if v != st.session_state.auto_year:
                            st.session_state.auto_engine = v
                            break

                # 6. Марка и Модель
                mapping = {
                    'А': 'A', 'В': 'B', 'С': 'C', 'Е': 'E', 'Н': 'H', 'К': 'K', 'М': 'M', 'О': 'O', 
                    'Р': 'P', 'Т': 'T', 'Х': 'X', 'У': 'Y', 'И': 'I', 'Л': 'L', 'Д': 'D', 'Ф': 'F', 
                    'Г': 'G', 'З': 'Z', 'Б': 'B', 'П': 'P', 'Ь': '', 'Ъ': '', 'Э': 'E', 'Ч': 'CH', 
                    'Я': 'YA', 'Ж': 'J'
                }
                lat_global_text = "".join(mapping.get(c, c) for c in clean_text_upper)
                
                found_brand = ""
                found_model = ""
                for b in CAR_BRANDS:
                    if re.search(rf'\b{b}\b', lat_global_text): found_brand = b; break
                for m in CAR_MODELS:
                    if re.search(rf'\b{m}\b', lat_global_text): found_model = m; break
                
                if found_brand or found_model:
                    st.session_state.auto_model = f"{found_brand} {found_model}".strip()

                # 7. Цвет и Тип кузова
                for line in upper_lines:
                    if not st.session_state.auto_color:
                        for c in KNOWN_COLORS:
                            if c in line and "ЦВЕТ" not in line:
                                st.session_state.auto_color = line.title(); break
                    if not st.session_state.auto_body_type:
                        for bt in KNOWN_BODIES:
                            if bt in line:
                                st.session_state.auto_body_type = line.title(); break

                # 8. Положение руля
                if "ПРАВ" in clean_text_upper: st.session_state.auto_steering = "Правый руль"
                elif "ЛЕВ" in clean_text_upper: st.session_state.auto_steering = "Левый руль"

                # ==============================================================
                # УМНЫЙ ПОИСК ТЕКСТОВЫХ БЛОКОВ (ФИО И АДРЕС)
                # ==============================================================
                
                # Запретные слова для ФИО (чтобы случайно не схватить адрес)
                bad_fio_words = ["КЫРГЫЗ", "РЕСПУБЛИКА", "ОБЛАСТЬ", "РАЙОН", "БИШКЕК", "УЛИЦА", "АДРЕС", "ДОМ", "КВ"]

                for i, text in enumerate(upper_lines):
                    
                    # ФИО
                    if "ВЛАДЕЛЕЦ" in text or "ЭЭСИ" in text or "OWNER" in text:
                        for j in range(1, 6):
                            if i + j < len(lines):
                                candidate = lines[i+j].strip()
                                # В имени не должно быть цифр (это отсеет ПИН), и должно быть минимум 2 слова
                                if not re.search(r'\d', candidate) and len(candidate.split()) >= 2:
                                    # Проверяем, нет ли в этой строке запретных слов из адреса
                                    if not any(bw in candidate.upper() for bw in bad_fio_words):
                                        st.session_state.auto_fio = candidate.title()
                                        break

                    # Адрес
                    if "ДАРЕГИ" in text or "АДРЕС" in text:
                        for j in range(1, 6):
                            if i + j < len(upper_lines):
                                if "БИШКЕК" in upper_lines[i+j] or "РЕСПУБЛИКА" in upper_lines[i+j] or "ОБЛ" in upper_lines[i+j]:
                                    addr = lines[i+j]
                                    if i + j + 1 < len(lines) and not re.search(r'\d{5}', lines[i+j+1]):
                                        if "БЕРГЕН" not in upper_lines[i+j+1] and "ОРГАН" not in upper_lines[i+j+1]:
                                            addr += " " + lines[i+j+1]
                                    st.session_state.auto_address = addr.strip()
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
