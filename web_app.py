import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import io
import os
import pytesseract
from PIL import Image
import re
import difflib

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

# =========================================================
# БАЗА ДАННЫХ АВТО (ОБНОВЛЕНО С УЧЕТОМ КИТАЙСКОГО АВТОПРОМА)
# =========================================================

CAR_BRANDS = [
    "TOYOTA", "HONDA", "HYUNDAI", "KIA", "LEXUS", "BMW", "MERCEDES-BENZ", "NISSAN", 
    "VOLKSWAGEN", "AUDI", "SUBARU", "CHEVROLET", "FORD", "MAZDA", "MITSUBISHI", 
    "RENAULT", "SKODA", "LADA", "PORSCHE", "LAND ROVER", "DAEWOO", 
    # КИТАЙСКИЕ БРЕНДЫ:
    "GEELY", "BYD", "ZEEKR", "LIXIANG", "LI", "CHANGAN", "CHERY", "HAVAL", 
    "EXEED", "OMODA", "JETOUR", "TANK", "HONGQI", "FAW", "JAC", "VOYAH", "DONGFENG"
]

CAR_MODELS = [
    "ALPHARD", "CAMRY", "COROLLA", "RAV4", "LAND CRUISER", "PRADO", "HIGHLANDER", "PRIUS", 
    "FIT", "ACCORD", "CR-V", "CIVIC", "ODYSSEY", "STEPWGN", "SONATA", "ELANTRA", "SANTA FE", 
    "TUCSON", "K5", "RIO", "SPORTAGE", "SORENTO", "MORNING", "RX", "LX", "ES", "GX", "GS", 
    "IS", "NX", "RX300", "RX330", "RX350", "LX470", "LX570", "X5", "X3", "X6", "X7", "E-CLASS", 
    "C-CLASS", "S-CLASS", "SPRINTER", "FOCUS", "TRANSIT", "GOLF", "PASSAT", "JETTA", "TIGUAN", 
    "POLO", "FORESTER", "OUTBACK", "IMPREZA", "LEGACY", "CRUZE", "MALIBU", "COBALT", "NEXIA", 
    "MATIZ", "SPARK", "ESTIMA", "WISH", "NOAH", "VOXY", "HARRIER", "AVANTE", "PALISADE",
    # КИТАЙСКИЕ МОДЕЛИ:
    "L7", "L8", "L9", "MONJARO", "COOLRAY", "TUGELLA", "OKAVANGO", "EMGRAND",
    "SONG", "HAN", "TANG", "YUAN", "CHAZOR", "SEAGULL", "001", "009", "X", 
    "TIGGO", "ARRIZO", "JOLION", "DARGO", "H6", "UNI-K", "UNI-T", "UNI-V", 
    "TXL", "VX", "RX", "DASHING", "300", "500", "FREE", "DREAM"
]

def fix_car_name(text, known_list):
    if not text: return ""
    text = re.sub(r'[^А-ЯA-Z0-9\s\-]', '', text.upper())
    
    mapping = {
        'А': 'A', 'В': 'B', 'С': 'C', 'Е': 'E', 'Н': 'H', 'К': 'K', 'М': 'M', 'О': 'O', 
        'Р': 'P', 'Т': 'T', 'Х': 'X', 'У': 'Y', 'И': 'I', 'Л': 'L', 'Д': 'D', 'Ф': 'F', 
        'Г': 'G', 'З': 'Z', 'Б': 'B', 'П': 'P', 'Ь': '', 'Ъ': '', 'Э': 'E', 'Ч': 'CH', 
        'Я': 'YA', 'Ю': 'YU', 'Ц': 'C', 'Й': 'Y', 'Ж': 'J'
    }
    lat_text = "".join(mapping.get(c, c) for c in text).strip()
    
    # ПОВЫСИЛИ ПОРОГ ДО 60% (cutoff=0.60), чтобы избежать галлюцинаций вроде "Одиссея"
    if lat_text:
        matches = difflib.get_close_matches(lat_text, known_list, n=1, cutoff=0.60)
        if matches:
            return matches[0] 
    return lat_text # Если не уверен, вернет то, что прочитал

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
        with st.spinner("Нейросеть Google Tesseract читает документы..."):
            try:
                all_extracted_text = ""
                for photo in sts_photos:
                    img = Image.open(photo)
                    text = pytesseract.image_to_string(img, lang='rus+eng')
                    all_extracted_text += text + "\n\n"
                
                lines = [line.strip() for line in all_extracted_text.split('\n') if line.strip()]
                upper_lines = [line.upper() for line in lines]
                
                brand = ""
                model = ""

                for i, text in enumerate(upper_lines):
                    
                    # 1. ФИО Владельца
                    if ("ВЛАДЕЛЕЦ" in text or "ЭЭСИ" in text or "OWNER" in text) and "АДРЕС" not in text and "ADDRESS" not in text:
                        if i + 1 < len(lines): 
                            st.session_state.auto_fio = lines[i+1].title()

                    # 2. Адрес
                    if "АДРЕС СОБСТВЕННИКА" in text or "ДАРЕГИ" in text:
                        if i + 1 < len(lines):
                            addr = lines[i+1]
                            if i + 2 < len(lines) and len(lines[i+2]) > 5 and "МАРКА" not in upper_lines[i+2]:
                                addr += " " + lines[i+2]
                            st.session_state.auto_address = addr
                            
                    # 3. Марка
                    if "МАРКАСЫ" in text or "МАРКА /" in text or text == "МАРКА" or "BRAND" in text:
                        clean_text = text.replace("МАРКАСЫ", "").replace("МАРКА", "").replace("VEHICLE", "").replace("BRAND", "").replace("/", "").strip()
                        if len(clean_text) > 2:
                            brand = clean_text
                        elif i + 1 < len(upper_lines): 
                            brand = upper_lines[i+1]
                            
                    # 4. Модель
                    if "МОДЕЛИ" in text or "МОДЕЛЬ /" in text or text == "МОДЕЛЬ" or "MODEL" in text:
                        clean_text = text.replace("МОДЕЛИ", "").replace("МОДЕЛЬ", "").replace("MODEL", "").replace("/", "").strip()
                        if len(clean_text) > 2:
                            model = clean_text
                        elif i + 1 < len(upper_lines): 
                            model = upper_lines[i+1]
                            
                    # 5. Год выпуска
                    if "ГОД ВЫПУСКА" in text or "ЖЫЛЫ" in text:
                        for j in range(0, 3):
                            if i + j < len(upper_lines):
                                y_match = re.search(r'\b(19|20)\d{2}\b', upper_lines[i+j])
                                if y_match: 
                                    st.session_state.auto_year = y_match.group(0)
                                    break

                    # 6. Номер Кузова / Шасси / VIN
                    if "ШАССИ" in text or "КУЗОВА" in text or "VIN" in text:
                        for j in range(0, 3):
                            if i + j < len(upper_lines):
                                v_match = re.search(r'[A-Z0-9\-]{9,17}', upper_lines[i+j].replace("O", "0"))
                                if v_match and "ОТМЕТКИ" not in upper_lines[i+j]:
                                    st.session_state.auto_vin = v_match.group(0)
                                    break

                    # 7. Объем двигателя
                    if "ОБЪЕМ" in text or "КӨЛӨМҮ" in text or "VOLUME" in text:
                        for j in range(0, 4):
                            if i + j < len(upper_lines):
                                vols = re.findall(r'\d{3,4}', upper_lines[i+j].replace("O", "0"))
                                if vols:
                                    st.session_state.auto_engine = vols[-1]
                                    break

                    # 8. Цвет
                    if "ЦВЕТ" in text or "ТҮСҮ" in text:
                        for j in range(1, 3):
                            if i + j < len(lines):
                                if len(lines[i+j]) > 3 and "КУЗОВ" not in upper_lines[i+j]:
                                    st.session_state.auto_color = lines[i+j].capitalize()
                                    break

                    # 9. Положение руля
                    if "РУЛЬ" in text or "РАСПОЛОЖЕНИЕ РУЛЯ" in text:
                        for j in range(0, 3):
                            if i + j < len(upper_lines):
                                if "ПРАВ" in upper_lines[i+j]:
                                    st.session_state.auto_steering = "Правый руль"
                                    break
                                elif "ЛЕВ" in upper_lines[i+j]:
                                    st.session_state.auto_steering = "Левый руль"
                                    break
                                    
                    # 10. Номер техпаспорта 
                    if "СЕРИЯ И НОМЕР" in text or "СЕРИЯСЫ" in text:
                        for j in range(0, 3):
                            if i + j < len(upper_lines):
                                clean_line = upper_lines[i+j].replace("O", "0")
                                tp = re.search(r'[A-ZА-Я]{2}\s?\d{6,7}', clean_line)
                                if tp:
                                    st.session_state.auto_tech_passport = tp.group(0)
                                    break
                                    
                    # 11. Тип кузова
                    if "ВИД КУЗОВА" in text or "ТИП ТС" in text or "КУЗОВТУН ТҮРҮ" in text:
                        for j in range(1, 3):
                            if i + j < len(lines):
                                if "VEHICLE" not in upper_lines[i+j] and len(lines[i+j]) > 3:
                                    st.session_state.auto_body_type = lines[i+j].capitalize()
                                    break

                # ПРИМЕНЯЕМ ОБНОВЛЕННЫЙ АВТОКОРРЕКТОР С БАЗОЙ КИТАЙСКИХ АВТО!
                fixed_brand = fix_car_name(brand, CAR_BRANDS)
                fixed_model = fix_car_name(model, CAR_MODELS)
                
                if fixed_brand or fixed_model:
                    st.session_state.auto_model = f"{fixed_brand} {fixed_model}".strip()

                # Госномер
                clean_string = all_extracted_text.replace(" ", "").replace("\n", "").upper().replace("O", "0")
                reg_match_kg = re.search(r'\d{2}KG\d{3}[A-Z]{3}', clean_string)
                
                if reg_match_kg:
                    st.session_state.auto_reg = reg_match_kg.group(0)
                else:
                    reg_match_old = re.search(r'\b[A-ZА-Я]\d{4}[A-ZА-Я]{1,2}\b', all_extracted_text.upper())
                    if reg_match_old:
                        st.session_state.auto_reg = reg_match_old.group(0).replace(" ", "")

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
