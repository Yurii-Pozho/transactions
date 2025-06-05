import pandas as pd
import streamlit as st
import re
from supabase import create_client, Client
import time


# Налаштування favicon і заголовка сторінки
st.set_page_config(
    page_title="Обробка транзакцій",
    page_icon="favicon.ico",  # Файл у корені репозиторію
)

# Центрований заголовок
st.markdown(
    """
    <h1 style='text-align: center; color: #333;'>Обробка транзакцій</h1>
    """,
    unsafe_allow_html=True
)
# Налаштування Supabase
SUPABASE_URL = "https://rxnsqvukzwahvjsawpso.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJ4bnNxdnVrendhaHZqc2F3cHNvIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDkxMTUwNzcsImV4cCI6MjA2NDY5MTA3N30.Uxankt2738yR0kcCNXKWsB5Osq2pnWDlsr6gPHb-wP4"
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

def clean(value):
    value = re.sub(r'[^\d\.,-]', '', str(value))
    value = value.replace(',', '.')
    try:
        result = float(value)
        return abs(int(result))
    except ValueError:
        return 0

def extract_digit(description):
    if isinstance(description, str):
        match = re.search(r'(\d{4})\.', description)
        return match.group(1) if match else None
    return None

# Завантаження словника карт із Supabase
def load_card_names():
    try:
        response = supabase.table("card_names").select("*").execute()
        return {row["card_number"]: row["card_name"] for row in response.data}
    except Exception as e:
        st.error(f"Помилка при читанні карток: {str(e)}")
        return {}

# Збереження нової карти в Supabase
def save_card_name(card_number, card_name):
    try:
        supabase.table("card_names").insert({"card_number": card_number, "card_name": card_name}).execute()
        st.session_state.card_names = load_card_names()
    except Exception as e:
        st.error(f"Помилка при збереженні картки: {str(e)}")

# Видалення карти з Supabase
def delete_card_name(card_number):
    try:
        supabase.table("card_names").delete().eq("card_number", card_number).execute()
        st.session_state.card_names = load_card_names()
    except Exception as e:
        st.error(f"Помилка при видаленні картки: {str(e)}")

# Ініціалізація кешу
if "card_names" not in st.session_state:
    st.session_state.card_names = load_card_names()
if "last_added" not in st.session_state:
    st.session_state.last_added = 0

# Додавання нової карти у випадаючому списку
with st.expander("Додати нову карту"):
    card_number = st.text_input("Номер карти (4 цифри):", max_chars=4, key="add_card_number")
    card_name = st.text_input("Назва карти:", key="add_card_name")
    if st.button("Додати карту"):
        current_time = time.time()
        if current_time - st.session_state.last_added < 10:
            st.error("Зачекайте 10 секунд перед додаванням нової карти!")
        elif not card_number or not card_name:
            st.error("Заповніть обидва поля!")
        elif not re.match(r'^\d{4}$', card_number):
            st.error("Номер карти має містити рівно 4 цифри!")
        elif len(card_name) > 50:
            st.error("Назва карти не може бути довшою за 50 символів!")
        elif not re.match(r'^[\w\s\-\.]+$', card_name):
            st.error("Назва карти може містити лише літери, цифри, пробіли, дефіси та крапки!")
        elif card_number in st.session_state.card_names:
            st.error("Такий номер карти вже існує!")
        else:
            save_card_name(card_number, card_name)
            st.session_state.last_added = current_time
            st.success(f"Карту {card_name} ({card_number}) додано!")

# Видалення карти у випадаючому списку
card_names = st.session_state.card_names
with st.expander("Видалити карту"):
    if card_names:
        card_options = ["— Вибрати карту —"] + list(card_names.keys())
        card_number_to_delete = st.selectbox("Вибір карти:", card_options, index=0, key="delete_card")
        if st.button("Видалити"):
            if card_number_to_delete == "— Вибрати карту —":
                st.error("Будь ласка, виберіть карту для видалення!")
            else:
                card_name = card_names[card_number_to_delete]
                delete_card_name(card_number_to_delete)
                st.success(f"Карту {card_name} ({card_number_to_delete}) видалено!")
    else:
        st.warning("Немає карток для видалення.")

# Виведення поточних карт у випадающем списку
with st.expander("Переглянути поточні карти"):
    if card_names:
        for number, name in sorted(card_names.items()):
            st.write(f"{number}: {name}")
    else:
        st.write("— немає карток —")

# Обробка Excel-файлу
pd.set_option('display.max_colwidth', None)
upload_file = st.file_uploader("Оберіть файл Excel:", type=['xlsx'])

if upload_file:
    try:
        data = pd.read_excel(upload_file)
        required_columns = ['Unnamed: 2', 'Unnamed: 5']
        if not all(col in data.columns for col in required_columns):
            st.error(f"Файл не містить потрібних стовпців: {required_columns}. Доступні: {data.columns.tolist()}")
            st.stop()
        data = data.rename(columns={'Unnamed: 2': 'Describe', 'Unnamed: 5': 'Credits'})
    except Exception as e:
        st.error(f"Помилка при читанні файлу: {str(e)}")
        st.stop()

    data['Credits'] = data['Credits'].apply(clean)
    data['Bank_acount'] = data['Describe'].apply(extract_digit)
    data['Card_Name'] = data['Bank_acount'].map(card_names)
    data['Card_Name'] = data['Card_Name'].fillna('<span style="color:red; font-weight:bold;">ФОП не вказаний</span>')

    filtered_data = data[data['Bank_acount'].notna()]
    
    if filtered_data.empty:
        st.warning("Немає даних із номерами карток для обробки.")
    else:
        result_data = filtered_data[['Bank_acount', 'Card_Name', 'Credits']]
        table = result_data.pivot_table('Credits', ['Bank_acount', 'Card_Name'], aggfunc='sum')
        table = table.reset_index().sort_values(by='Card_Name').set_index(['Bank_acount', 'Card_Name'])
        
        # Обчислюємо загальну суму до форматування
        total_sum = table['Credits'].sum()
        
        # Форматуємо Credits для відображення з комами
        table['Credits'] = table['Credits'].apply(lambda x: f"{x:,.0f}")
        
        st.write('Результат обробки даних')
        table_html = table.to_html(escape=False)
        table_html = table_html.replace(
            '<td>ФОП не вказаний</td>',
            '<td><span style="color:red; font-weight:bold;">ФОП не вказаний</span></td>'
        )
        st.markdown(table_html, unsafe_allow_html=True)
        st.write(f"Загальна сума: {total_sum:,.0f}")


