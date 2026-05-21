import pandas as pd
import streamlit as st
import re
import sqlite3
import time


# =========================
# Налаштування сторінки
# =========================

st.set_page_config(
    page_title="Обробка транзакцій",
    page_icon="favicon.ico",
)

st.markdown(
    """
    <h1 style='text-align: center; color: #333;'>Обробка транзакцій</h1>
    """,
    unsafe_allow_html=True
)


# =========================
# Локальна база SQLite
# =========================

DB_PATH = "local_cards.db"


def get_connection():
    return sqlite3.connect(DB_PATH)


def init_db():
    try:
        conn = get_connection()
        cursor = conn.cursor()

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS card_names (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                card_number TEXT NOT NULL UNIQUE,
                card_name TEXT NOT NULL
            )
        """)

        conn.commit()
        conn.close()

    except Exception as e:
        st.error(f"Помилка при створенні локальної бази: {str(e)}")


init_db()


# =========================
# Допоміжні функції
# =========================

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


# =========================
# Робота з картами в SQLite
# =========================

def load_card_names():
    try:
        conn = get_connection()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT card_number, card_name
            FROM card_names
            ORDER BY card_number
        """)

        rows = cursor.fetchall()
        conn.close()

        return {
            card_number: card_name
            for card_number, card_name in rows
        }

    except Exception as e:
        st.error(f"Помилка при читанні карток з локальної бази: {str(e)}")
        return {}


def save_card_name(card_number, card_name):
    try:
        conn = get_connection()
        cursor = conn.cursor()

        cursor.execute("""
            INSERT INTO card_names (card_number, card_name)
            VALUES (?, ?)
        """, (card_number, card_name))

        conn.commit()
        conn.close()

        st.session_state.card_names = load_card_names()

    except sqlite3.IntegrityError:
        st.error("Такий номер карти вже існує!")

    except Exception as e:
        st.error(f"Помилка при збереженні картки в локальну базу: {str(e)}")


def delete_card_name(card_number):
    try:
        conn = get_connection()
        cursor = conn.cursor()

        cursor.execute("""
            DELETE FROM card_names
            WHERE card_number = ?
        """, (card_number,))

        conn.commit()
        conn.close()

        st.session_state.card_names = load_card_names()

    except Exception as e:
        st.error(f"Помилка при видаленні картки з локальної бази: {str(e)}")


# =========================
# Ініціалізація session_state
# =========================

if "card_names" not in st.session_state:
    st.session_state.card_names = load_card_names()

if "last_added" not in st.session_state:
    st.session_state.last_added = 0


# =========================
# Додавання нової карти
# =========================

with st.expander("Додати нову карту"):
    card_number = st.text_input(
        "Номер карти (4 цифри):",
        max_chars=4,
        key="add_card_number"
    )

    card_name = st.text_input(
        "Назва карти:",
        key="add_card_name"
    )

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


# =========================
# Видалення карти
# =========================

card_names = st.session_state.card_names

with st.expander("Видалити карту"):
    if card_names:
        card_options = ["— Вибрати карту —"] + list(card_names.keys())

        card_number_to_delete = st.selectbox(
            "Вибір карти:",
            card_options,
            index=0,
            key="delete_card"
        )

        if st.button("Видалити"):
            if card_number_to_delete == "— Вибрати карту —":
                st.error("Будь ласка, виберіть карту для видалення!")
            else:
                card_name = card_names[card_number_to_delete]
                delete_card_name(card_number_to_delete)
                st.success(f"Карту {card_name} ({card_number_to_delete}) видалено!")

    else:
        st.warning("Немає карток для видалення.")


# =========================
# Перегляд поточних карт
# =========================

with st.expander("Переглянути поточні карти"):
    card_names = st.session_state.card_names

    if card_names:
        for number, name in sorted(card_names.items()):
            st.write(f"{number}: {name}")
    else:
        st.write("— немає карток —")


# =========================
# Обробка Excel-файлу
# =========================

pd.set_option('display.max_colwidth', None)

upload_file = st.file_uploader(
    "Оберіть файл Excel:",
    type=["xlsx"]
)

if upload_file:
    try:
        data = pd.read_excel(upload_file)

        required_columns = ["Unnamed: 2", "Unnamed: 5"]

        if not all(col in data.columns for col in required_columns):
            st.error(
                f"Файл не містить потрібних стовпців: {required_columns}. "
                f"Доступні стовпці: {data.columns.tolist()}"
            )
            st.stop()

        data = data.rename(columns={
            "Unnamed: 2": "Describe",
            "Unnamed: 5": "Credits"
        })

    except Exception as e:
        st.error(f"Помилка при читанні файлу: {str(e)}")
        st.stop()

    card_names = st.session_state.card_names

    data["Credits"] = data["Credits"].apply(clean)
    data["Bank_acount"] = data["Describe"].apply(extract_digit)
    data["Card_Name"] = data["Bank_acount"].map(card_names)

    data["Card_Name"] = data["Card_Name"].fillna(
        '<span style="color:red; font-weight:bold;">ФОП не вказаний</span>'
    )

    filtered_data = data[data["Bank_acount"].notna()]

    if filtered_data.empty:
        st.warning("Немає даних із номерами карток для обробки.")

    else:
        result_data = filtered_data[[
            "Bank_acount",
            "Card_Name",
            "Credits"
        ]]

        table = result_data.pivot_table(
            values="Credits",
            index=["Bank_acount", "Card_Name"],
            aggfunc="sum"
        )

        table = (
            table
            .reset_index()
            .sort_values(by="Card_Name")
            .set_index(["Bank_acount", "Card_Name"])
        )

        total_sum = table["Credits"].sum()

        table["Credits"] = table["Credits"].apply(lambda x: f"{x:,.0f}")

        st.write("Результат обробки даних")

        table_html = table.to_html(escape=False)

        table_html = table_html.replace(
            "<td>ФОП не вказаний</td>",
            '<td><span style="color:red; font-weight:bold;">ФОП не вказаний</span></td>'
        )

        st.markdown(table_html, unsafe_allow_html=True)

        st.write(f"Загальна сума: {total_sum:,.0f}")