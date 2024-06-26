import streamlit as st
import pandas as pd
import openpyxl
import requests
from io import BytesIO
import datetime

# Загрузка данных из Excel по URL
@st.cache_data
def load_data(url):
    response = requests.get(url)
    df = pd.read_excel(BytesIO(response.content), engine='openpyxl')
    return df

# Фильтрация данных на основе условий
def filter_data(df, week, report_type):
    if 'account' not in df.columns or 'partner' not in df.columns:
        st.error("'account' или 'partner' колонки не найдены в данных.")
        return pd.DataFrame()  # Возвращаем пустой DataFrame
    
    if report_type == 'со счетом':
        df_filtered = df[(df['week'] == week) & (df['account'].str.lower() == 'да') & (df['partner'].str.lower() == 'да')]
    else:
        df_filtered = df[(df['week'] == week) & (df['account'].str.lower() == 'нет') & (df['partner'].str.lower() == 'нет')]
        mask_keywords = ['банк', 'пумб', 'держ', 'обл', 'дтек', 'вдвс', 'мвс', 'дсу', 'дснс', 'дпс', 'митна', 'гук']
        df_filtered = df_filtered[~df_filtered['payer'].str.contains('|'.join(mask_keywords), case=False, na=False)]
        df_filtered = df_filtered[~df_filtered['payer'].str.contains('район', case=False, na=False) | df_filtered['payer'].str.contains('крайон', case=False, na=False)]
    return df_filtered

# Добавление "прочих" и общей суммы к топ-10 спискам
def add_others_and_total(data, col_name):
    top_data = data.nlargest(10, col_name)
    others_sum = data[~data.index.isin(top_data.index)][col_name].sum()
    top_data.loc['Others'] = others_sum
    total_sum = data[col_name].sum()
    top_data.loc['Gross Total'] = total_sum
    return top_data

# Функция для получения диапазона дат для заданной недели
def get_date_range_for_week(week_number, year):
    first_day_of_year = datetime.datetime(year, 1, 1)
    monday = first_day_of_year + datetime.timedelta(weeks=int(week_number) - 1, days=-first_day_of_year.weekday())
    sunday = monday + datetime.timedelta(days=6)
    return monday, sunday

# Создание дашборда
def create_dashboard(df):
    st.sidebar.header("Фильтры")
    selected_week = st.sidebar.selectbox("Выберите неделю", sorted(df['week'].unique()))
    selected_report_type = st.sidebar.radio("Выберите тип отчета", ['со счетом', 'без счета'])

    st.write(f"Выбранная неделя: {selected_week}")
    st.write(f"Выбранный тип отчета: {selected_report_type}")

    filtered_data = filter_data(df, selected_week, selected_report_type)
    st.write(f"Фильтрованные данные: {filtered_data.shape}")

    start_date, end_date = get_date_range_for_week(selected_week, 2024)
    start_date_str = start_date.strftime('%d.%m.%Y')
    end_date_str = end_date.strftime('%d.%м.%Y')

    # Заголовок и стили
    st.markdown(f"""
        <div style="background-color:#FFA500;padding:10px;border-radius:10px">
            <h1 style="color:white;text-align:center;">Платежи на крупных контрагентов ФОЗЗИ за пределы Востока за период {start_date_str} - {end_date_str}</h1>
            <h2 style="color:white;text-align:right;">Неделя {selected_week}</h2>
        </div>
    """, unsafe_allow_html=True)

    st.header("Динамика платежей")
    if not filtered_data.empty:
        dynamics_data = df.groupby('week')['sum'].sum().reset_index()
        st.line_chart(dynamics_data.set_index('week'))
    else:
        st.write("Нет данных для выбранных фильтров.")

    st.header("Топ платежей")
    if not filtered_data.empty:
        top_payments = filtered_data.groupby('payer')['sum'].sum().nlargest(10).reset_index()
        st.table(top_payments)
    else:
        st.write("Нет данных для выбранных фильтров.")

    st.header("Матрица Поставщик-Плательщик")
    if not filtered_data.empty:
        matrix_data = filtered_data.pivot_table(values='sum', index='payer', columns='recipient', aggfunc='sum', fill_value=0)

        supplier_sums = matrix_data.sum(axis=1).sort_values(ascending=False)
        payer_sums = matrix_data.sum(axis=0).sort_values(ascending=False)

        top_suppliers = add_others_and_total(supplier_sums.reset_index(), 0).index
        top_payers = add_others_and_total(payer_sums.reset_index(), 0).index

        matrix_data.loc['Others'] = matrix_data.loc[~matrix_data.index.isin(top_suppliers)].sum()
        matrix_data.loc['Gross Total'] = matrix_data.sum()
        matrix_data['Others'] = matrix_data[~matrix_data.columns.isin(top_payers)].sum(axis=1)
        matrix_data['Gross Total'] = matrix_data.sum(axis=1)

        top_suppliers = top_suppliers.tolist() + ['Others', 'Gross Total']
        top_payers = top_payers.tolist() + ['Others', 'Gross Total']
        matrix_data_filtered = matrix_data.loc[top_suppliers, top_payers]

        st.table(matrix_data_filtered)
    else:
        st.write("Нет данных для выбранных фильтров.")

    st.header("Топ поставщиков")
    if not filtered_data.empty:
        supplier_data = add_others_and_total(filtered_data.groupby('payer')['sum'].sum().reset_index(), 'sum')
        st.table(supplier_data)
    else:
        st.write("Нет данных для выбранных фильтров.")

    if st.button("Скачать отчет в формате Excel"):
        output_excel(filtered_data, selected_week, selected_report_type, start_date_str, end_date_str)

def output_excel(df, week, report_type, start_date, end_date):
    with pd.ExcelWriter('financial_report.xlsx') as writer:
        dynamics_data = df.groupby('week')['sum'].sum().reset_index()
        dynamics_data.to_excel(writer, sheet_name='Динамика', index=False)

        supplier_data = df.groupby(['week', 'payer'])['sum'].sum().reset_index()
        supplier_data.to_excel(writer, sheet_name='Платежи по поставщикам', index=False)

        matrix_data = df.pivot_table(values='sum', index='payer', columns='recipient', aggfunc='sum', fill_value=0)
        top_suppliers = add_others_and_total(matrix_data.sum(axis=1).reset_index(), 0).index
        top_payers = add_others_and_total(matrix_data.sum(axis=0).reset_index(), 0).index

        matrix_data.loc['Others'] = matrix_data.loc[~matrix_data.index.isin(top_suppliers)].sum()
        matrix_data.loc['Gross Total'] = matrix_data.sum()
        matrix_data['Others'] = matrix_data[~matrix_data.columns.isin(top_payers)].sum(axis=1)
        matrix_data['Gross Total'] = matrix_data.sum(axis=1)

        top_suppliers = top_suppliers.tolist() + ['Others', 'Gross Total']
        top_payers = top_payers.tolist() + ['Others', 'Gross Total']
        matrix_data_filtered = matrix_data.loc[top_suppliers, top_payers]

        matrix_data_filtered.to_excel(writer, sheet_name='Матрица поставщик-плательщик', index=True)

    with open('financial_report.xlsx', 'rb') as f:
        st.download_button('Скачать отчет в формате Excel', f, file_name='financial_report.xlsx')

# Основная функция для запуска Streamlit приложения
def main():
    st.set_page_config(layout="wide")

    df = load_data("https://raw.githubusercontent.com/Havrilukuriy2004/Fozzi_report/main/raw_data_for_python_final.xlsx")
    st.write("Данные успешно загружены.")
    create_dashboard(df)

if __name__ == "__main__":
    main()
