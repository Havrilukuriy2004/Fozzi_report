import streamlit as st
import pandas as pd
import openpyxl
import requests
from io import BytesIO
import datetime

# Load the Excel file from URL
@st.cache
def load_data(url):
    response = requests.get(url)
    df = pd.read_excel(BytesIO(response.content), engine='openpyxl')
    return df

# Filter data based on conditions
def filter_data(df, week, report_type):
    if report_type == 'со счетом':
        df_filtered = df[(df['week'] <= week) & (df['account'] == 'Да') & (df['Партнер'] == 'Да')]
    else:
        df_filtered = df[(df['week'] <= week) & (df['account'] == 'Нет') & (df['Партнер'] == 'Нет')]
        mask_keywords = ['банк', 'пумб', 'держ', 'обл', 'дтек', 'вдвс', 'мвс', 'дсу', 'дснс', 'дпс', 'митна', 'гук']
        df_filtered = df_filtered[~df_filtered['Плательщик'].str.contains('|'.join(mask_keywords), case=False, na=False)]
        df_filtered = df_filtered[~df_filtered['Плательщик'].str.contains('район', case=False, na=False) | df_filtered[
            'Плательщик'].str.contains('крайон', case=False, na=False)]
    return df_filtered

# Add "others" and gross total to top 10 lists
def add_others_and_total(data, col_name):
    top_data = data.nlargest(10, col_name)
    others_sum = data[~data.index.isin(top_data.index)][col_name].sum()
    top_data.loc['Others'] = others_sum
    total_sum = data[col_name].sum()
    top_data.loc['Gross Total'] = total_sum
    return top_data

# Function to calculate the date range for a given week number
def get_date_range_for_week(week_number, year):
    first_day_of_year = datetime.datetime(year, 1, 1)
    # Calculate the Monday of the specified week
    monday = first_day_of_year + datetime.timedelta(weeks=int(week_number) - 1, days=-first_day_of_year.weekday())
    # Calculate the Sunday of the specified week
    sunday = monday + datetime.timedelta(days=6)
    return monday, sunday

# Create the dashboard
def create_dashboard(df):
    st.sidebar.header("Filters")
    selected_week = st.sidebar.selectbox("Select Week", sorted(df['week'].unique()))
    selected_report_type = st.sidebar.radio("Select Report Type", ['со счетом', 'без счета'])

    st.write(f"Selected Week: {selected_week}")
    st.write(f"Selected Report Type: {selected_report_type}")

    filtered_data = filter_data(df, selected_week, selected_report_type)
    st.write(f"Filtered data shape: {filtered_data.shape}")  # Check the shape of filtered data

    start_date, end_date = get_date_range_for_week(selected_week, 2024)
    start_date_str = start_date.strftime('%d.%m.%Y')
    end_date_str = end_date.strftime('%d.%m.%Y')

    # Title and styling
    st.markdown(f"""
        <div style="background-color:#FFA500;padding:10px;border-radius:10px">
            <h1 style="color:white;text-align:center;">Платежи на крупных контрагентов ФОЗЗИ за пределы Востока за период {start_date_str} - {end_date_str}</h1>
            <h2 style="color:white;text-align:right;">Неделя {selected_week}</h2>
        </div>
    """, unsafe_allow_html=True)

    st.header("Payments Dynamics")
    dynamics_data = filtered_data.groupby('week')['sum'].sum().reset_index()
    st.line_chart(dynamics_data.set_index('week'))

    st.header("Top Payments")
    top_payments = filtered_data.groupby('Плательщик')['sum'].sum().nlargest(10).reset_index()
    st.table(top_payments)

    st.header("Supplier-Payer Matrix")
    matrix_data = filtered_data.pivot_table(values='sum', index='Плательщик', columns='Получатель', aggfunc='sum',
                                            fill_value=0)
    top_suppliers = add_others_and_total(matrix_data.sum(axis=1).reset_index(), 0).index
    top_payers = add_others_and_total(matrix_data.sum(axis=0).reset_index(), 0).index
    matrix_data_filtered = matrix_data.loc[top_suppliers, top_payers]
    st.table(matrix_data_filtered)

    st.header("Top Suppliers")
    supplier_data = add_others_and_total(filtered_data.groupby('Плательщик')['sum'].sum().reset_index(), 'sum')
    st.table(supplier_data)

    # Button to download Excel report
    if st.button("Download Excel Report"):
        output_excel(filtered_data, selected_week, selected_report_type, start_date_str, end_date_str)

def output_excel(df, week, report_type, start_date, end_date):
    with pd.ExcelWriter('financial_report.xlsx', engine='openpyxl') as writer:
        # Sheet 1: Dynamics of payments
        dynamics_data = df.groupby('week')['sum'].sum().reset_index()
        dynamics_data.to_excel(writer, sheet_name='Dynamics', index=False)

        # Sheet 2: Top payments by supplier and week
        supplier_data = df.groupby(['week', 'Плательщик'])['sum'].sum().reset_index()
        supplier_data.to_excel(writer, sheet_name='Supplier Payments', index=False)

        # Sheet 3: Supplier-Payer Matrix
        matrix_data = df.pivot_table(values='sum', index='Плательщик', columns='Получатель', aggfunc='sum',
                                     fill_value=0)
        top_suppliers = add_others_and_total(matrix_data.sum(axis=1).reset_index(), 0).index
        top_payers = add_others_and_total(matrix_data.sum(axis=0).reset_index(), 0).index
        matrix_data_filtered = matrix_data.loc[top_suppliers, top_payers]
        matrix_data_filtered.to_excel(writer, sheet_name='Supplier-Payer Matrix', index=True)

    with open('financial_report.xlsx', 'rb') as f:
        st.download_button('Download Excel report', f, file_name='financial_report.xlsx')

# Main function to run the Streamlit app
def main():
    st.set_page_config(layout="wide")

    st.sidebar.header("")
    file_url = st.sidebar.text_input("Enter URL to the Excel file", value="https://raw.githubusercontent.com/Havrilukuriy2004/Fozzi_report/main/raw_data_for_python.xlsx")

    if file_url:
        st.write(f"Loading file from URL: {file_url}")
        try:
            df = load_data(file_url)
            st.write("Data loaded successfully.")
            create_dashboard(df)
        except Exception as e:
            st.error(f"Error loading data: {e}")
    else:
        st.info("Please enter the URL to the Excel file.")

if __name__ == "__main__":
    main()
