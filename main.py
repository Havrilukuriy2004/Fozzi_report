import streamlit as st
import pandas as pd
import openpyxl
import os


# Load the Excel file
@st.cache_data
def load_data(file_path):
    df = pd.read_excel(file_path)
    return df


# Filter data based on conditions
def filter_data(df, week, report_type):
    if report_type == 'со счетом':
        df_filtered = df[(df['Неделя'] <= week) &
                         (df['Наличие открытого UAH счета 2600, 2605  или 2650 у партнера в дату проводки'] == 'Да') &
                         (df['Партнер'] == 'Да')]
    else:
        df_filtered = df[(df['Неделя'] <= week) &
                         (df['Наличие открытого UAH счета 2600, 2605  или 2650 у партнера в дату проводки'] == 'Нет') &
                         (df['Партнер'] == 'Нет')]
        mask_keywords = ['банк', 'пумб', 'держ', 'обл', 'дтек', 'вдвс', 'мвс', 'дсу', 'дснс', 'дпс', 'митна', 'гук']
        df_filtered = df_filtered[~df_filtered['Поставщик'].str.contains('|'.join(mask_keywords), case=False, na=False)]
        df_filtered = df_filtered[~df_filtered['Поставщик'].str.contains('район', case=False, na=False) | df_filtered[
            'Поставщик'].str.contains('крайон', case=False, na=False)]
    return df_filtered


# Add "others" and gross total to top 10 lists
def add_others_and_total(data, col_name):
    top_data = data.nlargest(10, col_name)
    others_sum = data[~data.index.isin(top_data.index)][col_name].sum()
    top_data.loc['Others'] = others_sum
    total_sum = data[col_name].sum()
    top_data.loc['Gross Total'] = total_sum
    return top_data


# Create the dashboard
def create_dashboard(df):
    st.sidebar.header("Filters")
    selected_week = st.sidebar.selectbox("Select Week", sorted(df['Неделя'].unique()))
    selected_report_type = st.sidebar.radio("Select Report Type", ['со счетом', 'без счета'])

    filtered_data = filter_data(df, selected_week, selected_report_type)

    start_date = df[df['Неделя'] == selected_week]['Дата проводки'].min().strftime('%d.%m.%Y')
    end_date = df[df['Неделя'] == selected_week]['Дата проводки'].max().strftime('%d.%m.%Y')

    # Title and styling
    st.markdown(f"""
        <div style="background-color:#FFA500;padding:10px;border-radius:10px">
            <h1 style="color:white;text-align:center;">Платежи на крупных контрагентов ФОЗЗИ за пределы Востока за период {start_date} - {end_date}</h1>
            <h2 style="color:white;text-align:right;">Неделя {selected_week}</h2>
        </div>
    """, unsafe_allow_html=True)

    st.header("Payments Dynamics")
    dynamics_data = filtered_data.groupby('Неделя')['Сумма'].sum().reset_index()
    st.line_chart(dynamics_data.set_index('Неделя'))

    st.header("Top Payments")
    top_payments = filtered_data.groupby('Поставщик')['Сумма'].sum().nlargest(10).reset_index()
    st.table(top_payments)

    st.header("Supplier-Payer Matrix")
    matrix_data = filtered_data.pivot_table(values='Сумма', index='Поставщик', columns='Плательщик', aggfunc='sum',
                                            fill_value=0)
    top_suppliers = add_others_and_total(matrix_data.sum(axis=1).reset_index(), 0).index
    top_payers = add_others_and_total(matrix_data.sum(axis=0).reset_index(), 0).index
    matrix_data_filtered = matrix_data.loc[top_suppliers, top_payers]
    st.table(matrix_data_filtered)

    st.header("Top Suppliers")
    supplier_data = add_others_and_total(filtered_data.groupby('Поставщик')['Сумма'].sum().reset_index(), 'Сумма')
    st.table(supplier_data)

    # Button to download Excel report
    if st.button("Download Excel Report"):
        output_excel(filtered_data, selected_week, selected_report_type, start_date, end_date)


def output_excel(df, week, report_type, start_date, end_date):
    with pd.ExcelWriter('financial_report.xlsx') as writer:
        # Sheet 1: Dynamics of payments
        dynamics_data = df.groupby('Неделя')['Сумма'].sum().reset_index()
        dynamics_data.to_excel(writer, sheet_name='Dynamics', index=False)

        # Sheet 2: Top payments by supplier and week
        supplier_data = df.groupby(['Неделя', 'Поставщик'])['Сумма'].sum().reset_index()
        supplier_data.to_excel(writer, sheet_name='Supplier Payments', index=False)

        # Sheet 3: Supplier-Payer Matrix
        matrix_data = df.pivot_table(values='Сумма', index='Поставщик', columns='Плательщик', aggfunc='sum',
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

    st.sidebar.header("raw_data_for_python.xlsx")
    file_path = st.sidebar.text_input("raw_data_for_python.xlsx")

    if file_path:
        st.write(f"Loading file from path: {file_path}")
        if os.path.exists(file_path):
            st.write("File found, loading data...")
            df = load_data(file_path)
            st.write("Data loaded successfully.")
            create_dashboard(df)
        else:
            st.error("File not found. Please enter a valid file path.")
    else:
        st.info("Please enter the file path to the Excel file.")


if __name__ == "__main__":
    main()

