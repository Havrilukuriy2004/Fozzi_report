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

# Main function to run the Streamlit app
def main():
    st.set_page_config(layout="wide")

    st.sidebar.header("Upload Excel File")
    file_path = st.sidebar.text_input("Enter the path to the Excel file:", value="")

    if file_path:
        st.write(f"Loading file from path: {file_path}")
        if os.path.exists(file_path):
            st.write("File found, loading data...")
            df = load_data(file_path)
            st.write("Data loaded successfully.")
            
            # Display the first few rows of the dataframe for verification
            st.write(df.head())

            # Rest of the dashboard code
            st.sidebar.header("Filters")
            selected_week = st.sidebar.selectbox("Select Week", sorted(df['Неделя'].unique()))
            selected_report_type = st.sidebar.radio("Select Report Type", ['со счетом', 'без счета'])

            st.write(f"Selected Week: {selected_week}")
            st.write(f"Selected Report Type: {selected_report_type}")

            filtered_data = filter_data(df, selected_week, selected_report_type)
            st.write(f"Filtered data shape: {filtered_data.shape}")  # Check the shape of filtered data

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
            
        else:
            st.error("File not found. Please enter a valid file path.")
    else:
        st.info("Please enter the file path to the Excel file.")

if __name__ == "__main__":
    main()
