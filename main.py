import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import datetime
import altair as alt

@st.cache
def load_data(url):
    response = requests.get(url)
    df = pd.read_excel(BytesIO(response.content), engine='openpyxl')
    return df

def filter_data(df, week, report_type):
    if 'account' not in df.columns or 'partner' not in df.columns:
        st.error("'account' или 'partner' колонки не найдены в данных.")
        return pd.DataFrame()

    if report_type == 'з відкритим рахунком':
        df_filtered = df[
            (df['week'] <= week) & (df['account'].str.lower() == 'да') & (df['partner'].str.lower() == 'да')]
    else:
        df_filtered = df[
            (df['week'] <= week) & (df['account'].str.lower() == 'нет') & (df['partner'].str.lower() == 'нет')]

    mask_keywords = ['банк', 'пумб', 'держ', 'обл', 'дтек', 'вдвс', 'мвс', 'дсу', 'дснс', 'дпс', 'митна', 'гук']
    df_filtered = df_filtered[~df_filtered['recipient'].str.contains('|'.join(mask_keywords), case=False, na=False)]
    df_filtered = df_filtered[~df_filtered['recipient'].str.contains('район', case=False, na=False) |
                              df_filtered['recipient'].str.contains('крайон', case=False, na=False)]

    return df_filtered

def get_date_range_for_week(week_number, year):
    first_day_of_year = datetime.datetime(year, 1, 1)
    monday = first_day_of_year + datetime.timedelta(weeks=int(week_number) - 1, days=-first_day_of_year.weekday())
    sunday = monday + datetime.timedelta(days=6)
    return monday, sunday

def highlight_first_last(val):
    color = 'orange'
    return f'background-color: {color}'

def highlight_values(val):
    color = '#FFD580'  # Light orange
    return f'background-color: {color}'

def create_dashboard(df):
    st.sidebar.header("Фільтри")
    selected_week = st.sidebar.selectbox("Оберіть тиждень", sorted(df['week'].unique()))
    selected_report_type = st.sidebar.radio("Оберіть тип звіту", ['з відкритим рахунком', 'без відкритого рахунку'])

    filtered_data = filter_data(df, selected_week, selected_report_type)

    start_date, end_date = get_date_range_for_week(selected_week, 2024)
    start_date_str = start_date.strftime('%d.%m.%Y')
    end_date_str = end_date.strftime('%d.%m.%Y')

    st.markdown(f"""
        <div style="background-color:#FFA500;padding:10px;border-radius:10px">
            <h1 style="color:white;text-align:center;">Виплати великим контрагентам FOZZI за межами ПАТ "БАНК ВОСТОК" за період {start_date_str} - {end_date_str}</h1>
            <h2 style="color:white;text-align:right;">Тиждень {selected_week}</h2>
        </div>
    """, unsafe_allow_html=True)

    st.header("Динаміка виплат")
    if not filtered_data.empty:
        dynamics_data = df[df['week'] <= selected_week].groupby('week')['sum'].sum().reset_index()
        dynamics_data['sum'] = dynamics_data['sum'] / 1000  # Перевод в тыс. грн

        line_chart = alt.Chart(dynamics_data).mark_line(point=alt.OverlayMarkDef()).encode(
            x='week:O',
            y=alt.Y('sum:Q', axis=alt.Axis(format=',.0f', title='Сума (тис. грн)')),
            tooltip=['week', alt.Tooltip('sum:Q', format=',.0f')]
        ).properties(
            title='Динаміка виплат по тижням'
        ).interactive()

        st.altair_chart(line_chart, use_container_width=True)

        recipient_totals = filtered_data.groupby("recipient")["sum"].sum()
        top_10_recipients = recipient_totals.nlargest(10).index

        recipients_pivot = filtered_data.pivot_table(values='sum', index='recipient', columns='week', aggfunc='sum',
