import streamlit as st
import pandas as pd
import openpyxl
import requests
from io import BytesIO
import datetime
import altair as alt

@st.cache_data
def load_data(url):
    response = requests.get(url)
    df = pd.read_excel(BytesIO(response.content), engine='openpyxl')
    return df

def filter_data(df, week, report_type):
    if 'account' not in df.columns or 'partner' not in df.columns:
        st.error("'account' или 'partner' колонки не найдены в данных.")
        return pd.DataFrame()
    
    if report_type == 'со счетом':
        df_filtered = df[(df['week'] == week) & (df['account'].str.lower() == 'да') & (df['partner'].str.lower() == 'да')]
    else:
        df_filtered = df[(df['week'] == week) & (df['account'].str.lower() == 'нет') & (df['partner'].str.lower() == 'нет')]

    mask_keywords = ['банк', 'пумб', 'держ', 'обл', 'дтек', 'вдвс', 'мвс', 'дсу', 'дснс', 'дпс', 'митна', 'гук']
    df_filtered = df_filtered[~df_filtered['recipient'].str.contains('|'.join(mask_keywords), case=False, na=False)]
    df_filtered = df_filtered[~df_filtered['recipient'].str.contains('район', case=False, na=False) | 
                              df_filtered['recipient'].str.contains('крайон', case=False, na=False)]
    
    return df_filtered

def add_others_and_total(data, col_name):
    top_data = data.nlargest(10, col_name)
    others_sum = data[~data.index.isin(top_data.index)][col_name].sum()
    top_data.loc['Others'] = others_sum
    total_sum = data[col_name].sum()
    top_data.loc['Gross Total'] = total_sum
    return top_data

def get_date_range_for_week(week_number, year):
    first_day_of_year = datetime.datetime(year, 1, 1)
    monday = first_day_of_year + datetime.timedelta(weeks=int(week_number) - 1, days=-first_day_of_year.weekday())
    sunday = monday + datetime.timedelta(days=6)
    return monday, sunday

def create_dashboard(df):
    st.sidebar.header("Фильтры")
    selected_week = st.sidebar.selectbox("Выберите неделю", sorted(df['week'].unique()))
    selected_report_type = st.sidebar.radio("Выберите тип отчета", ['со счетом', 'без счета'])

    filtered_data = filter_data(df, selected_week, selected_report_type)

    start_date, end_date = get_date_range_for_week(selected_week, 2024)
    start_date_str = start_date.strftime('%d.%m.%Y')
    end_date_str = end_date.strftime('%d.%m.%Y')

    st.markdown(f"""
        <div style="background-color:#FFA500;padding:10px;border-radius:10px">
            <h1 style="color:white;text-align:center;">Платежи на крупных контрагентов ФОЗЗИ за пределы Востока за период {start_date_str} - {end_date_str}</h1>
            <h2 style="color:white;text-align:right;">Неделя {selected_week}</h2>
        </div>
    """, unsafe_allow_html=True)

    st.header("Динамика платежей")
    if not filtered_data.empty:
        dynamics_data = df[df['week'] <= selected_week].groupby('week')['sum'].sum().reset_index()
        dynamics_data['sum'] = dynamics_data['sum'] / 1000  # Перевод в тыс. грн

        line_chart = alt.Chart(dynamics_data).mark_line(point=alt.OverlayMarkDef(), color='#FF4500').encode(
            x='week:O',
            y=alt.Y('sum:Q', axis=alt.Axis(format=',.0f', title='Сумма (тыс. грн)')),
            tooltip=['week', alt.Tooltip('sum:Q', format=',.0f')]
        ).properties(
            title='Динамика платежей по неделям'
        ).configure_view(
            background='#FFE4B5'
        ).configure_axis(
            grid=False
        ).configure_title(
            color='#FF4500'
        ).interactive()

        st.altair_chart(line_chart, use_container_width=True)

        recipient_totals = filtered_data.groupby("recipient")["sum"].sum()
        top_10_recipients = recipient_totals.nlargest(10).index

        recipients_pivot = filtered_data.pivot_table(values='sum', index='recipient', columns='week', aggfunc='sum', fill_value=0)
        recipients_pivot = recipients_pivot.loc[top_10_recipients]
        recipients_pivot['Total'] = recipients_pivot.sum(axis=1) / 1000  # Перевод в тыс. грн

        other_data = filtered_data[~filtered_data["recipient"].isin(top_10_recipients)]
        other_totals = other_data.groupby('week')['sum'].sum()
        other_totals['Total'] = other_totals.sum() / 1000  # Перевод в тыс. грн
        recipients_pivot.loc['Others'] = other_totals

        recipients_pivot = recipients_pivot.apply(pd.to_numeric, errors='coerce')  # Convert all values to numeric, setting errors to NaN
        recipients_pivot = recipients_pivot.fillna(0)  # Fill NaN with 0

        st.table(recipients_pivot.style.format("{:,.0f}").set_table_styles([
            {
                'selector': 'th',
                'props': [('background-color', '#FFA500'), ('color', 'white')]
            },
            {
                'selector': 'td',
                'props': [('background-color', '#FFE4B5')]
            }
        ]))
    else:
        st.write("Нет данных для выбранных фильтров.")

    st.header("Топ получателей")
    if not filtered_data.empty:
        top_recipients = filtered_data.groupby(['code', 'recipient'])['sum'].sum().nlargest(10).reset_index()
        others_sum = filtered_data[~filtered_data['recipient'].isin(top_recipients['recipient'])]['sum'].sum() / 1000  # Перевод в тыс. грн
        total_sum = filtered_data['sum'].sum() / 1000  # Перевод в тыс. грн

        top_recipients['sum'] = top_recipients['sum'] / 1000  # Перевод в тыс. грн
        top_recipients.loc[len(top_recipients.index)] = ['Другие', 'Другие', others_sum]
        top_recipients.loc[len(top_recipients.index)] = ['Всего', 'Всего', total_sum]

        top_recipients = top_recipients.apply(pd.to_numeric, errors='coerce')  # Convert all values to numeric, setting errors to NaN
        top_recipients = top_recipients.fillna(0)  # Fill NaN with 0

        st.table(top_recipients.rename(columns={'code': 'Код получателя', 'recipient': 'Получатель', 'sum': 'Сума (тыс. грн)'}).style.format({'Сума (тыс. грн)': '{:,.0f}'}).set_table_styles([
            {
                'selector': 'th',
                'props': [('background-color', '#FFA500'), ('color', 'white')]
            },
            {
                'selector': 'td',
                'props': [('background-color', '#FFE4B5')]
            }
        ]))
    else:
        st.write("Нет данных для выбранных фильтров.")

    st.header("Матрица Поставщик-Плательщик")
    if not filtered_data.empty:
        recipient_totals = filtered_data.groupby("recipient")["sum"].sum().reset_index()
        top_10_recipients = recipient_totals.sort_values(by="sum", ascending=False).head(10)["recipient"]

        payer_totals = filtered_data.groupby("payer")["sum"].sum().reset_index()
        top_10_payers = payer_totals.sort_values(by="sum", ascending=False).head(10)["payer"]

        summary_data = []
        for recipient in top_10_recipients:
            recipient_data = filtered_data[filtered_data["recipient"] == recipient]
            row = [recipient] + [recipient_data[recipient_data["payer"] == payer]["sum"].sum() for payer in top_10_payers] + [recipient_data["sum"].sum()]
            summary_data.append(row)

        other_data = filtered_data[~filtered_data["recipient"].isin(top_10_recipients)]
        other_row = ["Others"] + [other_data[other_data["payer"] == payer]["sum"].sum() for payer in top_10_payers] + [other_data["sum"].sum()]

        totals_row = ["Total"] + [filtered_data[filtered_data["payer"] == payer]["sum"].sum() for payer in top_10_payers] + [filtered_data["sum"].sum()]

        summary_data.append(other_row)
        summary_data.append(totals_row)

        columns = ["Recipient"] + list(top_10_payers) + ["Total"]
        summary_df = pd.DataFrame(summary_data, columns=columns)
        summary_df.iloc[:, 1:] = summary_df.iloc[:, 1:] / 1000  # Перевод в тыс. грн

        styled_df = summary_df.style.format({
            'Recipient': '{}',
            **{payer: '{:,.0f}' for payer in top_10_payers},
            'Total': '{:,.0f}'
        }).set_table_styles([
            {
                'selector': 'th',
                'props': [('background-color', '#FFA500'), ('color', 'white')]
            },
            {
                'selector': 'td',
                'props': [('background-color', '#FFE4B5')]
            }
        ])

        st.table(styled_df)
    else:
        st.write("Нет данных для выбранных фильтров.")

    st.header("Топ плательщиков")
    if not filtered_data.empty:
        supplier_totals = filtered_data.groupby("payer")["sum"].sum().nlargest(10).reset_index()
        others_sum = filtered_data[~filtered_data["payer"].isin(supplier_totals["payer"])]["sum"].sum() / 1000  # Перевод в тыс. грн
        total_sum = filtered_data["sum"].sum() / 1000  # Перевод в тыс. грн

        supplier_totals["sum"] = supplier_totals["sum"] / 1000  # Перевод в тыс. грн
        supplier_totals.loc[len(supplier_totals.index)] = ["Другие", others_sum]
        supplier_totals.loc[len(supplier_totals.index)] = ["Всего", total_sum]

        supplier_totals = supplier_totals.apply(pd.to_numeric, errors='coerce')  # Convert all values to numeric, setting errors to NaN
        supplier_totals = supplier_totals.fillna(0)  # Fill NaN with 0

        st.table(supplier_totals.rename(columns={"payer": "Плательщик", "sum": "Сума (тыс. грн)"}).style.format({"Сума (тыс. грн)": "{:,.0f}"}).set_table_styles([
            {
                'selector': 'th',
                'props': [('background-color', '#FFA500'), ('color', 'white')]
            },
            {
                'selector': 'td',
                'props': [('background-color', '#FFE4B5')]
            }
        ]))
    else:
        st.write("Нет данных для выбранных фильтров.")

    if st.button("Скачать отчет в формате Excel"):
        output_excel(filtered_data, selected_week, selected_report_type, start_date_str, end_date_str)

def output_excel(df, week, report_type, start_date, end_date):
    with pd.ExcelWriter("financial_report.xlsx", engine="openpyxl") as writer:
        dynamics_data = df[df["week"] <= week].groupby("week")["sum"].sum().reset_index()
        dynamics_data["sum"] = dynamics_data["sum"] / 1000  # Перевод в тыс. грн
        dynamics_data.to_excel(writer, sheet_name="Динамика", index=False, float_format="%.2f")

        supplier_data = df.groupby(["week", "payer"])["sum"].sum().reset_index()
        supplier_data["sum"] = supplier_data["sum"] / 1000  # Перевод в тыс. грн
        supplier_data.to_excel(writer, sheet_name="Платежи по поставщикам", index=False, float_format="%.2f")

        matrix_data = df.pivot_table(values="sum", index="payer", columns="recipient", aggfunc="sum", fill_value=0)
        top_suppliers = matrix_data.sum(axis=1).nlargest(10).index
        top_payers = matrix_data.sum(axis=0).nlargest(10).index

        matrix_data.loc["Others"] = matrix_data.loc[~matrix_data.index.isin(top_suppliers)].sum()
        matrix_data.loc["Gross Total"] = matrix_data.sum()
        matrix_data["Others"] = matrix_data[~matrix_data.columns.isin(top_payers)].sum(axis=1)
        matrix_data["Gross Total"] = matrix_data.sum(axis=1)

        top_suppliers = top_suppliers.tolist() + ["Others", "Gross Total"]
        top_payers = top_payers.tolist() + ["Others", "Gross Total"]
        matrix_data_filtered = matrix_data.loc[top_suppliers, top_payers] / 1000  # Перевод в тыс. грн

        matrix_data_filtered.to_excel(writer, sheet_name="Матрица", index=True, float_format="%.2f")

    st.write("Отчет успешно создан: [скачать отчет](financial_report.xlsx)")

st.set_page_config(layout="wide")

excel_url = "https://raw.githubusercontent.com/Havrilukuriy2004/Fozzi_report/main/raw_data_for_python_final.xlsx"
df = load_data(excel_url)

if not df.empty:
    create_dashboard(df)
else:
    st.error("Не удалось загрузить данные. Проверьте URL и попробуйте снова.")

if __name__ == "__main__":
    main()
