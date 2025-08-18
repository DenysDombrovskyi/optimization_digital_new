import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import numbers

st.set_page_config(page_title="Digital Split Classifier", layout="wide")
st.title("Digital Split Classifier – Автоматична класифікація")

# ==== Налаштування ====
num_instruments = st.number_input("Кількість інструментів:", min_value=1, max_value=50, value=20, step=1)
total_audience = st.number_input("Загальний розмір потенційної аудиторії:", value=50000, step=1000)
total_budget = st.number_input("Заданий бюджет ($):", value=100000, step=1000)

# ==== Початкові дані ====
if "df" not in st.session_state or len(st.session_state.df) != num_instruments:
    st.session_state.df = pd.DataFrame({
        "Instrument": [f"Instrument {i+1}" for i in range(num_instruments)],
        "CPM": [50 + i*2 for i in range(num_instruments)],
        "Freq": [1.0 + 0.05*i for i in range(num_instruments)],
        "MinShare": [0.02 for _ in range(num_instruments)],
        "MaxShare": [0.2 for i in range(num_instruments)]
    })

df = st.session_state.df.copy()

# ==== Форма редагування ====
with st.form("instrument_form"):
    for i in range(len(df)):
        st.subheader(f"Інструмент {i+1}")
        df.loc[i, "Instrument"] = st.text_input(f"Назва {i+1}", value=df.loc[i, "Instrument"])
        df.loc[i, "CPM"] = st.number_input(f"CPM {i+1}", value=float(df.loc[i, "CPM"]), step=1.0)
        df.loc[i, "Freq"] = st.number_input(f"Frequency {i+1}", value=float(df.loc[i, "Freq"]), step=0.01)
        df.loc[i, "MinShare"] = st.number_input(f"Min Share {i+1}", value=float(df.loc[i, "MinShare"]), step=0.01)
        df.loc[i, "MaxShare"] = st.number_input(f"Max Share {i+1}", value=float(df.loc[i, "MaxShare"]), step=0.01)
    submitted = st.form_submit_button("Перерахувати спліт")

if submitted:
    st.session_state.df = df.copy()
    
    df["Efficiency"] = 1000 / (df["CPM"] * df["Freq"])
    df["CPR"] = (df["CPM"] / 1000) * (df["Freq"] * total_audience) / total_audience
    
    # Сортування за CPM для класифікації
    df = df.sort_values(by="CPM", ascending=True).reset_index(drop=True)
    
    # Динамічне визначення груп
    n_total = len(df)
    n_gold = max(1, round(n_total * 0.2))  # 20% найкращих, мінімум 1
    n_silver = max(1, round(n_total * 0.3)) # 30% середніх, мінімум 1
    n_bronze = n_total - n_gold - n_silver

    df['Tier'] = 'Бронзова група'
    df.loc[df.index[:n_gold], 'Tier'] = 'Золота група'
    df.loc[df.index[n_gold:n_gold+n_silver], 'Tier'] = 'Срібна група'

    # Розрахунок бюджету
    df['Budget'] = 0.0
    
    # 1. Забезпечуємо мінімальний бюджет для всіх
    df['Budget'] = df['MinShare'] * total_budget
    remaining_budget = total_budget - df['Budget'].sum()
    
    # 2. Розподіляємо додатковий бюджет на "золоту" групу
    gold_indices = df[df['Tier'] == 'Золота група'].index
    if remaining_budget > 0 and len(gold_indices) > 0:
        gold_max_share_budget = df.loc[gold_indices, 'MaxShare'] * total_budget
        gold_additional_budget = gold_max_share_budget - df.loc[gold_indices, 'Budget']
        
        total_gold_additional = np.sum(gold_additional_budget)
        
        if total_gold_additional > 0:
            to_allocate = min(remaining_budget, total_gold_additional)
            
            # Пропорційно розподіляємо між інструментами "золотої" групи
            df.loc[gold_indices, 'Budget'] += gold_additional_budget / total_gold_additional * to_allocate
            remaining_budget -= to_allocate
            
    # 3. Розподіляємо решту бюджету на "срібну" групу
    silver_indices = df[df['Tier'] == 'Срібна група'].index
    if remaining_budget > 0 and len(silver_indices) > 0:
        silver_max_share_budget = df.loc[silver_indices, 'MaxShare'] * total_budget
        silver_additional_budget = silver_max_share_budget - df.loc[silver_indices, 'Budget']
        
        total_silver_additional = np.sum(silver_additional_budget)
        
        if total_silver_additional > 0:
            to_allocate = min(remaining_budget, total_silver_additional)
            
            # Пропорційно розподіляємо між інструментами "срібної" групи
            df.loc[silver_indices, 'Budget'] += silver_additional_budget / total_silver_additional * to_allocate
            remaining_budget -= to_allocate
            
    # Фінальний розрахунок метрик
    df["BudgetSharePct"] = df["Budget"] / total_budget
    df["Impressions"] = df["Budget"] / df["CPM"] * 1000
    df["Unique Reach (People)"] = df["Impressions"] / df["Freq"]
    df["ReachPct"] = df["Unique Reach (People)"] / total_audience
    df["ReachPct"] = df["ReachPct"].clip(upper=1.0)
    
    total_reach_prob = 1 - np.prod(1 - df["ReachPct"])
    total_reach_people = total_reach_prob * total_audience

    st.subheader("Результат автоматичної класифікації")
    st.write(f"Total Reach: **{total_reach_prob*100:.2f}%**")
    
    display_cols = ["Instrument", "CPM", "Tier", "MinShare", "MaxShare", "Budget", "BudgetSharePct", "Impressions", "Unique Reach (People)", "ReachPct"]
    st.dataframe(df[display_cols])

    # ==== Завантаження результатів у Excel ====
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Автоматичний спліт"
    
    for r in dataframe_to_rows(df[display_cols], index=False, header=True):
        ws.append(r)
    
    ws.append([])
    ws.append(["TOTAL", "", "", "", "", df["Budget"].sum(), 1.0, df["Impressions"].sum(), df["Unique Reach (People)"].sum(), f"{total_reach_prob*100:.2f}%"])
    ws.append(["TOTAL PEOPLE", "", "", "", "", "", "", "", int(total_reach_people), ""])
    
    for row in ws.iter_rows(min_row=2, max_row=1+len(df), min_col=7, max_col=7):
        for cell in row:
            cell.number_format = '0.00%'
            
    chart = BarChart()
    chart.type = "col"
    chart.title = "Бюджет по інструментам (%)"
    chart.y_axis.title = "Budget Share (%)"
    chart.x_axis.title = "Інструменти"
    data = Reference(ws, min_col=7, min_row=1, max_row=1+len(df))
    categories = Reference(ws, min_col=1, min_row=2, max_row=1+len(df))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    ws.add_chart(chart, "L2")

    wb.save(output)
    output.seek(0)
    
    st.download_button(
        label="Завантажити результати (Excel)",
        data=output,
        file_name="Digital_Split_Automatic.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
