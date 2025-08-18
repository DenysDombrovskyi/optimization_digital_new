import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import numbers

st.set_page_config(page_title="Digital Split Classifier", layout="wide")
st.title("Digital Split Classifier – Класифікація за CPM")

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
    
    # Класифікація інструментів
    df['Tier'] = 'Інші'
    if len(df) >= 3:
        df.loc[df.index[:3], 'Tier'] = 'Найвищий пріоритет'
    if len(df) >= 6:
        df.loc[df.index[3:6], 'Tier'] = 'Середній пріоритет'

    # Розрахунок бюджету
    df['Budget'] = 0.0
    
    # 1. Розподіляємо мінімальний бюджет на всі інструменти
    df['Budget'] = df['MinShare'] * total_budget
    remaining_budget = total_budget - df['Budget'].sum()
    
    # 2. Розподіляємо решту бюджету на пріоритетні групи
    if remaining_budget > 0:
        tier1_indices = df[df['Tier'] == 'Найвищий пріоритет'].index
        tier2_indices = df[df['Tier'] == 'Середній пріоритет'].index
        
        # Визначаємо, скільки ще можна виділити бюджету на кожну групу
        tier1_allocatable_budget = np.sum(df.loc[tier1_indices, 'MaxShare'] * total_budget - df.loc[tier1_indices, 'Budget'])
        tier2_allocatable_budget = np.sum(df.loc[tier2_indices, 'MaxShare'] * total_budget - df.loc[tier2_indices, 'Budget'])
        
        # Надаємо перевагу найвищому пріоритету
        if tier1_allocatable_budget > 0:
            tier1_to_allocate = min(remaining_budget, tier1_allocatable_budget)
            
            # Пропорційно розподіляємо між інструментами Tier 1
            if tier1_to_allocate > 0:
                tier1_budget_shares = (df.loc[tier1_indices, 'MaxShare'] - df.loc[tier1_indices, 'MinShare'])
                tier1_budget_shares_norm = tier1_budget_shares / tier1_budget_shares.sum()
                df.loc[tier1_indices, 'Budget'] += tier1_budget_shares_norm * tier1_to_allocate
            
            remaining_budget -= tier1_to_allocate
        
        # Розподіляємо решту на середній пріоритет
        if remaining_budget > 0 and tier2_allocatable_budget > 0:
            tier2_to_allocate = min(remaining_budget, tier2_allocatable_budget)
            
            # Пропорційно розподіляємо між інструментами Tier 2
            if tier2_to_allocate > 0:
                tier2_budget_shares = (df.loc[tier2_indices, 'MaxShare'] - df.loc[tier2_indices, 'MinShare'])
                tier2_budget_shares_norm = tier2_budget_shares / tier2_budget_shares.sum()
                df.loc[tier2_indices, 'Budget'] += tier2_budget_shares_norm * tier2_to_allocate
                
            remaining_budget -= tier2_to_allocate
            
    # Фінальний розрахунок метрик
    df["BudgetSharePct"] = df["Budget"] / total_budget
    df["Impressions"] = df["Budget"] / df["CPM"] * 1000
    df["Unique Reach (People)"] = df["Impressions"] / df["Freq"]
    df["ReachPct"] = df["Unique Reach (People)"] / total_audience
    df["ReachPct"] = df["ReachPct"].clip(upper=1.0)
    
    total_reach_prob = 1 - np.prod(1 - df["ReachPct"])
    total_reach_people = total_reach_prob * total_audience

    st.subheader("Результат класифікованого розподілу")
    st.write(f"Total Reach: **{total_reach_prob*100:.2f}%**")
    
    display_cols = ["Instrument", "CPM", "Tier", "MinShare", "MaxShare", "Budget", "BudgetSharePct", "Impressions", "Unique Reach (People)", "ReachPct"]
    st.dataframe(df[display_cols])

    # ==== Завантаження результатів у Excel ====
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Класифікований спліт"
    
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
        file_name="Digital_Split_Classified.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
