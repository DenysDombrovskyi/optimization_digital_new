import streamlit as st
import pandas as pd
import numpy as np
from scipy.optimize import minimize
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import numbers

st.set_page_config(page_title="Digital Split Optimizer", layout="wide")
st.title("Digital Split Optimizer – Плавне зниження долей")

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
    df = df.sort_values(by="CPM", ascending=True).reset_index(drop=True)
    
    def total_reach(budgets, df_current):
        impressions = budgets / df_current["CPM"].values * 1000
        reach_i = np.clip(impressions / total_audience, 0, 1)
        return 1 - np.prod(1 - reach_i)
    
    # Визначення ідеальних долей за допомогою експоненційного спадання
    # Використовуємо індекс як "ранг" для плавного зниження
    ranks = np.arange(1, len(df) + 1)
    
    # Параметр a контролює швидкість спадання. Вищий a -> крутіше падіння.
    # Я обрав a=2 для помітного, але не надто різкого спадання
    a = 2
    ideal_shares = 1 / (ranks ** a)
    
    # Нормалізація ідеальних долей
    ideal_shares_norm = ideal_shares / ideal_shares.sum()
    
    # Розрахунок бюджету на основі ідеальних долей та Min/Max обмежень
    df_result = df.copy()
    df_result['IdealShare'] = ideal_shares_norm
    df_result['Budget'] = df_result['IdealShare'] * total_budget
    
    # Ітеративне коригування, щоб дотриматися MinShare та MaxShare
    for _ in range(10):  # Виконуємо кілька ітерацій для стабільності
        
        # Обчислюємо поточні частки перед кожною перевіркою
        df_result['BudgetSharePct'] = df_result['Budget'] / total_budget
        
        # Обробка інструментів, що виходять за межі MinShare
        min_breach_indices = df_result[df_result['BudgetSharePct'] < df_result['MinShare']].index
        
        if not min_breach_indices.empty:
            df_result.loc[min_breach_indices, 'Budget'] = df_result.loc[min_breach_indices, 'MinShare'] * total_budget
            
        # Обробка інструментів, що виходять за межі MaxShare
        max_breach_indices = df_result[df_result['BudgetSharePct'] > df_result['MaxShare']].index
        if not max_breach_indices.empty:
            df_result.loc[max_breach_indices, 'Budget'] = df_result.loc[max_breach_indices, 'MaxShare'] * total_budget
        
        # Перерозподіл залишку
        current_total_budget = df_result['Budget'].sum()
        remaining_budget = total_budget - current_total_budget
        
        # Визначення інструментів, які ще не досягли своїх Min/Max
        free_indices = df_result[(df_result['BudgetSharePct'] > df_result['MinShare']) & (df_result['BudgetSharePct'] < df_result['MaxShare'])].index
        
        if len(free_indices) > 0 and abs(remaining_budget) > 0.01:
            free_ideal_shares = df_result.loc[free_indices, 'IdealShare']
            free_ideal_shares_norm = free_ideal_shares / free_ideal_shares.sum()
            df_result.loc[free_indices, 'Budget'] += free_ideal_shares_norm * remaining_budget
        else:
            break

    # Фінальний розрахунок метрик
    df_result["BudgetSharePct"] = df_result["Budget"] / total_budget
    df_result["Impressions"] = df_result["Budget"] / df_result["CPM"] * 1000
    df_result["Unique Reach (People)"] = df_result["Impressions"] / df_result["Freq"]
    df_result["ReachPct"] = df_result["Unique Reach (People)"] / total_audience
    df_result["ReachPct"] = df_result["ReachPct"].clip(upper=1.0)
    
    total_reach_prob = total_reach(df_result["Budget"].values, df_result)
    total_reach_people = total_reach_prob * total_audience

    st.subheader("Оптимальний спліт з плавним зниженням")
    st.write(f"Total Reach: **{total_reach_prob*100:.2f}%**")
    
    display_cols = ["Instrument", "CPM", "Freq", "MinShare", "MaxShare", "Budget", "BudgetSharePct", "Impressions", "Unique Reach (People)", "ReachPct"]
    st.dataframe(df_result[display_cols])

    # ==== Завантаження результатів у Excel ====
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Плавний спліт"
    
    for r in dataframe_to_rows(df_result[display_cols], index=False, header=True):
        ws.append(r)
    
    ws.append([])
    ws.append(["TOTAL", "", "", "", "", df_result["Budget"].sum(), 1.0, df_result["Impressions"].sum(), df_result["Unique Reach (People)"].sum(), f"{total_reach_prob*100:.2f}%"])
    ws.append(["TOTAL PEOPLE", "", "", "", "", "", "", "", int(total_reach_people), ""])
    
    for row in ws.iter_rows(min_row=2, max_row=1+len(df_result), min_col=7, max_col=7):
        for cell in row:
            cell.number_format = '0.00%'
            
    chart = BarChart()
    chart.type = "col"
    chart.title = "Бюджет по інструментам (%)"
    chart.y_axis.title = "Budget Share (%)"
    chart.x_axis.title = "Інструменти"
    data = Reference(ws, min_col=7, min_row=1, max_row=1+len(df_result))
    categories = Reference(ws, min_col=1, min_row=2, max_row=1+len(df_result))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    ws.add_chart(chart, "L2")

    wb.save(output)
    output.seek(0)
    
    st.download_button(
        label="Завантажити результати (Excel)",
        data=output,
        file_name="Digital_Split_Smooth.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
