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
st.title("Digital Split Optimizer – Max Reach + Min Cost оновлений")

# ==== Налаштування ====
num_instruments = st.number_input("Кількість інструментів:", min_value=1, max_value=50, value=20, step=1)
total_audience = st.number_input("Загальний розмір потенційної аудиторії:", value=50000, step=1000)
mode = st.radio("Режим оптимізації:", ["Max Reach (при бюджеті)", "Min Cost (при охопленні)"])
total_budget = st.number_input("Максимальний бюджет ($):", value=100000, step=1000)
if mode == "Min Cost (при охопленні)":
    target_reach = st.number_input("Цільове охоплення (0-1):", value=0.8, step=0.01)
    lambda_balance = st.slider("Коефіцієнт балансу (0 - менш збалансований, 1 - більш збалансований):",
                               min_value=0.0, max_value=1.0, value=0.1, step=0.01)

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
    
    # Розрахунок ефективності та CPR
    df["Efficiency"] = 1000 / (df["CPM"] * df["Freq"])
    df["CPR"] = (df["CPM"] / 1000) * (df["Freq"] * total_audience) / total_audience
    df = df.sort_values(by="CPR", ascending=True).reset_index(drop=True)
    
    # Визначення початкового значення x0 та меж
    if mode == "Max Reach (при бюджеті)":
        eff_weights = df["Efficiency"].values / df["Efficiency"].sum()
        x0 = df["MinShare"].values * total_budget + \
             (df["MaxShare"].values - df["MinShare"].values) * total_budget * eff_weights
        bounds = [(df.loc[i, "MinShare"] * total_budget, df.loc[i, "MaxShare"] * total_budget) for i in range(len(df))]
        
    else: # Min Cost (при охопленні)
        # Створюємо початкове значення, де бюджет розподіляється за рангом CPR
        cpr_weights = (1 / df["CPR"]).values
        cpr_weights_normalized = cpr_weights / cpr_weights.sum()
        
        # Початкове значення x0, що ігнорує Min/MaxShare, щоб дати оптимізатору сильний сигнал
        x0 = cpr_weights_normalized * total_budget
        
        # Межі залишаються на основі MinShare та MaxShare
        min_budgets = df["MinShare"].values * total_budget
        max_budgets = df["MaxShare"].values * total_budget
        bounds = list(zip(min_budgets, max_budgets))

    def total_reach(budgets):
        impressions = budgets / df["CPM"].values * 1000
        reach_i = np.clip(impressions / total_audience, 0, 1)
        return 1 - np.prod(1 - reach_i)

    if mode == "Max Reach (при бюджеті)":
        lambda_balance_max_reach = 0.2
        def objective_maxreach(budgets):
            eff_weights = df["Efficiency"].values / df["Efficiency"].sum()
            balance_penalty = np.sum((budgets / np.sum(budgets) - eff_weights)**2)
            return - (total_reach(budgets) - lambda_balance_max_reach * balance_penalty)

        res = minimize(objective_maxreach, x0=x0, bounds=bounds, method='SLSQP')

    else:  # Min Cost (при охопленні)
        def objective_min_budget_with_balance(budgets):
            total_b = np.sum(budgets)
            if total_b == 0:
                return float('inf')
            
            # Розрахунок штрафу за дисбаланс, базуючись на CPR
            budget_shares = budgets / total_b
            cpr_weights_normalized = (1 / df["CPR"]).values / (1 / df["CPR"]).sum()
            balance_penalty = np.sum((budget_shares - cpr_weights_normalized)**2)
            
            # Оновлена цільова функція: мінімізуємо суму бюджетів + штраф
            return total_b + lambda_balance * total_b * balance_penalty
        
        def constraint_reach(budgets):
            return total_reach(budgets) - target_reach

        cons = [{'type': 'ineq', 'fun': constraint_reach}]
        
        res = minimize(objective_min_budget_with_balance, x0=x0, bounds=bounds, constraints=cons, method='SLSQP', options={'disp': False})

    if res.success:
        df["Budget"] = res.x
        df["Impressions"] = df["Budget"] / df["CPM"] * 1000
    else:
        st.error("Не вдалося знайти оптимальний спліт. Спробуйте змінити початкові дані (наприклад, цільове охоплення чи частки інструментів).")
        df["Budget"] = 0.0
        df["Impressions"] = 0.0

    df["BudgetSharePct"] = df["Budget"] / df["Budget"].sum()
    df["ReachPct"] = df["Impressions"] / total_audience
    df["ReachPct"] = df["ReachPct"].clip(upper=1.0)
    total_reach_prob = 1 - np.prod(1 - df["ReachPct"])
    total_reach_people = total_reach_prob * total_audience

    display_cols = ["Instrument", "CPM", "Freq", "MinShare", "MaxShare", "Efficiency", "CPR", "Budget", "BudgetSharePct", "Impressions", "ReachPct"]
    st.subheader(f"Розрахунок спліту (Total Audience={total_audience})")
    st.dataframe(df[display_cols])
    st.write(f"Total Reach (ймовірність): {total_reach_prob*100:.2f}%")
    st.write(f"Total Reach (людей): {int(total_reach_people)}")

    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button("Завантажити CSV", data=csv, file_name="Digital_Split_Result.csv", mime="text/csv")

    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Digital Split"
    for r in dataframe_to_rows(df[display_cols], index=False, header=True):
        ws.append(r)

    total_budget_sum = df["Budget"].sum()
    total_impressions_sum = df["Impressions"].sum()
    ws.append([])
    ws.append(["TOTAL", "", "", "", "", "", "", total_budget_sum, 1.0, total_impressions_sum, f"{total_reach_prob*100:.2f}%"])
    ws.append(["TOTAL PEOPLE", "", "", "", "", "", "", "", int(total_reach_people)])

    for row in ws.iter_rows(min_row=2, max_row=1+len(df), min_col=9, max_col=9):
        for cell in row:
            cell.number_format = '0.00%'

    chart = BarChart()
    chart.type = "col"
    chart.title = "Бюджет по інструментам (%)"
    chart.y_axis.title = "Budget Share (%)"
    chart.x_axis.title = "Інструменти"
    data = Reference(ws, min_col=9, min_row=1, max_row=1+len(df))
    categories = Reference(ws, min_col=1, min_row=2, max_row=1+len(df))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    chart.shape = 4
    ws.add_chart(chart, "L2")

    wb.save(output)
    output.seek(0)

    st.download_button(
        label="Завантажити Excel з графіком",
        data=output,
        file_name="Digital_Split_Result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
