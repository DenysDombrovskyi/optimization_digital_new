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
st.title("Digital Split Optimizer – 3 варіанти")

# ==== Налаштування ====
num_instruments = st.number_input("Кількість інструментів:", min_value=1, max_value=50, value=20, step=1)
total_audience = st.number_input("Загальний розмір потенційної аудиторії:", value=50000, step=1000)
mode = st.radio("Режим оптимізації:", ["Оцінка спліту (при заданому бюджеті)", "Мінімальна вартість (при заданому охопленні)"])

if mode == "Оцінка спліту (при заданому бюджеті)":
    total_budget = st.number_input("Заданий бюджет ($):", value=100000, step=1000)
else:
    total_budget = st.number_input("Максимальний бюджет ($):", value=100000, step=1000)
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
    if mode == "Оцінка спліту (при заданому бюджеті)":
        
        eff_weights = df["Efficiency"].values / df["Efficiency"].sum()
        x0 = df["MinShare"].values * total_budget + \
             (df["MaxShare"].values - df["MinShare"].values) * total_budget * eff_weights
        bounds = [(df.loc[i, "MinShare"] * total_budget, df.loc[i, "MaxShare"] * total_budget) for i in range(len(df))]
        
        constraints = [{'type': 'eq', 'fun': lambda budgets: np.sum(budgets) - total_budget}]
        
    else: # Min Cost (при охопленні)
        cpr_weights = (1 / df["CPR"]).values
        cpr_weights_normalized = cpr_weights / cpr_weights.sum()
        
        x0 = cpr_weights_normalized * total_budget
        
        min_budgets = df["MinShare"].values * total_budget
        max_budgets = df["MaxShare"].values * total_budget
        bounds = list(zip(min_budgets, max_budgets))

    def total_reach(budgets):
        impressions = budgets / df["CPM"].values * 1000
        reach_i = np.clip(impressions / total_audience, 0, 1)
        return 1 - np.prod(1 - reach_i)
    
    # --- Нові функції для 3 варіантів ---
    def objective_cheapest(budgets):
        # Мінімізуємо середньозважений CPM
        total_b = np.sum(budgets)
        if total_b == 0:
            return float('inf')
        
        weighted_cpm = np.sum(budgets * df["CPM"].values) / total_b
        return weighted_cpm
    
    def objective_best_reach(budgets):
        # Максимізуємо охоплення
        return -total_reach(budgets)
        
    def objective_balanced(budgets):
        # Компроміс: максимізуємо охоплення і винагороджуємо за ефективність
        total_b = np.sum(budgets)
        if total_b == 0:
            return float('inf')
            
        reach = total_reach(budgets)
        budget_shares = budgets / total_b
        eff_weights = df["Efficiency"].values / df["Efficiency"].sum()
        
        # Нагорода за розподіл, близький до ідеального за ефективністю
        efficiency_reward = np.sum(budget_shares * eff_weights)
        return - (reach + 0.5 * efficiency_reward)

    if mode == "Оцінка спліту (при заданому бюджеті)":
        
        st.subheader("Найдешевший спліт")
        res_cheapest = minimize(objective_cheapest, x0=x0, bounds=bounds, constraints=constraints, method='SLSQP')
        df_cheapest = df.copy()
        if res_cheapest.success:
            df_cheapest["Budget"] = res_cheapest.x
        else:
            df_cheapest["Budget"] = 0.0
        
        df_cheapest["BudgetSharePct"] = df_cheapest["Budget"] / df_cheapest["Budget"].sum()
        df_cheapest["Impressions"] = df_cheapest["Budget"] / df_cheapest["CPM"] * 1000
        df_cheapest["ReachPct"] = df_cheapest["Impressions"] / total_audience
        df_cheapest["ReachPct"] = df_cheapest["ReachPct"].clip(upper=1.0)
        total_reach_prob_cheapest = 1 - np.prod(1 - df_cheapest["ReachPct"])
        st.write(f"Total Reach: {total_reach_prob_cheapest*100:.2f}%")
        st.dataframe(df_cheapest)
        
        st.subheader("Найкраще охоплення")
        res_best_reach = minimize(objective_best_reach, x0=x0, bounds=bounds, constraints=constraints, method='SLSQP')
        df_best_reach = df.copy()
        if res_best_reach.success:
            df_best_reach["Budget"] = res_best_reach.x
        else:
            df_best_reach["Budget"] = 0.0
        
        df_best_reach["BudgetSharePct"] = df_best_reach["Budget"] / df_best_reach["Budget"].sum()
        df_best_reach["Impressions"] = df_best_reach["Budget"] / df_best_reach["CPM"] * 1000
        df_best_reach["ReachPct"] = df_best_reach["Impressions"] / total_audience
        df_best_reach["ReachPct"] = df_best_reach["ReachPct"].clip(upper=1.0)
        total_reach_prob_best_reach = 1 - np.prod(1 - df_best_reach["ReachPct"])
        st.write(f"Total Reach: {total_reach_prob_best_reach*100:.2f}%")
        st.dataframe(df_best_reach)
        
        st.subheader("Збалансований спліт")
        res_balanced = minimize(objective_balanced, x0=x0, bounds=bounds, constraints=constraints, method='SLSQP')
        df_balanced = df.copy()
        if res_balanced.success:
            df_balanced["Budget"] = res_balanced.x
        else:
            df_balanced["Budget"] = 0.0
        
        df_balanced["BudgetSharePct"] = df_balanced["Budget"] / df_balanced["Budget"].sum()
        df_balanced["Impressions"] = df_balanced["Budget"] / df_balanced["CPM"] * 1000
        df_balanced["ReachPct"] = df_balanced["Impressions"] / total_audience
        df_balanced["ReachPct"] = df_balanced["ReachPct"].clip(upper=1.0)
        total_reach_prob_balanced = 1 - np.prod(1 - df_balanced["ReachPct"])
        st.write(f"Total Reach: {total_reach_prob_balanced*100:.2f}%")
        st.dataframe(df_balanced)
        
    else:  # Min Cost (при охопленні)
        def objective_min_budget_with_balance(budgets):
            total_b = np.sum(budgets)
            if total_b == 0:
                return float('inf')
            
            budget_shares = budgets / total_b
            cpr_weights_normalized = (1 / df["CPR"]).values / (1 / df["CPR"]).sum()
            balance_penalty = np.sum((budget_shares - cpr_weights_normalized)**2)
            
            return total_b + lambda_balance * balance_penalty * total_budget
        
        def constraint_reach(budgets):
            return total_reach(budgets) - target_reach
            
        cons = [{'type': 'ineq', 'fun': constraint_reach}]
        for i in range(len(df)):
            max_share = df.loc[i, "MaxShare"]
            def constraint_max_share(budgets, i=i, max_share=max_share):
                if np.sum(budgets) == 0:
                    return 1.0
                return max_share - (budgets[i] / np.sum(budgets))
            
            cons.append({'type': 'ineq', 'fun': constraint_max_share})

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
