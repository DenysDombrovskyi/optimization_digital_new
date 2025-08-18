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
    df = df.sort_values(by="CPR", ascending=True).reset_index(drop=True)
    
    x0 = df["MinShare"].values * total_budget + \
         (df["MaxShare"].values - df["MinShare"].values) * total_budget / len(df)
    bounds = [(df.loc[i, "MinShare"] * total_budget, df.loc[i, "MaxShare"] * total_budget) for i in range(len(df))]
    
    def total_reach(budgets, df_current):
        impressions = budgets / df_current["CPM"].values * 1000
        reach_i = np.clip(impressions / total_audience, 0, 1)
        return 1 - np.prod(1 - reach_i)

    def calculate_results(df_to_calc, res, total_budget_calc=None):
        if res.success:
            df_to_calc["Budget"] = res.x
        else:
            df_to_calc["Budget"] = 0.0
        
        if total_budget_calc is None:
            total_budget_calc = df_to_calc["Budget"].sum()
        
        df_to_calc["BudgetSharePct"] = df_to_calc["Budget"] / total_budget_calc
        df_to_calc["Impressions"] = df_to_calc["Budget"] / df_to_calc["CPM"] * 1000
        df_to_calc["ReachPct"] = df_to_calc["Impressions"] / total_audience
        df_to_calc["ReachPct"] = df_to_calc["ReachPct"].clip(upper=1.0)
        df_to_calc["Unique Reach (People)"] = df_to_calc["Impressions"] / df_to_calc["Freq"]
        total_reach_prob = total_reach(df_to_calc["Budget"].values, df_to_calc)
        
        return df_to_calc, total_reach_prob

    results = {}
    display_cols = ["Instrument", "CPM", "Freq", "MinShare", "MaxShare", "Efficiency", "CPR", "Budget", "BudgetSharePct", "Impressions", "Unique Reach (People)", "ReachPct"]
    
    # ==== Варіант 1: Найдешевший спліт (залежність від CPM/CPR) ====
    st.subheader("1. Найдешевший спліт (частки залежать від ефективності)")
    
    def objective_cheapest(budgets):
        total_b = np.sum(budgets)
        if total_b == 0:
            return float('inf')
        
        # Мінімізуємо середньозважений CPM/CPR
        # Це прямо змушує оптимізатор віддавати перевагу найдешевшим інструментам
        weighted_cpm = np.sum(budgets * df["CPM"].values) / total_b
        return weighted_cpm
    
    res_cheapest = minimize(objective_cheapest, x0=x0, bounds=bounds, constraints={'type': 'eq', 'fun': lambda b: np.sum(b) - total_budget}, method='SLSQP')
    df_cheapest, reach_cheapest = calculate_results(df.copy(), res_cheapest, total_budget)
    st.write(f"Total Reach: **{reach_cheapest*100:.2f}%**")
    st.dataframe(df_cheapest[display_cols])
    results['Найдешевший спліт'] = df_cheapest
    
    # ==== Варіант 2: Спліт з відсіканням (мінімум 2 інструменти) ====
    st.subheader("2. Спліт з відсіканням (максимальне охоплення)")
    df_temp = df.copy()
    num_instruments_to_keep = len(df_temp)
    best_reach = 0
    best_df = pd.DataFrame()

    while num_instruments_to_keep >= 2:
        df_temp = df_temp.sort_values(by="Efficiency", ascending=False)
        df_temp = df_temp.head(num_instruments_to_keep)
        
        temp_x0 = df_temp["MinShare"].values * total_budget + (df_temp["MaxShare"].values - df_temp["MinShare"].values) * total_budget / num_instruments_to_keep
        temp_bounds = [(df_temp.loc[i, "MinShare"] * total_budget, df_temp.loc[i, "MaxShare"] * total_budget) for i in df_temp.index]
        temp_constraints = [{'type': 'eq', 'fun': lambda budgets: np.sum(budgets) - total_budget}]
        
        def objective_reach_prune(budgets):
            return -total_reach(budgets, df_temp)
            
        res = minimize(objective_reach_prune, x0=temp_x0, bounds=temp_bounds, constraints=temp_constraints, method='SLSQP')
        
        if res.success:
            current_reach = total_reach(res.x, df_temp)
            if current_reach > best_reach:
                best_reach = current_reach
                best_df = df_temp.copy()
                best_df["Budget"] = res.x
                
        num_instruments_to_keep -= 1
        
    if not best_df.empty:
        df_pruned = best_df
        df_pruned, reach_pruned = calculate_results(df_pruned.copy(), res, total_budget)
        st.write(f"Total Reach: **{reach_pruned*100:.2f}%**")
        st.dataframe(df_pruned[display_cols])
        results['Спліт з відсіканням'] = df_pruned
    else:
        st.error("Не вдалося знайти оптимальний спліт з відсіканням.")

    # ==== Варіант 3: Максимальне охоплення (ідеальний спліт) ====
    st.subheader("3. Максимальне охоплення (ідеальний спліт)")
    def objective_max_reach_unconstrained(budgets):
        return -total_reach(budgets, df)
    
    x0_ideal = df["MinShare"].values * 10000 
    bounds_ideal = [(df.loc[i, "MinShare"] * total_budget, None) for i in range(len(df))]
    
    res_ideal = minimize(objective_max_reach_unconstrained, x0=x0_ideal, bounds=bounds_ideal, method='SLSQP')
    
    df_ideal = df.copy()
    if res_ideal.success:
        df_ideal["Budget"] = res_ideal.x
    else:
        df_ideal["Budget"] = 0.0
        
    df_ideal["BudgetSharePct"] = df_ideal["Budget"] / df_ideal["Budget"].sum()
    df_ideal["Impressions"] = df_ideal["Budget"] / df_ideal["CPM"] * 1000
    df_ideal["ReachPct"] = df_ideal["Impressions"] / total_audience
    df_ideal["ReachPct"] = df_ideal["ReachPct"].clip(upper=1.0)
    df_ideal["Unique Reach (People)"] = df_ideal["Impressions"] / df_ideal["Freq"]
    total_reach_prob_ideal = total_reach(df_ideal["Budget"].values, df_ideal)
    st.write(f"Total Reach: **{total_reach_prob_ideal*100:.2f}%**")
    st.write(f"**Бюджет для досягнення ідеального охоплення:** **{df_ideal['Budget'].sum():.2f}$**")
    st.dataframe(df_ideal[display_cols])
    results['Максимальне охоплення'] = df_ideal

    # ==== Завантаження результатів у Excel ====
    output = io.BytesIO()
    wb = Workbook()

    for sheet_name, result_df in results.items():
        ws = wb.create_sheet(title=sheet_name)
        for r in dataframe_to_rows(result_df[display_cols], index=False, header=True):
            ws.append(r)
        
        total_budget_sum = result_df["Budget"].sum()
        total_impressions_sum = result_df["Impressions"].sum()
        total_reach_prob = total_reach(result_df["Budget"].values, result_df)
        total_reach_people = total_reach_prob * total_audience
        
        ws.append([])
        ws.append(["TOTAL", "", "", "", "", "", "", total_budget_sum, 1.0, total_impressions_sum, np.sum(result_df["Unique Reach (People)"]), f"{total_reach_prob*100:.2f}%"])
        ws.append(["TOTAL PEOPLE", "", "", "", "", "", "", "", "", "", int(total_reach_people)])
        
        for row in ws.iter_rows(min_row=2, max_row=1+len(result_df), min_col=9, max_col=9):
            for cell in row:
                cell.number_format = '0.00%'
                
        chart = BarChart()
        chart.type = "col"
        chart.title = f"Бюджет по інструментам (%) - {sheet_name}"
        chart.y_axis.title = "Budget Share (%)"
        chart.x_axis.title = "Інструменти"
        data = Reference(ws, min_col=9, min_row=1, max_row=1+len(result_df))
        categories = Reference(ws, min_col=1, min_row=2, max_row=1+len(result_df))
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(categories)
        ws.add_chart(chart, "L2")

    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    wb.save(output)
    output.seek(0)
    
    st.download_button(
        label="Завантажити результати (Excel)",
        data=output,
        file_name="Digital_Split_3_Options.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

