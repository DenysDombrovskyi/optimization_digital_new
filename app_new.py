import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import numbers

st.set_page_config(page_title="Digital Split Optimizer", layout="wide")
st.title("Digital Split Optimizer – Гібридний спліт")

# ==== Налаштування ====
optimization_goal = st.radio(
    "Оберіть мету оптимізації:",
    ('Максимізація охоплення', 'Мінімізація бюджету'),
    horizontal=True
)
st.markdown("---")

num_instruments = st.number_input("Кількість інструментів:", min_value=1, max_value=50, value=25, step=1)
total_audience = st.number_input("Загальний розмір потенційної аудиторії:", value=50000, step=1000)

if optimization_goal == 'Максимізація охоплення':
    total_budget = st.number_input("Заданий бюджет ($):", value=100000, step=1000)
else:
    reach_target_pct = st.number_input("Бажаний відсоток охоплення (%):", min_value=1, max_value=100, value=50, step=1)
    
# ==== Початкові дані ====
if "df" not in st.session_state or len(st.session_state.df) != num_instruments:
    st.session_state.df = pd.DataFrame({
        "Instrument": [f"Instrument {i+1}" for i in range(num_instruments)],
        "CPM": [50 + i*2 for i in range(num_instruments)],
        "Freq": [1.0 + 0.05*i for i in range(num_instruments)],
        "MinShare": [0.01 for _ in range(num_instruments)],
        "MaxShare": [0.15 for i in range(num_instruments)]
    })

df = st.session_state.df.copy()

# ==== Форма редагування ====
with st.form("instrument_form"):
    for i in range(len(df)):
        st.subheader(f"Інструмент {i+1}")
        df.loc[i, "Instrument"] = st.text_input(f"Назва {i+1}", value=df.loc[i, "Instrument"])
        df.loc[i, "CPM"] = st.number_input(f"CPM {i+1}", value=float(df.loc[i, "CPM"]), step=1.0)
        df.loc[i, "Freq"] = st.number_input(f"Frequency {i+1}", value=float(df.loc[i, "Freq"]), step=0.01)
        if optimization_goal == 'Максимізація охоплення':
            df.loc[i, "MinShare"] = st.number_input(f"Min Share {i+1}", value=float(df.loc[i, "MinShare"]), step=0.01)
            df.loc[i, "MaxShare"] = st.number_input(f"Max Share {i+1}", value=float(df.loc[i, "MaxShare"]), step=0.01)
    submitted = st.form_submit_button("Перерахувати спліт")

if submitted:
    st.session_state.df = df.copy()
    
    df["Efficiency"] = 1000 / (df["CPM"] * df["Freq"])
    df["CPR"] = (df["CPM"] / 1000) * (df["Freq"] * total_audience) / total_audience
    
    def total_reach(budgets, df_current):
        impressions = budgets / df_current["CPM"].values * 1000
        reach_i = np.clip(impressions / total_audience, 0, 1)
        return 1 - np.prod(1 - reach_i)

    # ==== Мінімізація бюджету (новий варіант) ====
    if optimization_goal == 'Мінімізація бюджету':
        df_result = df.copy()
        df_result = df_result.sort_values(by="CPR", ascending=True).reset_index(drop=True)
        
        reach_target_people = total_audience * (reach_target_pct / 100)
        current_reach_people = 0
        current_budget = 0
        
        df_result['Budget'] = 0.0
        
        for index, row in df_result.iterrows():
            unique_reach_instrument = row["Impressions"] / row["Freq"] if row["Freq"] > 0 else 0
            
            # Якщо інструмент може забезпечити все, що залишилось
            if current_reach_people + unique_reach_instrument >= reach_target_people:
                needed_reach = reach_target_people - current_reach_people
                impressions_needed = needed_reach * row["Freq"]
                budget_needed = impressions_needed * row["CPM"] / 1000
                
                df_result.loc[index, 'Budget'] = budget_needed
                current_budget += budget_needed
                current_reach_people = reach_target_people
                break
            else:
                # Використовуємо інструмент повністю
                budget_instrument = row['Budget']
                df_result.loc[index, 'Budget'] = budget_instrument
                current_budget += budget_instrument
                current_reach_people += unique_reach_instrument
        
        # Перераховуємо всі частки на основі отриманого мінімального бюджету
        df_result["BudgetSharePct"] = df_result["Budget"] / current_budget if current_budget > 0 else 0
        df_result["Impressions"] = df_result["Budget"] / df_result["CPM"] * 1000
        df_result["Unique Reach (People)"] = df_result["Impressions"] / df_result["Freq"]
        df_result["ReachPct"] = df_result["Unique Reach (People)"] / total_audience
        
        st.subheader(f"Результат: Мінімізація бюджету")
        st.write(f"Мінімальний бюджет для охоплення **{reach_target_pct}%** аудиторії: **{current_budget:,.2f} $**")
        st.dataframe(df_result)

    # ==== Максимізація охоплення (існуючий варіант) ====
    elif optimization_goal == 'Максимізація охоплення':
        df = df.sort_values(by="CPR", ascending=True).reset_index(drop=True)
        df_result = df.copy()
        
        # 1. Розрахунок мінімального бюджету для всіх
        df_result['Budget'] = df_result['MinShare'] * total_budget

        # 2. Виділяємо максимальний бюджет найдешевшому інструменту (за CPR)
        cheapest_instrument_index = df_result.index[0]
        df_result.loc[cheapest_instrument_index, 'Budget'] = df_result.loc[cheapest_instrument_index, 'MaxShare'] * total_budget
        
        # 3. Розрахунок доступного бюджету для розподілу між іншими інструментами
        remaining_budget = total_budget - df_result['Budget'].sum()

        # 4. Розподіл залишку на основі CPR для всіх, хто не досяг максимуму
        if remaining_budget > 0:
            indices_to_distribute_to = df_result.index[1:]
            
            indices_to_distribute_to = indices_to_distribute_to[df_result.loc[indices_to_distribute_to, 'Budget'] < df_result.loc[indices_to_distribute_to, 'MaxShare'] * total_budget]
            
            if len(indices_to_distribute_to) > 0:
                ideal_shares = 1 / df_result.loc[indices_to_distribute_to, 'CPR']
                ideal_shares_norm = ideal_shares / ideal_shares.sum()
                
                available_budget = (df_result.loc[indices_to_distribute_to, 'MaxShare'] * total_budget - df_result.loc[indices_to_distribute_to, 'Budget'])

                to_allocate = min(remaining_budget, available_budget.sum())
                
                if to_allocate > 0:
                    budget_share = ideal_shares_norm * to_allocate
                    
                    df_result.loc[indices_to_distribute_to, 'Budget'] += np.minimum(budget_share.values, available_budget.values)
                    
        # 5. Фінальне перезважування для отримання 100%
        df_result["BudgetSharePct"] = df_result["Budget"] / df_result["Budget"].sum()

        # Фінальний розрахунок інших метрик
        df_result["Impressions"] = df_result["Budget"] / df_result["CPM"] * 1000
        df_result["Unique Reach (People)"] = df_result["Impressions"] / df_result["Freq"]
        df_result["ReachPct"] = df_result["Unique Reach (People)"] / total_audience
        df_result["ReachPct"] = df_result["ReachPct"].clip(upper=1.0)
        
        total_reach_prob = total_reach(df_result["Budget"].values, df_result)
        total_reach_people = total_reach_prob * total_audience

        st.subheader("Гібридний спліт (Максимізація охоплення)")
        st.write(f"Total Reach: **{total_reach_prob*100:.2f}%**")
        
        display_cols = ["Instrument", "CPM", "CPR", "Freq", "MinShare", "MaxShare", "Budget", "BudgetSharePct", "Impressions", "Unique Reach (People)", "ReachPct"]
        st.dataframe(df_result[display_cols])

    # ==== Завантаження результатів у Excel ====
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Гібридний спліт"
    
    # Визначаємо, які колонки виводити в Excel
    if optimization_goal == 'Мінімізація бюджету':
        excel_cols = ["Instrument", "CPM", "CPR", "Freq", "Budget", "BudgetSharePct", "Impressions", "Unique Reach (People)", "ReachPct"]
    else:
        excel_cols = ["Instrument", "CPM", "CPR", "Freq", "MinShare", "MaxShare", "Budget", "BudgetSharePct", "Impressions", "Unique Reach (People)", "ReachPct"]
        
    for r in dataframe_to_rows(df_result[excel_cols], index=False, header=True):
        ws.append(r)
    
    ws.append([])
    ws.append(["TOTAL", "", "", "", "", df_result["Budget"].sum(), 1.0, df_result["Impressions"].sum(), df_result["Unique Reach (People)"].sum(), f"{total_reach_prob*100:.2f}%"])
    ws.append(["TOTAL PEOPLE", "", "", "", "", "", "", "", int(total_reach_people), ""])
    
    for row in ws.iter_rows(min_row=2, max_row=1+len(df_result), min_col=8, max_col=8):
        for cell in row:
            cell.number_format = '0.00%'
            
    chart = BarChart()
    chart.type = "col"
    chart.title = "Бюджет по інструментам (%)"
    chart.y_axis.title = "Budget Share (%)"
    chart.x_axis.title = "Інструменти"
    data = Reference(ws, min_col=8, min_row=1, max_row=1+len(df_result))
    categories = Reference(ws, min_col=1, min_row=2, max_row=1+len(df_result))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    ws.add_chart(chart, "L2")

    wb.save(output)
    output.seek(0)
    
    st.download_button(
        label="Завантажити результати (Excel)",
        data=output,
        file_name="Digital_Split_Hybrid.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
