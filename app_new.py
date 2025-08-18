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
        # MinShare and MaxShare are now always required
        df.loc[i, "MinShare"] = st.number_input(f"Min Share {i+1}", value=float(df.loc[i, "MinShare"]), step=0.01)
        df.loc[i, "MaxShare"] = st.number_input(f"Max Share {i+1}", value=float(df.loc[i, "MaxShare"]), step=0.01)
    submitted = st.form_submit_button("Перерахувати спліт")

if submitted:
    st.session_state.df = df.copy()
    
    df["Efficiency"] = 1000 / (df["CPM"] * df["Freq"])
    # Перевірка ділення на нуль
    df["CPR"] = df["CPM"] / 1000 * df["Freq"]
    df.loc[df["Freq"] == 0, "CPR"] = np.inf
    
    def total_reach(budgets, df_current):
        impressions = budgets / df_current["CPM"].values * 1000
        reach_i = np.clip(impressions / total_audience, 0, 1)
        return 1 - np.prod(1 - reach_i)

    # ==== Загальний алгоритм розподілу, який використовується в обох опціях ====
    def distribute_budget(df_input, budget):
        df_result = df_input.sort_values(by="CPR", ascending=True).reset_index(drop=True)

        # 1. Розрахунок мінімального бюджету для всіх
        df_result['Budget'] = df_result['MinShare'] * budget

        # 2. Виділяємо максимальний бюджет найдешевшому інструменту (за CPR)
        cheapest_instrument_index = df_result.index[0]
        df_result.loc[cheapest_instrument_index, 'Budget'] = df_result.loc[cheapest_instrument_index, 'MaxShare'] * budget
        
        # 3. Розрахунок доступного бюджету для розподілу між іншими інструментами
        remaining_budget = budget - df_result['Budget'].sum()

        # 4. Розподіл залишку на основі CPR для всіх, хто не досяг максимуму
        if remaining_budget > 0:
            indices_to_distribute_to = df_result.index[1:]
            
            indices_to_distribute_to = indices_to_distribute_to[df_result.loc[indices_to_distribute_to, 'Budget'] < df_result.loc[indices_to_distribute_to, 'MaxShare'] * budget]
            
            if len(indices_to_distribute_to) > 0:
                ideal_shares = 1 / df_result.loc[indices_to_distribute_to, 'CPR']
                ideal_shares_norm = ideal_shares / ideal_shares.sum()
                
                available_budget = (df_result.loc[indices_to_distribute_to, 'MaxShare'] * budget - df_result.loc[indices_to_distribute_to, 'Budget'])

                to_allocate = min(remaining_budget, available_budget.sum())
                
                if to_allocate > 0:
                    budget_share = ideal_shares_norm * to_allocate
                    
                    df_result.loc[indices_to_distribute_to, 'Budget'] += np.minimum(budget_share.values, available_budget.values)
                    
        # 5. Фінальне перезважування для отримання 100%
        df_result["BudgetSharePct"] = df_result["Budget"] / df_result["Budget"].sum() if df_result["Budget"].sum() > 0 else 0

        # Фінальний розрахунок інших метрик
        df_result["Impressions"] = df_result["Budget"] / df_result["CPM"] * 1000
        df_result["Unique Reach (People)"] = df_result["Impressions"] / df_result["Freq"]
        df_result["ReachPct"] = df_result["Unique Reach (People)"] / total_audience
        df_result["ReachPct"] = df_result["ReachPct"].clip(upper=1.0)
        
        return df_result
    
    # ==== Мінімізація бюджету ====
    if optimization_goal == 'Мінімізація бюджету':
        reach_target_people = total_audience * (reach_target_pct / 100)
        
        # Розраховуємо рекомендований бюджет для досягнення цілі
        # Використовуємо ітерацію для пошуку мінімального бюджету
        low_budget = 0
        high_budget = 1000000 # Велика початкова оцінка
        recommended_budget = 0
        
        for _ in range(100): # Використовуємо 100 ітерацій для точності
            mid_budget = (low_budget + high_budget) / 2
            if mid_budget == 0:
                mid_budget = 1
            
            temp_df = distribute_budget(df, mid_budget)
            current_reach = total_reach(temp_df['Budget'].values, temp_df) * total_audience
            
            if current_reach < reach_target_people:
                low_budget = mid_budget
            else:
                high_budget = mid_budget
                recommended_budget = mid_budget
        
        df_result = distribute_budget(df, recommended_budget)
        final_total_budget = df_result['Budget'].sum()
        total_reach_prob = total_reach(df_result["Budget"].values, df_result)
        
        st.subheader(f"Результат: Мінімізація бюджету")
        st.write(f"Мінімальний бюджет для охоплення **{total_reach_prob*100:.2f}%** аудиторії: **{final_total_budget:,.2f} $**")
        st.dataframe(df_result)
        
    # ==== Максимізація охоплення ====
    elif optimization_goal == 'Максимізація охоплення':
        df_result = distribute_budget(df, total_budget)
        
        total_reach_prob = total_reach(df_result["Budget"].values, df_result)
        
        st.subheader("Гібридний спліт (Максимізація охоплення)")
        st.write(f"Total Reach: **{total_reach_prob*100:.2f}%**")
        
        display_cols = ["Instrument", "CPM", "CPR", "Freq", "MinShare", "MaxShare", "Budget", "BudgetSharePct", "Impressions", "Unique Reach (People)", "ReachPct"]
        st.dataframe(df_result[display_cols])

    # ==== Завантаження результатів у Excel ====
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Гібридний спліт"
    
    excel_cols = ["Instrument", "CPM", "CPR", "Freq", "MinShare", "MaxShare", "Budget", "BudgetSharePct", "Impressions", "Unique Reach (People)", "ReachPct"]

    df_to_save = df_result[excel_cols].copy()
    final_total_budget = df_result["Budget"].sum()
    final_total_reach_prob = total_reach(df_result["Budget"].values, df_result)
    final_total_reach_people = final_total_reach_prob * total_audience

    df_to_save.loc[len(df_to_save)] = ["TOTAL", "", "", "", df_result['MinShare'].sum(), df_result['MaxShare'].sum(), final_total_budget, 1.0, df_result["Impressions"].sum(), df_result["Unique Reach (People)"].sum(), f"{final_total_reach_prob*100:.2f}%"]

    for r in dataframe_to_rows(df_to_save, index=False, header=True):
        ws.append(r)
    
    # Format percentage column
    budget_share_pct_col = excel_cols.index("BudgetSharePct") + 1
    for row in ws.iter_rows(min_row=2, max_row=1+len(df_result), min_col=budget_share_pct_col, max_col=budget_share_pct_col):
        for cell in row:
            cell.number_format = '0.00%'
            
    chart = BarChart()
    chart.type = "col"
    chart.title = "Бюджет по інструментам (%)"
    chart.y_axis.title = "Budget Share (%)"
    chart.x_axis.title = "Інструменти"
    data = Reference(ws, min_col=budget_share_pct_col, min_row=1, max_row=1+len(df_result))
    categories = Reference(ws, min_col=1, min_row=2, max_row=1+len(df_result))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    ws.add_chart(chart, f"L{budget_share_pct_col}")

    wb.save(output)
    output.seek(0)
    
    st.download_button(
        label="Завантажити результати (Excel)",
        data=output,
        file_name="Digital_Split_Hybrid.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
