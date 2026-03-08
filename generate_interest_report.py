import pandas as pd

# =========================
# 1. Load data
# =========================
df = pd.read_excel("borrowings_2025.xlsx")

# Strip any spaces from column headers
df.columns = df.columns.str.strip()

# Ensure Borrowing Date is datetime
df['Borrowing Date'] = pd.to_datetime(df['Borrowing Date'])

# =========================
# 2. Create Month column for Excel charts
# =========================
df['Month'] = df['Borrowing Date'].dt.strftime('%b-%Y')  # e.g., Jan-2025

# =========================
# 3. Pivot table: Interest per Bank per Month
# =========================
pivot = df.pivot_table(
    index='Month',          # Month as index for chart
    columns='Bank Name',
    values='Interest Amount',
    aggfunc='sum'
).fillna(0)

# Sort months chronologically
pivot = pivot.sort_index(key=lambda x: pd.to_datetime(x, format='%b-%Y'))

# =========================
# 4. Annual total per bank (for Pie Chart)
# =========================
annual_pivot = df.groupby('Bank Name')['Interest Amount'].sum()

# =========================
# 5. Create Excel with charts
# =========================
with pd.ExcelWriter("Interest_on_Borrowings_2025.xlsx", engine='xlsxwriter') as writer:

    # --- Raw Data ---
    df.to_excel(writer, sheet_name='Raw Data', index=False)

    # --- Pivot Table ---
    pivot.to_excel(writer, sheet_name='Bank-wise Summary', index=True)

    workbook  = writer.book
    worksheet = writer.sheets['Bank-wise Summary']

    # ===== Line Chart: Bank-wise Trend =====
    line_chart = workbook.add_chart({'type': 'line'})
    for i, bank in enumerate(pivot.columns):
        line_chart.add_series({
            'name':       bank,
            'categories': ['Bank-wise Summary', 1, 0, len(pivot), 0],   # Month column
            'values':     ['Bank-wise Summary', 1, i+1, len(pivot), i+1],  # Bank column
        })
    line_chart.set_title({'name': 'Bank-wise Interest Trend 2025'})
    line_chart.set_x_axis({'name': 'Month'})
    line_chart.set_y_axis({'name': 'Interest Amount'})
    line_chart.set_style(10)
    worksheet.insert_chart('H2', line_chart)

    # ===== Stacked Column Chart: Monthly Bank Contribution =====
    col_chart = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
    for i, bank in enumerate(pivot.columns):
        col_chart.add_series({
            'name': bank,
            'categories': ['Bank-wise Summary', 1, 0, len(pivot), 0],
            'values': ['Bank-wise Summary', 1, i+1, len(pivot), i+1],
        })
    col_chart.set_title({'name': 'Monthly Bank-wise Interest Contribution'})
    col_chart.set_x_axis({'name': 'Month'})
    col_chart.set_y_axis({'name': 'Interest Amount'})
    col_chart.set_style(11)
    worksheet.insert_chart('H18', col_chart)

    # ===== Pie Chart: Annual Share per Bank =====
    pie_chart = workbook.add_chart({'type': 'pie'})
    pie_chart.add_series({
        'name': 'Annual Interest Share per Bank',
        'categories': ['Bank-wise Summary', 0, 1, 0, len(annual_pivot)],
        'values': ['Bank-wise Summary', 1, 1, 1, len(annual_pivot)],
        'data_labels': {'percentage': True}
    })
    pie_chart.set_title({'name': 'Annual Interest Share per Bank'})
    worksheet.insert_chart('H34', pie_chart)

print("Excel report created: Interest_on_Borrowings_2025.xlsx")