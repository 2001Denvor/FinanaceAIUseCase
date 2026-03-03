import pandas as pd
from pptx import Presentation
from pptx.util import Inches
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

# -----------------------------
# Helper function to clean dataframe columns
# -----------------------------
def clean_columns(df):
    df.columns = df.columns.str.strip().str.lower()
    return df

# -----------------------------
# Load CSV files safely
# -----------------------------
pnl = clean_columns(pd.read_csv("vw_pnl_final.csv"))
waterfall = clean_columns(pd.read_csv("vw_pnl_waterfall.csv"))
yoy = clean_columns(pd.read_csv("vw_yoy_variance.csv"))
budget = clean_columns(pd.read_csv("vw_budget_vs_actual.csv"))
expense = clean_columns(pd.read_csv("vw_expense_positive.csv"))
ytd = clean_columns(pd.read_csv("vw_ytd_summary.csv"))


# Fill missing values
pnl = pnl.fillna(0)
yoy = yoy.fillna(0)
budget = budget.fillna(0)
expense = expense.fillna(0)
ytd = ytd.fillna(0)

# -----------------------------
# Create Presentation
# -----------------------------
prs = Presentation()
blank_layout = prs.slide_layouts[5]

# -----------------------------
# 1. Title Slide
# -----------------------------
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "Auto Financial Pack - FY 2025"

# -----------------------------
# 2. Revenue & Profit Trend
# -----------------------------
slide = prs.slides.add_slide(blank_layout)
slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5)).text = "Revenue & Profit Trend"

chart_data = CategoryChartData()
chart_data.categories = pnl['period'].astype(str).tolist()
chart_data.add_series('Revenue', pnl['revenue'].tolist())
chart_data.add_series('EBITDA', pnl['ebitda'].tolist())
chart_data.add_series('Profit After Tax', pnl['profit_after_tax'].tolist())

slide.shapes.add_chart(
    XL_CHART_TYPE.LINE,
    Inches(1), Inches(1),
    Inches(8), Inches(4.5),
    chart_data
)

# -----------------------------
# 3. YOY Growth
# -----------------------------
slide = prs.slides.add_slide(blank_layout)
slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5)).text = "Year Over Year Growth %"

chart_data = CategoryChartData()
chart_data.categories = yoy['quarter'].astype(str).tolist()
chart_data.add_series('YOY Growth %', yoy['yoy_growth_percentage'].tolist())

slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(1), Inches(1),
    Inches(8), Inches(4.5),
    chart_data
)

# -----------------------------
# 4. Budget vs Actual
# -----------------------------
slide = prs.slides.add_slide(blank_layout)
slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5)).text = "Budget vs Actual"

chart_data = CategoryChartData()
chart_data.categories = budget['pnl_line'].astype(str).tolist()
chart_data.add_series('Actual', budget['actual_amount'].tolist())
chart_data.add_series('Budget', budget['budget_amount'].tolist())

slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(1), Inches(1),
    Inches(8), Inches(4.5),
    chart_data
)

# -----------------------------
# 5. Expense Breakdown
# -----------------------------
slide = prs.slides.add_slide(blank_layout)
slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5)).text = "Expense Breakdown"

chart_data = CategoryChartData()
chart_data.categories = expense['period'].astype(str).tolist()
chart_data.add_series('Opex', expense['opex'].tolist())
chart_data.add_series('Depreciation', expense['depreciation'].tolist())
chart_data.add_series('Tax', expense['tax'].tolist())

slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_STACKED,
    Inches(1), Inches(1),
    Inches(8), Inches(4.5),
    chart_data
)

# -----------------------------
# 6. YTD Summary
# -----------------------------
slide = prs.slides.add_slide(blank_layout)
slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5)).text = "YTD Summary"

textbox = slide.shapes.add_textbox(Inches(2), Inches(2), Inches(6), Inches(3))
tf = textbox.text_frame

tf.text = f"YTD Revenue: {ytd['ytd_revenue'].iloc[0]:,.2f}"

p = tf.add_paragraph()
p.text = f"YTD EBITDA: {ytd['ytd_ebitda'].iloc[0]:,.2f}"

p = tf.add_paragraph()
p.text = f"YTD Operating Profit: {ytd['ytd_operating_profit'].iloc[0]:,.2f}"

p = tf.add_paragraph()
p.text = f"YTD Profit After Tax: {ytd['ytd_profit_after_tax'].iloc[0]:,.2f}"


# =====================================================
# 7. Product Revenue Comparison (2025)
# =====================================================
product_rev = clean_columns(pd.read_csv("vw_product_revenue_comparison.csv"))
product_rev =_finalize = product_rev.fillna(0)

# Filter latest year (2025)
product_rev_2025 = product_rev[product_rev['year'] == 2025]

slide = prs.slides.add_slide(blank_layout)
slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5)).text = "Product Revenue Comparison - 2025"

chart_data = CategoryChartData()
chart_data.categories = product_rev_2025['product_line']

chart_data.add_series(
    'Revenue',
    product_rev_2025['revenue'].tolist()
)

slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(1), Inches(1),
    Inches(8), Inches(4.5),
    chart_data
)

# =====================================================
# 8. Product Profit Comparison (2025)
# =====================================================
product_profit = clean_columns(pd.read_csv("vw_product_profit_comparison.csv"))
product_profit = product_profit.fillna(0)

product_profit_2025 = product_profit[product_profit['year'] == 2025]

slide = prs.slides.add_slide(blank_layout)
slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5)).text = "Product Profit Comparison - 2025"

chart_data = CategoryChartData()
chart_data.categories = product_profit_2025['product_line']

chart_data.add_series(
    'Profit',
    product_profit_2025['profit'].tolist()
)

slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(1), Inches(1),
    Inches(8), Inches(4.5),
    chart_data
)

# =====================================================
# 9. Product Profit Margin %
# =====================================================
product_margin = clean_columns(pd.read_csv("vw_product_margin_comparison.csv"))
product_margin = product_margin.fillna(0)

product_margin_2025 = product_margin[product_margin['year'] == 2025]

slide = prs.slides.add_slide(blank_layout)
slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5)).text = "Product Profit Margin % - 2025"

chart_data = CategoryChartData()
chart_data.categories = product_margin_2025['product_line']

chart_data.add_series(
    'Profit Margin %',
    (product_margin_2025['profit_margin'] * 100).tolist()
)

slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(1), Inches(1),
    Inches(8), Inches(4.5),
    chart_data
)

# =====================================================
# 10. Product YOY Growth %
# =====================================================
product_yoy = clean_columns(pd.read_csv("vw_product_yoy.csv"))
product_yoy = product_yoy.fillna(0)

product_yoy_2025 = product_yoy[product_yoy['year'] == 2025]

slide = prs.slides.add_slide(blank_layout)
slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5)).text = "Product YOY Revenue Growth % - 2025"

chart_data = CategoryChartData()
chart_data.categories = product_yoy_2025['product_line']

chart_data.add_series(
    'Product YOY Growth %',
    (product_yoy_2025['yoy_growth_percent'] * 100).tolist()
)

slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(1), Inches(1),
    Inches(8), Inches(4.5),
    chart_data
)

# -----------------------------
# Save file
# -----------------------------
prs.save("Auto_Financial_Pack_2025.pptx")

print("Financial Pack Generated Successfully 🚀")