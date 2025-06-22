import pandas as pd
import streamlit as st
from datetime import datetime

# Load data from Excel files
@st.cache_data
def load_data():
    sales = pd.read_excel("sales_summary.xlsx")
    stock = pd.read_excel("Stocks.xlsx")
    return sales, stock

# Generate purchase plan based on sales and stock data
def generate_plan(sales, stock, target_month, target_year):
    last_year = target_year - 1
    prev_month = target_month - 1

    # Months from last year
    months_last_year = [target_month - i for i in range(1, 4) if (target_month - i) > 0]

    # Filter sales data
    sales_last = sales[
        (sales['Year'] == last_year) & (sales['Month'].isin(months_last_year))
    ]
    sales_now = sales[
        (sales['Year'] == target_year) & (sales['Month'] == prev_month)
    ]

    combined = pd.concat([sales_last, sales_now])
    sales_summary = combined.groupby('Barcode')['Quantity'].sum().reset_index()
    sales_summary.rename(columns={'Quantity': 'Recent_Sales'}, inplace=True)

    df = stock.merge(sales_summary, on='Barcode', how='left')
    df['Recent_Sales'] = df['Recent_Sales'].fillna(0)
    df['Recommended_Purchase'] = df['Recent_Sales'] - df['Quantity On Hand']
    df['Recommended_Purchase'] = df['Recommended_Purchase'].apply(lambda x: max(x, 0))

    return df[['Barcode', 'Name', 'Product Category/Complete Name', 'Quantity On Hand', 'Recent_Sales', 'Recommended_Purchase']]

# Streamlit user interface
def main():
    st.title("نموذج اقتراح المشتريات")
    target_month = st.selectbox("اختر الشهر", list(range(1, 13)))
    target_year = st.number_input("أدخل السنة", value=datetime.now().year)

    sales, stock = load_data()

    if st.button("توليد خطة الشراء"):
        plan = generate_plan(sales, stock, target_month, target_year)
        st.success("✅ تم توليد خطة الشراء.")
        st.dataframe(plan)
        st.download_button("📥 تحميل Excel", data=plan.to_csv(index=False), file_name="purchase_plan.csv", mime="text/csv")

if __name__ == "__main__":
    main()
