import pandas as pd
import streamlit as st
from datetime import datetime
import io

@st.cache_data
def load_data():
    sales = pd.read_excel("sales_summary.xlsx")
    stock = pd.read_excel("Stocks.xlsx")
    purchases = pd.read_excel("Purchase.xlsx")
    return sales, stock, purchases

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='خطة الشراء')
    return output.getvalue()

def generate_plan(sales, stock, purchases, target_month, target_year):
    last_year = target_year - 1
    prev_month = target_month - 1 if target_month > 1 else 12
    prev_year = target_year if target_month > 1 else target_year - 1

    sales_prev_month = sales[(sales['Year'] == prev_year) & (sales['Month'] == prev_month)]
    months_last_year = [target_month - i for i in range(1, 4) if target_month - i > 0]
    if len(months_last_year) < 3:
        months_last_year += list(range(12 - (3 - len(months_last_year)) + 1, 13))

    sales_last_year = sales[(sales['Year'] == last_year) & (sales['Month'].isin(months_last_year))]
    combined = pd.concat([sales_prev_month, sales_last_year])

    # عدد الشهور بمبيعات
    months_with_sales = combined[['Barcode', 'Year', 'Month']].drop_duplicates()
    months_with_sales = months_with_sales.groupby('Barcode').size().reset_index(name='Months_With_Sales')

    # عدد الفواتير
    if 'Order Reference' in combined.columns:
        invoice_count = combined[['Barcode', 'Order Reference']].drop_duplicates().groupby('Barcode').size().reset_index(name='Invoice_Count')
    else:
        invoice_count = pd.DataFrame(columns=['Barcode', 'Invoice_Count'])

    # مبيعات
    sales_summary = combined.groupby('Barcode')['Quantity'].sum().reset_index()
    sales_summary = sales_summary.merge(months_with_sales, on='Barcode', how='left')
    sales_summary['Average_Monthly_Sales'] = sales_summary['Quantity'] / sales_summary['Months_With_Sales']
    sales_summary.rename(columns={'Quantity': 'Total_Sales_Period'}, inplace=True)

    # مشتريات
    purchases['Date'] = pd.to_datetime(purchases['Date'])
    purchases['Year'] = purchases['Date'].dt.year
    purchases['Month'] = purchases['Date'].dt.month
    purchases_prev_month = purchases[(purchases['Year'] == prev_year) & (purchases['Month'] == prev_month)]
    purchases_last_year = purchases[(purchases['Year'] == last_year) & (purchases['Month'].isin(months_last_year))]
    combined_purchases = pd.concat([purchases_prev_month, purchases_last_year])
    months_with_purchases = combined_purchases[['Barcode', 'Year', 'Month']].drop_duplicates()
    months_with_purchases = months_with_purchases.groupby('Barcode').size().reset_index(name='Months_With_Purchases')
    purchases_summary = combined_purchases.groupby('Barcode').agg({
        'purchase': 'sum',
        'اسم المورد': lambda x: ', '.join(x.dropna().unique())
    }).reset_index()
    purchases_summary = purchases_summary.merge(months_with_purchases, on='Barcode', how='left')
    purchases_summary.rename(columns={'purchase': 'Total_Purchases_Period', 'اسم المورد': 'Suppliers'}, inplace=True)

    # دمج مع المخزون
    df = stock.merge(sales_summary, on='Barcode', how='left')
    df = df.merge(purchases_summary, on='Barcode', how='left')
    df = df.merge(invoice_count, on='Barcode', how='left')

    # الموردين من المخزون
    if 'اسم المورد' in stock.columns:
        stock_suppliers = stock[['Barcode', 'اسم المورد']].dropna().rename(columns={'اسم المورد': 'Stock_Supplier'})
        df = df.merge(stock_suppliers, on='Barcode', how='left')
    else:
        df['Stock_Supplier'] = None

    df['All_Suppliers'] = df.apply(lambda row: ', '.join(set(filter(None, [row.get('Suppliers'), row.get('Stock_Supplier')]))) or 'غير محدد', axis=1)

    # ملء النواقص
    df.fillna({'Total_Sales_Period': 0, 'Average_Monthly_Sales': 0, 'Months_With_Sales': 0,
               'Total_Purchases_Period': 0, 'Months_With_Purchases': 0, 'Suppliers': 'غير محدد',
               'Invoice_Count': 0}, inplace=True)

    # معدل الدوران العادي
    df['Average_Inventory'] = (df['Quantity On Hand'] + df['Total_Purchases_Period']) / 2
    df['Average_Inventory'] = df['Average_Inventory'].replace(0, 1)
    df['Inventory_Turnover_Rate'] = df['Total_Sales_Period'] / df['Average_Inventory']

    # معدل الدوران بالفواتير
    df['Invoice_Turnover_Rate'] = df['Invoice_Count'] / df['Average_Inventory']

    def classify_turnover(rate):
        if rate >= 4:
            return 'سريع جداً'
        elif rate >= 2:
            return 'سريع'
        elif rate >= 1:
            return 'متوسط'
        elif rate >= 0.5:
            return 'بطيء'
        else:
            return 'راكد'

    df['Turnover_Classification'] = df['Inventory_Turnover_Rate'].apply(classify_turnover)
    df['Recommended_Purchase'] = df['Average_Monthly_Sales'] - df['Quantity On Hand']
    df['Recommended_Purchase'] = df['Recommended_Purchase'].apply(lambda x: max(x, 0))

    result_df = df[[
        'Barcode', 'Name', 'Product Category/Complete Name', 'Quantity On Hand',
        'Total_Sales_Period', 'Months_With_Sales', 'Total_Purchases_Period', 'Months_With_Purchases',
        'Average_Monthly_Sales', 'Average_Inventory', 'Inventory_Turnover_Rate',
        'Invoice_Count', 'Invoice_Turnover_Rate',
        'Turnover_Classification', 'Recommended_Purchase', 'All_Suppliers'
    ]].copy()

    result_df.columns = [
        'الباركود', 'اسم المنتج', 'فئة المنتج', 'الكمية المتاحة',
        'إجمالي المبيعات في الفترة', 'عدد الشهور بمبيعات', 'إجمالي المشتريات في الفترة', 'عدد الشهور بمشتريات',
        'متوسط المبيعات الشهرية', 'متوسط المخزون', 'معدل الدوران',
        'عدد الفواتير', 'معدل الدوران حسب الفواتير',
        'تصنيف الدوران', 'الشراء المقترح', 'الموردين'
    ]

    return result_df

# واجهة التطبيق
def main():
    st.title("🛒 نموذج اقتراح المشتريات وتحليل الدوران")
    target_month = st.selectbox("اختر الشهر", list(range(1, 13)), format_func=lambda x: [
        "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
        "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"
    ][x-1])
    target_year = st.number_input("أدخل السنة", value=datetime.now().year, min_value=2020, max_value=2030)

    try:
        sales, stock, purchases = load_data()
    except:
        st.error("❌ تأكد من وجود ملفات sales_summary.xlsx و Stocks.xlsx و Purchase.xlsx")
        return

    if st.button("📊 توليد الخطة"):
        result = generate_plan(sales, stock, purchases, target_month, target_year)
        st.success("✅ تم إنشاء خطة الشراء بنجاح.")
        st.dataframe(result, use_container_width=True)
        excel_file = to_excel(result)
        st.download_button("📥 تحميل Excel", data=excel_file, file_name=f"purchase_plan_{target_year}_{target_month}.xlsx")

if __name__ == "__main__":
    main()
