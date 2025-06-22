import pandas as pd 
import streamlit as st 
from datetime import datetime 
import io
 
# Load data from Excel files 
@st.cache_data 
def load_data(): 
    sales = pd.read_excel("sales_summary.xlsx") 
    stock = pd.read_excel("Stocks.xlsx") 
    return sales, stock 

# Convert DataFrame to Excel bytes
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='خطة الشراء')
    processed_data = output.getvalue()
    return processed_data
 
# Generate purchase plan based on sales and stock data 
def generate_plan(sales, stock, target_month, target_year): 
    last_year = target_year - 1 
    prev_month = target_month - 1 if target_month > 1 else 12
    prev_year = target_year if target_month > 1 else target_year - 1
 
    # الشهر اللي قبل الفلتر
    sales_prev_month = sales[
        (sales['Year'] == prev_year) & (sales['Month'] == prev_month)
    ]
    
    # الـ 3 شهور في السنة اللي قبلها
    months_last_year = [target_month - i for i in range(1, 4) if (target_month - i) > 0]
    # إذا كان target_month أقل من 4، نأخذ الشهور من آخر السنة السابقة
    if len(months_last_year) < 3:
        remaining_months = 12 - (3 - len(months_last_year)) + 1
        months_last_year.extend(list(range(remaining_months, 13)))
    
    sales_last_year = sales[
        (sales['Year'] == last_year) & (sales['Month'].isin(months_last_year))
    ]
 
    # دمج البيانات
    combined = pd.concat([sales_prev_month, sales_last_year])
    sales_summary = combined.groupby('Barcode')['Quantity'].sum().reset_index()
    
    # حساب متوسط المبيعات (إجمالي المبيعات ÷ 4 شهور)
    sales_summary['Average_Monthly_Sales'] = sales_summary['Quantity'] / 4
    sales_summary.rename(columns={'Quantity': 'Total_Sales_4_Months'}, inplace=True)
 
    # دمج مع بيانات المخزون
    df = stock.merge(sales_summary, on='Barcode', how='left')
    df['Total_Sales_4_Months'] = df['Total_Sales_4_Months'].fillna(0)
    df['Average_Monthly_Sales'] = df['Average_Monthly_Sales'].fillna(0)
    
    # حساب الشراء المقترح بناءً على متوسط المبيعات الشهرية
    df['Recommended_Purchase'] = df['Average_Monthly_Sales'] - df['Quantity On Hand']
    df['Recommended_Purchase'] = df['Recommended_Purchase'].apply(lambda x: max(x, 0))
 
    # ترتيب الأعمدة
    result_df = df[[
        'Barcode', 
        'Name', 
        'Product Category/Complete Name', 
        'Quantity On Hand', 
        'Total_Sales_4_Months',
        'Average_Monthly_Sales', 
        'Recommended_Purchase'
    ]].copy()
    
    # إعادة تسمية الأعمدة بالعربية
    result_df.columns = [
        'الباركود',
        'اسم المنتج', 
        'فئة المنتج',
        'الكمية المتاحة',
        'إجمالي مبيعات 4 شهور',
        'متوسط المبيعات الشهرية',
        'الشراء المقترح'
    ]
    
    return result_df
 
# Streamlit user interface 
def main(): 
    st.title("نموذج اقتراح المشتريات") 
    
    st.write("سيتم حساب الشراء بناءً على:")
    st.write("- الشهر السابق للشهر المختار")
    st.write("- 3 شهور مقابلة في السنة السابقة")
    st.write("- متوسط المبيعات الشهرية للـ 4 شهور")
    
    target_month = st.selectbox("اختر الشهر", 
                               options=list(range(1, 13)),
                               format_func=lambda x: [
                                   "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
                                   "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"
                               ][x-1])
    
    target_year = st.number_input("أدخل السنة", value=datetime.now().year, min_value=2020, max_value=2030) 
 
    try:
        sales, stock = load_data() 
        st.success(f"✅ تم تحميل البيانات بنجاح")
        st.write(f"عدد منتجات المبيعات: {len(sales)}")
        st.write(f"عدد منتجات المخزون: {len(stock)}")
    except FileNotFoundError as e:
        st.error("❌ لم يتم العثور على ملفات البيانات. تأكد من وجود sales_summary.xlsx و Stocks.xlsx")
        return
    except Exception as e:
        st.error(f"❌ خطأ في تحميل البيانات: {str(e)}")
        return
 
    if st.button("توليد خطة الشراء"): 
        try:
            plan = generate_plan(sales, stock, target_month, target_year) 
            st.success("✅ تم توليد خطة الشراء بنجاح.")
            
            # عرض إحصائيات سريعة
            total_recommended = plan['الشراء المقترح'].sum()
            products_to_buy = len(plan[plan['الشراء المقترح'] > 0])
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("إجمالي القطع المقترح شراؤها", f"{total_recommended:,.0f}")
            with col2:
                st.metric("عدد المنتجات المطلوب شراؤها", products_to_buy)
            with col3:
                st.metric("إجمالي المنتجات", len(plan))
            
            # عرض الجدول
            st.dataframe(plan, use_container_width=True)
            
            # زر التحميل للإكسيل
            excel_file = to_excel(plan)
            st.download_button(
                label="📥 تحميل ملف Excel",
                data=excel_file,
                file_name=f"purchase_plan_{target_year}_{target_month:02d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"❌ خطأ في توليد خطة الشراء: {str(e)}")
 
if __name__ == "__main__": 
    main()
