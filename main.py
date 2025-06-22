import pandas as pd 
import streamlit as st 
from datetime import datetime 
import io
 
# Load data from Excel files 
@st.cache_data 
def load_data(): 
    sales = pd.read_excel("sales_summary.xlsx") 
    stock = pd.read_excel("Stocks.xlsx") 
    purchases = pd.read_excel("Purchase.xlsx")  # ملف المشتريات
    return sales, stock, purchases 

# Convert DataFrame to Excel bytes
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='خطة الشراء')
    processed_data = output.getvalue()
    return processed_data
 
# Generate purchase plan based on sales and stock data 
def generate_plan(sales, stock, purchases, target_month, target_year): 
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
    
    # حساب عدد الشهور اللي فيها مبيعات فعلية لكل منتج
    months_with_sales = combined.groupby('Barcode').size().reset_index(name='Months_With_Sales')
    
    # حساب إجمالي المبيعات
    sales_summary = combined.groupby('Barcode')['Quantity'].sum().reset_index()
    
    # دمج عدد الشهور مع المبيعات
    sales_summary = sales_summary.merge(months_with_sales, on='Barcode')
    
    # حساب متوسط المبيعات على أساس الشهور الفعلية
    sales_summary['Average_Monthly_Sales'] = sales_summary['Quantity'] / sales_summary['Months_With_Sales']
    sales_summary.rename(columns={'Quantity': 'Total_Sales_Period'}, inplace=True)
    
    # تحضير بيانات المشتريات لنفس الفترة
    purchases['Date'] = pd.to_datetime(purchases['Date'])
    purchases['Year'] = purchases['Date'].dt.year
    purchases['Month'] = purchases['Date'].dt.month
    
    # مشتريات الشهر السابق
    purchases_prev_month = purchases[
        (purchases['Year'] == prev_year) & (purchases['Month'] == prev_month)
    ]
    
    # مشتريات الـ 3 شهور في السنة السابقة
    purchases_last_year = purchases[
        (purchases['Year'] == last_year) & (purchases['Month'].isin(months_last_year))
    ]
    
    # دمج بيانات المشتريات
    combined_purchases = pd.concat([purchases_prev_month, purchases_last_year])
    
    # حساب عدد الشهور اللي فيها مشتريات فعلية
    months_with_purchases = combined_purchases.groupby('Barcode').size().reset_index(name='Months_With_Purchases')
    
    purchases_summary = combined_purchases.groupby('Barcode').agg({
        'purchase': 'sum',  # إجمالي المشتريات
        'اسم المورد': lambda x: ', '.join(x.unique())  # أسماء الموردين
    }).reset_index()
    
    # دمج عدد الشهور مع المشتريات
    purchases_summary = purchases_summary.merge(months_with_purchases, on='Barcode')
    
    purchases_summary.rename(columns={
        'purchase': 'Total_Purchases_Period',
        'اسم المورد': 'Suppliers'
    }, inplace=True)
    
    # إضافة بيانات الموردين من ملف المخزون إذا كان متوفر
    if 'اسم المورد' in stock.columns:
        stock_suppliers = stock[['Barcode', 'اسم المورد']].dropna()
        stock_suppliers.rename(columns={'اسم المورد': 'Stock_Supplier'}, inplace=True)
    else:
        stock_suppliers = pd.DataFrame(columns=['Barcode', 'Stock_Supplier'])
 
    # دمج مع بيانات المخزون
    df = stock.merge(sales_summary, on='Barcode', how='left')
    df = df.merge(purchases_summary, on='Barcode', how='left')
    
    # دمج بيانات الموردين من المخزون إذا كان متوفر
    if len(stock_suppliers) > 0:
        df = df.merge(stock_suppliers, on='Barcode', how='left')
        # دمج الموردين من المشتريات والمخزون
        df['All_Suppliers'] = df.apply(lambda row: 
            ', '.join(filter(None, [
                str(row.get('Suppliers', '')).strip() if pd.notna(row.get('Suppliers')) and str(row.get('Suppliers')).strip() != 'غير محدد' else '',
                str(row.get('Stock_Supplier', '')).strip() if pd.notna(row.get('Stock_Supplier')) else ''
            ])), axis=1)
        df['All_Suppliers'] = df['All_Suppliers'].apply(lambda x: x if x else 'غير محدد')
    else:
        df['All_Suppliers'] = df['Suppliers'].fillna('غير محدد')
    
    # ملء القيم المفقودة
    df['Total_Sales_Period'] = df['Total_Sales_Period'].fillna(0)
    df['Average_Monthly_Sales'] = df['Average_Monthly_Sales'].fillna(0)
    df['Months_With_Sales'] = df['Months_With_Sales'].fillna(0)
    df['Total_Purchases_Period'] = df['Total_Purchases_Period'].fillna(0)
    df['Months_With_Purchases'] = df['Months_With_Purchases'].fillna(0)
    df['Suppliers'] = df['Suppliers'].fillna('غير محدد')
    
    # التأكد من وجود عمود All_Suppliers
    if 'All_Suppliers' not in df.columns:
        df['All_Suppliers'] = df['Suppliers']
    
    # حساب معدل الدوران
    # معدل الدوران = إجمالي المبيعات ÷ متوسط المخزون
    # متوسط المخزون = (المخزون الحالي + المشتريات) ÷ 2
    df['Average_Inventory'] = (df['Quantity On Hand'] + df['Total_Purchases_Period']) / 2
    df['Average_Inventory'] = df['Average_Inventory'].replace(0, 1)  # تجنب القسمة على صفر
    
    df['Inventory_Turnover_Rate'] = df['Total_Sales_Period'] / df['Average_Inventory']
    df['Inventory_Turnover_Rate'] = df['Inventory_Turnover_Rate'].round(2)
    
    # تصنيف سرعة الدوران
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
    
    # حساب الشراء المقترح بناءً على متوسط المبيعات الشهرية
    df['Recommended_Purchase'] = df['Average_Monthly_Sales'] - df['Quantity On Hand']
    df['Recommended_Purchase'] = df['Recommended_Purchase'].apply(lambda x: max(x, 0))
 
    # ترتيب الأعمدة
    result_df = df[[
        'Barcode', 
        'Name', 
        'Product Category/Complete Name', 
        'Quantity On Hand', 
        'Total_Sales_Period',
        'Months_With_Sales',
        'Total_Purchases_Period',
        'Months_With_Purchases',
        'Average_Monthly_Sales',
        'Average_Inventory',
        'Inventory_Turnover_Rate',
        'Turnover_Classification',
        'Recommended_Purchase',
        'Suppliers'
    ]].copy()
    
    # إعادة تسمية الأعمدة بالعربية
    result_df.columns = [
        'الباركود',
        'اسم المنتج', 
        'فئة المنتج',
        'الكمية المتاحة',
        'إجمالي المبيعات في الفترة',
        'عدد الشهور بمبيعات',
        'إجمالي المشتريات في الفترة',
        'عدد الشهور بمشتريات',
        'متوسط المبيعات الشهرية',
        'متوسط المخزون',
        'معدل الدوران',
        'تصنيف الدوران',
        'الشراء المقترح',
        'الموردين'
    ]
    
    return result_df
 
# Streamlit user interface 
def main(): 
    st.title("نموذج اقتراح المشتريات") 
    
    st.write("سيتم حساب الشراء بناءً على:")
    st.write("- الشهر السابق للشهر المختار")
    st.write("- 3 شهور مقابلة في السنة السابقة")
    st.write("- متوسط المبيعات الشهرية (حسب الشهور الفعلية بمبيعات)")
    st.write("- معدل دوران المخزون وتصنيف المنتجات")
    st.write("- بيانات الموردين للمنتجات")
    
    target_month = st.selectbox("اختر الشهر", 
                               options=list(range(1, 13)),
                               format_func=lambda x: [
                                   "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
                                   "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"
                               ][x-1])
    
    target_year = st.number_input("أدخل السنة", value=datetime.now().year, min_value=2020, max_value=2030) 
 
    try:
        sales, stock, purchases = load_data() 
        st.success(f"✅ تم تحميل البيانات بنجاح")
        st.write(f"عدد منتجات المبيعات: {len(sales)}")
        st.write(f"عدد منتجات المخزون: {len(stock)}")
        st.write(f"عدد سجلات المشتريات: {len(purchases)}")
    except FileNotFoundError as e:
        st.error("❌ لم يتم العثور على ملفات البيانات. تأكد من وجود sales_summary.xlsx و Stocks.xlsx و purchases.xlsx")
        return
    except Exception as e:
        st.error(f"❌ خطأ في تحميل البيانات: {str(e)}")
        return
 
    if st.button("توليد خطة الشراء"): 
        try:
            plan = generate_plan(sales, stock, purchases, target_month, target_year) 
            st.success("✅ تم توليد خطة الشراء بنجاح.")
            
            # عرض إحصائيات سريعة
            total_recommended = plan['الشراء المقترح'].sum()
            products_to_buy = len(plan[plan['الشراء المقترح'] > 0])
            avg_turnover = plan['معدل الدوران'].mean()
            fast_moving = len(plan[plan['تصنيف الدوران'].isin(['سريع', 'سريع جداً'])])
            slow_moving = len(plan[plan['تصنيف الدوران'].isin(['بطيء', 'راكد'])])
            
            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                st.metric("إجمالي القطع المقترح شراؤها", f"{total_recommended:,.0f}")
            with col2:
                st.metric("عدد المنتجات المطلوب شراؤها", products_to_buy)
            with col3:
                st.metric("متوسط معدل الدوران", f"{avg_turnover:.2f}")
            with col4:
                st.metric("منتجات سريعة الحركة", fast_moving)
            with col5:
                st.metric("منتجات بطيئة/راكدة", slow_moving)
            
            # تصفية حسب تصنيف الدوران
            st.subheader("🔍 تصفية النتائج")
            turnover_filter = st.multiselect(
                "اختر تصنيف الدوران للعرض:",
                options=['سريع جداً', 'سريع', 'متوسط', 'بطيء', 'راكد'],
                default=['سريع جداً', 'سريع', 'متوسط', 'بطيء', 'راكد']
            )
            
            filtered_plan = plan[plan['تصنيف الدوران'].isin(turnover_filter)]
            
            # عرض الجدول
            st.dataframe(filtered_plan, use_container_width=True)
            
            # تحليل إضافي
            st.subheader("📊 تحليل معدل الدوران")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**توزيع المنتجات حسب تصنيف الدوران:**")
                turnover_counts = plan['تصنيف الدوران'].value_counts()
                for category, count in turnover_counts.items():
                    percentage = (count / len(plan)) * 100
                    st.write(f"• {category}: {count} منتج ({percentage:.1f}%)")
            
            with col2:
                st.write("**أهم الموردين (للمنتجات المطلوب شراؤها):**")
                products_to_purchase = plan[plan['الشراء المقترح'] > 0]
                if len(products_to_purchase) > 0:
                    # فصل الموردين المتعددين وعدهم
                    all_suppliers = []
                    for suppliers_str in products_to_purchase['الموردين']:
                        if suppliers_str != 'غير محدد':
                            all_suppliers.extend([s.strip() for s in suppliers_str.split(',')])
                    
                    if all_suppliers:
                        supplier_counts = pd.Series(all_suppliers).value_counts().head(5)
                        for supplier, count in supplier_counts.items():
                            st.write(f"• {supplier}: {count} منتج")
                    else:
                        st.write("لا توجد بيانات موردين متاحة")
                else:
                    st.write("لا توجد منتجات مطلوب شراؤها")
            
            # زر التحميل للإكسيل
            excel_file = to_excel(filtered_plan)
            st.download_button(
                label="📥 تحميل ملف Excel",
                data=excel_file,
                file_name=f"purchase_plan_with_turnover_{target_year}_{target_month:02d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"❌ خطأ في توليد خطة الشراء: {str(e)}")
            st.write("تفاصيل الخطأ:", str(e))
 
if __name__ == "__main__": 
    main()
