import pandas as pd 
import streamlit as st 
from datetime import datetime, timedelta
import io
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Configuration
st.set_page_config(
    page_title="نظام اقتراح المشتريات",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better RTL support
st.markdown("""
<style>
    .main > div {
        direction: rtl;
        text-align: right;
    }
    .stSelectbox > div > div {
        direction: rtl;
    }
    .metric-card {
        background: linear-gradient(45deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin: 0.5rem 0;
    }
    .alert-high { background-color: #ff4444; color: white; padding: 0.5rem; border-radius: 5px; }
    .alert-medium { background-color: #ffaa00; color: white; padding: 0.5rem; border-radius: 5px; }
    .alert-low { background-color: #00aa00; color: white; padding: 0.5rem; border-radius: 5px; }
</style>
""", unsafe_allow_html=True)

# Load data from Excel files with improved error handling
@st.cache_data(ttl=300)  # Cache for 5 minutes
def load_data():
    """Load data from Excel files with enhanced error handling"""
    try:
        sales = pd.read_excel("sales_summary.xlsx")
        stock = pd.read_excel("Stocks.xlsx") 
        purchases = pd.read_excel("Purchase.xlsx")
        
        required_sales_cols = ['Barcode', 'Year', 'Month', 'Quantity']
        required_stock_cols = ['Barcode', 'Name', 'Quantity On Hand']
        required_purchase_cols = ['Barcode', 'Date', 'purchase']
        
        missing_sales = [col for col in required_sales_cols if col not in sales.columns]
        missing_stock = [col for col in required_stock_cols if col not in stock.columns]
        missing_purchase = [col for col in required_purchase_cols if col not in purchases.columns]
        
        if missing_sales or missing_stock or missing_purchase:
            if missing_sales: st.error(f"أعمدة مفقودة في ملف المبيعات: {missing_sales}")
            if missing_stock: st.error(f"أعمدة مفقودة في ملف المخزون: {missing_stock}")
            if missing_purchase: st.error(f"أعمدة مفقودة في ملف المشتريات: {missing_purchase}")
            return None, None, None
            
        sales['Barcode'] = sales['Barcode'].astype(str)
        stock['Barcode'] = stock['Barcode'].astype(str)
        purchases['Barcode'] = purchases['Barcode'].astype(str)
        
        sales['Quantity'] = sales['Quantity'].clip(lower=0)
        stock['Quantity On Hand'] = stock['Quantity On Hand'].clip(lower=0)
        purchases['purchase'] = purchases['purchase'].clip(lower=0)
        
        return sales, stock, purchases
        
    except FileNotFoundError as e:
        st.error(f"❌ ملف غير موجود: {str(e)}. تأكد من وجود الملفات sales_summary.xlsx, Stocks.xlsx, Purchase.xlsx في نفس المجلد.")
        return None, None, None
    except Exception as e:
        st.error(f"❌ خطأ في تحميل البيانات: {str(e)}")
        return None, None, None

# Convert DataFrame to Excel bytes with formatting
def to_excel(df, filename_prefix="purchase_plan"):
    """Convert DataFrame to Excel with enhanced formatting"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='خطة الشراء')
        workbook = writer.book
        worksheet = writer.sheets['خطة الشراء']
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
        
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        
        for col_num, column_title in enumerate(df.columns, 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            column_letter = get_column_letter(col_num)
            max_length = max(df[column_title].astype(str).map(len).max(), len(column_title))
            adjusted_width = (max_length + 2) * 1.2
            worksheet.column_dimensions[column_letter].width = min(adjusted_width, 50)
    return output.getvalue()

# Enhanced purchase plan generation
def generate_plan(sales, stock, purchases, target_month, target_year, safety_stock_days=30):
    """Generate purchase plan with improved algorithms and annualized turnover rates"""
    
    last_year = target_year - 1 
    prev_month = target_month - 1 if target_month > 1 else 12
    prev_year = target_year if target_month > 1 else target_year - 1
 
    sales_prev_month = sales[(sales['Year'] == prev_year) & (sales['Month'] == prev_month)]
    months_last_year = [(target_month - i - 1) % 12 + 1 for i in range(3)]
    sales_last_year = sales[(sales['Year'] == last_year) & (sales['Month'].isin(months_last_year))]
 
    combined_sales = pd.concat([sales_prev_month, sales_last_year])
    
    if combined_sales.empty:
        st.warning("⚠️ لا توجد بيانات مبيعات للفترة المحددة")
        return pd.DataFrame()
    
    invoice_col = 'Order Reference' if 'Order Reference' in combined_sales.columns else 'Month'
    
    months_with_sales = combined_sales.groupby('Barcode').agg(
        Months_With_Sales=('Month', 'nunique'),
        Total_Sales_Period=('Quantity', 'sum'),
        Invoice_Count=(invoice_col, 'nunique')
    ).reset_index()
    
    months_with_sales['Average_Monthly_Sales'] = (months_with_sales['Total_Sales_Period'] / months_with_sales['Months_With_Sales']).replace([np.inf, -np.inf], 0).fillna(0)
    months_with_sales['Average_Monthly_Invoices'] = (months_with_sales['Invoice_Count'] / months_with_sales['Months_With_Sales']).replace([np.inf, -np.inf], 0).fillna(0)
    
    purchases['Date'] = pd.to_datetime(purchases['Date'], errors='coerce')
    purchases = purchases.dropna(subset=['Date'])
    purchases['Year'] = purchases['Date'].dt.year
    purchases['Month'] = purchases['Date'].dt.month
    
    purchases_prev_month = purchases[(purchases['Year'] == prev_year) & (purchases['Month'] == prev_month)]
    purchases_last_year = purchases[(purchases['Year'] == last_year) & (purchases['Month'].isin(months_last_year))]
    combined_purchases = pd.concat([purchases_prev_month, purchases_last_year])
    
    if not combined_purchases.empty:
        supplier_col = 'اسم المورد' if 'اسم المورد' in combined_purchases.columns else 'Barcode'
        purchases_summary = combined_purchases.groupby('Barcode').agg(
            Total_Purchases_Period=('purchase', 'sum'),
            Suppliers=(supplier_col, lambda x: ', '.join(x.dropna().unique()))
        ).reset_index()
    else:
        purchases_summary = pd.DataFrame(columns=['Barcode', 'Total_Purchases_Period', 'Suppliers'])
    
    df = stock.merge(months_with_sales, on='Barcode', how='left').merge(purchases_summary, on='Barcode', how='left')
    
    for col in ['Total_Sales_Period', 'Average_Monthly_Sales', 'Months_With_Sales', 'Invoice_Count', 'Average_Monthly_Invoices', 'Total_Purchases_Period']:
        if col in df.columns:
            df[col] = df[col].fillna(0)
    df['Suppliers'] = df['Suppliers'].fillna('غير محدد')

    df['Beginning_Inventory'] = df['Quantity On Hand'] - df['Total_Purchases_Period'] + df['Total_Sales_Period']
    df['Average_Inventory'] = (df['Beginning_Inventory'] + df['Quantity On Hand']) / 2
    df['Average_Inventory'] = df['Average_Inventory'].clip(lower=1)

    # --- تحويل معدل الدوران إلى سنوي بضربه في 3 (لأن الفترة 4 أشهر) ---
    df['Quantity_Turnover_Rate'] = (df['Total_Sales_Period'] / df['Average_Inventory']) * 3
    df['Invoice_Turnover_Rate'] = (df['Invoice_Count'] * 3) / df['Average_Inventory']
    
    # --- تعديل عتبات التصنيف لتناسب المعدلات السنوية الجديدة ---
    def classify_quantity_turnover(rate):
        if rate >= 18: return 'سريع جداً'
        elif rate >= 9: return 'سريع'
        elif rate >= 4: return 'متوسط'
        elif rate >= 1: return 'بطيء'
        else: return 'راكد'
    
    def classify_invoice_turnover(rate):
        if rate >= 24: return 'عالي التكرار'
        elif rate >= 12: return 'متوسط التكرار'
        elif rate >= 6: return 'منخفض التكرار'
        else: return 'نادر البيع'
    
    df['Quantity_Turnover_Classification'] = df['Quantity_Turnover_Rate'].apply(classify_quantity_turnover)
    df['Invoice_Turnover_Classification'] = df['Invoice_Turnover_Rate'].apply(classify_invoice_turnover)

    df['Safety_Stock'] = (df['Average_Monthly_Sales'] * safety_stock_days) / 30
    df['Days_Of_Stock'] = np.where(df['Average_Monthly_Sales'] > 0, (df['Quantity On Hand'] / df['Average_Monthly_Sales']) * 30, 999)
    df['Recommended_Purchase'] = np.maximum((df['Average_Monthly_Sales'] + df['Safety_Stock']) - df['Quantity On Hand'], 0)
    
    def calculate_priority(row):
        if row['Days_Of_Stock'] <= 7: return 'عاجل جداً'
        elif row['Days_Of_Stock'] <= 15: return 'عاجل'
        elif row['Days_Of_Stock'] <= 30: return 'متوسط'
        else: return 'منخفض'
    
    df['Priority'] = df.apply(calculate_priority, axis=1)
    
    df['Total_Cost'] = df['Recommended_Purchase'] * df['Cost'] if 'Cost' in df.columns else 0
    
    result_columns = [
        'Barcode', 'Name', 'Product Category/Complete Name', 'Quantity On Hand', 'Average_Monthly_Sales', 
        'Days_Of_Stock', 'Quantity_Turnover_Rate', 'Invoice_Turnover_Rate', 'Quantity_Turnover_Classification', 
        'Invoice_Turnover_Classification', 'Priority', 'Recommended_Purchase', 'Total_Cost', 'Suppliers'
    ]
    available_columns = [col for col in result_columns if col in df.columns]
    result_df = df[available_columns]
    
    arabic_names = {
        'Barcode': 'الباركود', 'Name': 'اسم المنتج', 'Product Category/Complete Name': 'فئة المنتج',
        'Quantity On Hand': 'الكمية المتاحة', 'Average_Monthly_Sales': 'متوسط المبيعات الشهرية',
        'Days_Of_Stock': 'أيام التغطية', 'Quantity_Turnover_Rate': 'معدل دوران الكميات',
        'Invoice_Turnover_Rate': 'معدل دوران الفواتير', 'Quantity_Turnover_Classification': 'تصنيف دوران الكميات',
        'Invoice_Turnover_Classification': 'تصنيف دوران الفواتير', 'Priority': 'الأولوية',
        'Recommended_Purchase': 'الشراء المقترح', 'Total_Cost': 'التكلفة الإجمالية', 'Suppliers': 'الموردين'
    }
    return result_df.rename(columns=arabic_names)

# Visualization functions
def create_turnover_charts(df):
    fig = make_subplots(rows=1, cols=2, subplot_titles=["تحليل دوران الكميات (سنوي)", "تحليل دوران الفواتير (سنوي)"], specs=[[{"type": "pie"}, {"type": "pie"}]])
    quantity_counts = df['تصنيف دوران الكميات'].value_counts()
    fig.add_trace(go.Pie(labels=quantity_counts.index, values=quantity_counts.values, name="دوران الكميات"), row=1, col=1)
    invoice_counts = df['تصنيف دوران الفواتير'].value_counts()
    fig.add_trace(go.Pie(labels=invoice_counts.index, values=invoice_counts.values, name="دوران الفواتير"), row=1, col=2)
    fig.update_traces(textposition='inside', textinfo='percent+label', hole=.3)
    fig.update_layout(title_text="📊 تحليل معدلات دوران المخزون السنوية", font=dict(size=12), height=450, showlegend=False)
    return fig

def create_combined_analysis_chart(df):
    df_filtered = df[df['الشراء المقترح'] > 0]
    if df_filtered.empty: return None
    fig = px.scatter(df_filtered, x='معدل دوران الكميات', y='معدل دوران الفواتير', size='الشراء المقترح', color='الأولوية',
                     hover_data=['اسم المنتج', 'الكمية المتاحة'], title="تحليل مقارن للمنتجات المقترحة",
                     labels={'معدل دوران الكميات': 'معدل دوران الكميات (بطيء -> سريع)', 'معدل دوران الفواتير': 'معدل دوران الفواتير (نادر -> متكرر)'},
                     color_discrete_map={'عاجل جداً': '#FF4444', 'عاجل': '#FF8800', 'متوسط': '#FFAA00', 'منخفض': '#00AA00'})
    fig.update_layout(height=500)
    return fig

def create_priority_chart(df):
    priority_counts = df['الأولوية'].value_counts()
    colors = {'عاجل جداً': '#FF4444', 'عاجل': '#FF8800', 'متوسط': '#FFAA00', 'منخفض': '#00AA00'}
    fig = px.bar(x=priority_counts.index, y=priority_counts.values, title="توزيع المنتجات حسب الأولوية",
                 color=priority_counts.index, color_discrete_map=colors, labels={'x':'الأولوية', 'y':'عدد المنتجات'})
    fig.update_layout(showlegend=False, height=400)
    return fig

def create_stock_days_chart(df):
    df_filtered = df[df['الشراء المقترح'] > 0]
    if df_filtered.empty: return None
    fig = px.histogram(df_filtered, x='أيام التغطية', nbins=30, title="توزيع أيام التغطية للمنتجات المطلوب شراؤها")
    fig.add_vline(x=df_filtered['أيام التغطية'].median(), line_dash="dash", line_color="red", annotation_text=f"الوسيط: {df_filtered['أيام التغطية'].median():.1f} يوم")
    fig.update_layout(xaxis_title="أيام التغطية", yaxis_title="عدد المنتجات", height=400)
    return fig

# Main Streamlit application
def main():
    st.title("🛒 نظام اقتراح المشتريات المتقدم")
    st.markdown("---")
    
    with st.sidebar:
        st.header("⚙️ إعدادات النظام")
        target_month = st.selectbox("اختر الشهر المستهدف", options=list(range(1, 13)), index=datetime.now().month - 1,
                                    format_func=lambda x: ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"][x-1])
        target_year = st.number_input("أدخل السنة المستهدفة", value=datetime.now().year, min_value=2020, max_value=2030)
        safety_stock_days = st.slider("أيام مخزون الأمان", min_value=7, max_value=90, value=30, help="عدد الأيام الإضافية للحماية من نفاد المخزون")
        st.markdown("---")
        st.info("**يتم حساب الشراء بناءً على:**\n- الشهر السابق للشهر المختار\n- 3 شهور مقابلة في السنة السابقة")
    
    with st.spinner("جاري تحميل البيانات..."):
        sales, stock, purchases = load_data()
    
    if sales is None or stock is None or purchases is None:
        return
    
    col1, col2, col3 = st.columns(3)
    col1.metric("📈 سجلات المبيعات", f"{len(sales):,}")
    col2.metric("📦 منتجات المخزون", f"{len(stock):,}")
    col3.metric("🛍️ سجلات المشتريات", f"{len(purchases):,}")
    st.markdown("---")
    
    if st.button("🔄 توليد خطة الشراء", type="primary", use_container_width=True):
        with st.spinner("جاري توليد خطة الشراء..."):
            try:
                plan = generate_plan(sales, stock, purchases, target_month, target_year, safety_stock_days)
                if plan.empty:
                    st.warning("⚠️ لم يتم إنشاء خطة شراء. تحقق من البيانات.")
                    return
                
                st.success("✅ تم توليد خطة الشراء بنجاح!")
                st.session_state.plan = plan

            except Exception as e:
                st.error(f"❌ خطأ في توليد خطة الشراء: {str(e)}")
                st.exception(e)
                return

    if 'plan' in st.session_state:
        plan = st.session_state.plan
        st.subheader("📊 مؤشرات رئيسية للخطة")
        total_recommended = plan['الشراء المقترح'].sum()
        products_to_buy = len(plan[plan['الشراء المقترح'] > 0])
        avg_quantity_turnover = plan['معدل دوران الكميات'][plan['معدل دوران الكميات'] > 0].mean()
        avg_invoice_turnover = plan['معدل دوران الفواتير'][plan['معدل دوران الفواتير'] > 0].mean()
        urgent_products = len(plan[plan['الأولوية'].isin(['عاجل', 'عاجل جداً'])])
        
        kpi_cols = st.columns(5)
        kpi_cols[0].metric("إجمالي القطع المقترحة", f"{total_recommended:,.0f}")
        kpi_cols[1].metric("المنتجات المطلوب شراؤها", f"{products_to_buy:,}")
        kpi_cols[2].metric("متوسط دوران الكميات", f"{avg_quantity_turnover:.2f}")
        kpi_cols[3].metric("متوسط دوران الفواتير", f"{avg_invoice_turnover:.2f}")
        kpi_cols[4].metric("المنتجات العاجلة", f"{urgent_products:,}")
        
        st.subheader("📈 التحليل البصري")
        st.plotly_chart(create_turnover_charts(plan), use_container_width=True)
        
        chart_cols = st.columns(2)
        with chart_cols[0]: st.plotly_chart(create_priority_chart(plan), use_container_width=True)
        with chart_cols[1]: 
            combined_chart = create_combined_analysis_chart(plan)
            if combined_chart: st.plotly_chart(combined_chart, use_container_width=True)
        
        stock_chart = create_stock_days_chart(plan)
        if stock_chart: st.plotly_chart(stock_chart, use_container_width=True)
        
        st.subheader("🔍 تصفية وعرض النتائج")
        filter_cols = st.columns(4)
        quantity_turnover_filter = filter_cols[0].multiselect("تصنيف دوران الكميات:", options=plan['تصنيف دوران الكميات'].unique(), default=plan['تصنيف دوران الكميات'].unique())
        invoice_turnover_filter = filter_cols[1].multiselect("تصنيف دوران الفواتير:", options=plan['تصنيف دوران الفواتير'].unique(), default=plan['تصنيف دوران الفواتير'].unique())
        priority_filter = filter_cols[2].multiselect("الأولوية:", options=plan['الأولوية'].unique(), default=plan['الأولوية'].unique())
        min_purchase = filter_cols[3].number_input("الحد الأدنى للشراء:", min_value=0.0, value=1.0, step=1.0)
        
        filtered_plan = plan[
            (plan['تصنيف دوران الكميات'].isin(quantity_turnover_filter)) &
            (plan['تصنيف دوران الفواتير'].isin(invoice_turnover_filter)) &
            (plan['الأولوية'].isin(priority_filter)) &
            (plan['الشراء المقترح'] >= min_purchase)
        ]
        
        st.subheader(f"📋 خطة الشراء المفصلة ({len(filtered_plan)} منتج)")
        st.dataframe(
            filtered_plan,
            use_container_width=True,
            hide_index=True,
            column_config={
                "الأولوية": st.column_config.SelectboxColumn("الأولوية", options=["عاجل جداً", "عاجل", "متوسط", "منخفض"]),
                "الشراء المقترح": st.column_config.NumberColumn("الشراء المقترح", format="%.0f"),
                "معدل دوران الكميات": st.column_config.NumberColumn("معدل دوران الكميات", help="المعدل السنوي لدوران المخزون بناءً على الكميات المباعة.", format="%.2f"),
                "معدل دوران الفواتير": st.column_config.NumberColumn("معدل دوران الفواتير", help="المعدل السنوي لدوران المخزون بناءً على تكرار البيع.", format="%.2f"),
                "الكمية المتاحة": st.column_config.NumberColumn(format="%.0f"),
                "متوسط المبيعات الشهرية": st.column_config.NumberColumn(format="%.1f"),
                "أيام التغطية": st.column_config.NumberColumn(help="عدد الأيام التي تكفيها الكمية الحالية", format="%.1f يوم"),
                "التكلفة الإجمالية": st.column_config.NumberColumn("التكلفة الإجمالية", format="EGP%.2f")
            }
        )

        if 'فئة المنتج' in filtered_plan.columns:
            st.subheader("📋 ملخص حسب الفئة")
            category_summary = filtered_plan.groupby('فئة المنتج').agg(إجمالي_الشراء_المقترح=('الشراء المقترح', 'sum'), عدد_المنتجات=('الباركود', 'count')).round(0).sort_values('إجمالي_الشراء_المقترح', ascending=False)
            st.dataframe(category_summary, use_container_width=True)
        
        st.subheader("📥 تحميل النتائج")
        download_cols = st.columns(2)
        excel_file = to_excel(filtered_plan)
        download_cols[0].download_button(label="📊 تحميل الخطة المفلترة (Excel)", data=excel_file, file_name=f"purchase_plan_{target_year}_{target_month:02d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        
        urgent_products = filtered_plan[filtered_plan['الأولوية'].isin(['عاجل', 'عاجل جداً'])]
        if not urgent_products.empty:
            urgent_excel = to_excel(urgent_products)
            download_cols[1].download_button(label="🚨 تحميل المنتجات العاجلة فقط", data=urgent_excel, file_name=f"urgent_purchases_{target_year}_{target_month:02d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

if __name__ == "__main__":
    main()
