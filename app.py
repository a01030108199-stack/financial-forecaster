import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from sklearn.linear_model import LinearRegression
import os
import base64

# 1. إعدادات الصفحة العامة وتحسين الواجهة
st.set_page_config(
    page_title="نظام التنبؤ المالي والتحليل الذكي - أكرم سعد",
    page_icon="📊",
    layout="wide"
)

# دالة لتحويل الصور المحلية إلى ترميز Base64 لدمجها في HTML
def get_image_base64(path):
    if os.path.exists(path):
        try:
            with open(path, "rb") as image_file:
                return base64.b64encode(image_file.read()).decode()
        except Exception:
            return ""
    return ""

avatar_b64 = get_image_base64("avatar.png")
logo_b64 = get_image_base64("logo.png")

# دعم اللغة العربية ومحاذاة النصوص من اليمين لليسار (RTL) وتصميم الكروت الإحصائية والهوية البصرية
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;500;600;700;800&display=swap');
        
        html, body, [data-testid="stSidebar"], .main, h1, h2, h3, h4, h5, h6, p, span, label, button, select, input {
            font-family: 'Cairo', sans-serif !important;
            direction: rtl !important;
            text-align: right !important;
        }
        
        /* بانر الهوية والاسم */
        .branding-banner {
            background: linear-gradient(135deg, #0d1527 0%, #070a13 100%);
            border: 1px solid rgba(139, 92, 246, 0.15);
            border-right: 6px solid #8b5cf6;
            border-radius: 16px;
            padding: 20px 30px;
            margin-bottom: 30px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
        }
        
        /* تصميم مخصص للكروت الإحصائية */
        .metric-card {
            background: rgba(17, 24, 39, 0.65);
            border: 1px solid rgba(255, 255, 255, 0.05);
            border-radius: 16px;
            padding: 22px 15px;
            text-align: center !important;
            box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.2);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            margin-bottom: 20px;
        }
        
        .metric-card:hover {
            transform: translateY(-5px);
            border-color: rgba(139, 92, 246, 0.35);
            box-shadow: 0 12px 40px 0 rgba(139, 92, 246, 0.15);
            background: rgba(17, 24, 39, 0.85);
        }
        
        .metric-title {
            font-size: 14px;
            color: #9ca3af;
            margin-bottom: 10px;
            font-weight: 500;
        }
        
        .metric-value {
            font-size: 26px;
            font-weight: 800;
            margin-bottom: 6px;
        }
        
        .metric-subtitle {
            font-size: 12px;
            color: #9ca3af;
            font-weight: 400;
        }
        
        /* تعديل محاذاة عناصر التصفية الجانبية */
        [data-testid="stSidebar"] {
            text-align: right !important;
            background-color: #0b0f19 !important;
            border-left: 1px solid rgba(255, 255, 255, 0.05);
        }
        
        /* تحسين مظهر الأزرار */
        .stButton>button {
            width: 100%;
            background: linear-gradient(135deg, #8b5cf6, #4f46e5) !important;
            color: white !important;
            font-weight: 700 !important;
            font-size: 16px !important;
            border: none !important;
            border-radius: 10px !important;
            padding: 10px 20px !important;
            box-shadow: 0 4px 15px rgba(139, 92, 246, 0.2) !important;
            transition: all 0.3s ease !important;
        }

        .stButton>button:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 6px 20px rgba(139, 92, 246, 0.35) !important;
            background: linear-gradient(135deg, #a78bfa, #6366f1) !important;
        }

        /* تنسيق التبويبات Tabs */
        .stTabs [data-baseweb="tab-list"] {
            gap: 24px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.05);
        }

        .stTabs [data-baseweb="tab"] {
            height: 50px;
            background-color: transparent !important;
            border: none !important;
            color: #9ca3af !important;
            font-weight: 600 !important;
            font-size: 16px !important;
            padding: 0 8px !important;
        }

        .stTabs [data-baseweb="tab"]:hover {
            color: #ffffff !important;
        }

        .stTabs [aria-selected="true"] {
            color: #8b5cf6 !important;
            border-bottom: 2px solid #8b5cf6 !important;
        }
    </style>
""", unsafe_allow_html=True)

# 2. دالة تحميل البيانات المالية من ملف Excel
@st.cache_data
def load_financial_data():
    file_path = "financial_data.xlsx"
    if not os.path.exists(file_path):
        return None
    
    # قراءة الملف وتحويل التاريخ لعمود زمني
    df = pd.read_excel(file_path)
    df['التاريخ'] = pd.to_datetime(df['التاريخ'])
    return df

df = load_financial_data()

if df is None:
    st.error("❌ لم يتم العثور على ملف البيانات 'financial_data.xlsx'. يرجى تشغيل كود توليد البيانات أولاً.")
    st.stop()

# 3. شريط التحكم الجانبي (Sidebar)
if logo_b64:
    st.sidebar.markdown(f"""
        <div style="text-align: center; margin-bottom: 15px; padding-top: 10px;">
            <img src="data:image/png;base64,{logo_b64}" style="width: 120px; max-width: 80%;" />
        </div>
    """, unsafe_allow_html=True)

st.sidebar.markdown("<h3 style='text-align: center; color: #8b5cf6; margin-top: 0; font-size: 18px;'>🎛️ لوحة التحكم والتحليل</h3>", unsafe_allow_html=True)
st.sidebar.markdown("<hr style='margin: 10px 0; opacity: 0.15;'>", unsafe_allow_html=True)

# فلترة السنوات
years = sorted(df['التاريخ'].dt.year.unique())
selected_years = st.sidebar.multiselect("📅 اختر السنوات للتحليل:", years, default=years)

# فلترة البيانات بناءً على الاختيار
filtered_df = df[df['التاريخ'].dt.year.isin(selected_years)].sort_values(by='التاريخ')

# محاكاة السيناريوهات المستقبلية (What-If Analysis)
st.sidebar.markdown("<h3 style='color: #06b6d4;'>🧠 محاكاة السيناريوهات لعام 2026</h3>", unsafe_allow_html=True)
marketing_multiplier = st.sidebar.slider("📢 زيادة الاستثمار التسويقي (%):", 0, 100, 10, step=5) / 100
opex_savings = st.sidebar.slider("✂️ خفض المصاريف التشغيلية (%):", 0, 50, 5, step=5) / 100

# 4. واجهة التطبيق الرئيسية (Main Layout)
logo_html = f'<img src="data:image/png;base64,{logo_b64}" style="height: 55px;" />' if logo_b64 else '📊'
avatar_html = f'<img src="data:image/png;base64,{avatar_b64}" style="width: 50px; height: 50px; border-radius: 50%; border: 2px solid #8b5cf6; object-fit: cover;" />' if avatar_b64 else '👤'

st.markdown(f"""
    <div class="branding-banner">
        <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 20px;">
            <div style="display: flex; align-items: center; gap: 15px; min-width: 300px;">
                <div style="font-size: 40px; line-height: 1;">{logo_html}</div>
                <div>
                    <h1 style="margin: 0; font-size: 24px; color: #ffffff; font-weight: 700; line-height: 1.2;">نظام التحليل المالي والتنبؤ الذكي بالذكاء الاصطناعي</h1>
                    <p style="margin: 4px 0 0 0; color: #9ca3af; font-size: 13px;">منصة متطورة لتحليل الأداء المالي التاريخي والتنبؤ بالموازنات والمبيعات المستقبلية</p>
                </div>
            </div>
            <div style="display: flex; align-items: center; gap: 12px; background: rgba(255, 255, 255, 0.02); padding: 8px 16px; border-radius: 12px; border: 1px solid rgba(255, 255, 255, 0.05); min-width: 250px;">
                {avatar_html}
                <div style="text-align: right;">
                    <h4 style="margin: 0; color: #a78bfa; font-size: 15px; font-weight: 700;">أكرم سعد</h4>
                    <p style="margin: 2px 0 0 0; color: #38bdf8; font-size: 11px; font-weight: 600; letter-spacing: 0.5px;">محاسب مالي • برمجة • تحليل مالي</p>
                </div>
            </div>
        </div>
    </div>
""", unsafe_allow_html=True)

# صف الكروت الإحصائية الرئيسية للبيانات المفلترة
total_revenue = filtered_df['الإيرادات'].sum()
total_cogs = filtered_df['تكلفة المبيعات'].sum()
total_opex = filtered_df['المصاريف التشغيلية'].sum()
total_marketing = filtered_df['المصاريف التسويقية'].sum()
total_net_profit = filtered_df['صافي الربح'].sum()
avg_margin = (total_net_profit / total_revenue) * 100 if total_revenue > 0 else 0

col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.markdown(f"""
    <div class="metric-card" style="border-top: 4px solid #34d399;">
        <div class="metric-title">📈 إجمالي الإيرادات</div>
        <div class="metric-value" style="color: #34d399;">{total_revenue:,.0f} ج.م</div>
        <div class="metric-subtitle">خلال الفترة المحددة</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="metric-card" style="border-top: 4px solid #f43f5e;">
        <div class="metric-title">📉 تكلفة المبيعات (COGS)</div>
        <div class="metric-value" style="color: #f43f5e;">{total_cogs:,.0f} ج.م</div>
        <div class="metric-subtitle">{(total_cogs/total_revenue)*100:.1f}% من الإيرادات</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="metric-card" style="border-top: 4px solid #fb923c;">
        <div class="metric-title">💸 المصاريف التشغيلية</div>
        <div class="metric-value" style="color: #fb923c;">{total_opex:,.0f} ج.م</div>
        <div class="metric-subtitle">{(total_opex/total_revenue)*100:.1f}% من الإيرادات</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown(f"""
    <div class="metric-card" style="border-top: 4px solid #10b981;">
        <div class="metric-title">💵 صافي الأرباح</div>
        <div class="metric-value" style="color: #10b981;">{total_net_profit:,.0f} ج.م</div>
        <div class="metric-subtitle">إجمالي القيمة الصافية</div>
    </div>
    """, unsafe_allow_html=True)

with col5:
    st.markdown(f"""
    <div class="metric-card" style="border-top: 4px solid #38bdf8;">
        <div class="metric-title">🎯 هامش صافي الربح</div>
        <div class="metric-value" style="color: #38bdf8;">{avg_margin:.2f}%</div>
        <div class="metric-subtitle">كفاءة الأداء المالي</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# 5. المخططات البيانية للتحليل التاريخي
tab1, tab2 = st.tabs(["📈 تحليل الاتجاهات المالية", "🍰 هيكل التكاليف والمصاريف"])

with tab1:
    st.markdown("<h3 style='color: #8b5cf6; font-size: 20px; font-weight: 700;'>حركة الإيرادات مقارنة بصافي الربح</h3>", unsafe_allow_html=True)
    # رسم خطي تفاعلي للإيرادات والأرباح
    fig_trends = go.Figure()
    fig_trends.add_trace(go.Scatter(
        x=filtered_df['التاريخ'], y=filtered_df['الإيرادات'],
        mode='lines+markers', name='الإيرادات',
        line=dict(color='#8b5cf6', width=3, shape='spline'),
        marker=dict(size=6, color='#8b5cf6'),
        hovertemplate='الإيرادات: %{y:,.0f} ج.م<extra></extra>'
    ))
    fig_trends.add_trace(go.Scatter(
        x=filtered_df['التاريخ'], y=filtered_df['صافي الربح'],
        mode='lines+markers', name='صافي الربح',
        line=dict(color='#10b981', width=3, shape='spline'),
        marker=dict(size=6, color='#10b981'),
        hovertemplate='صافي الربح: %{y:,.0f} ج.م<extra></extra>'
    ))
    fig_trends.update_layout(
        template='plotly_dark',
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        xaxis=dict(showgrid=False, title='التاريخ'),
        yaxis=dict(showgrid=True, gridcolor='rgba(255,255,255,0.05)', title='القيمة (ج.م)'),
        margin=dict(l=20, r=20, t=30, b=20),
        legend=dict(orientation="h", y=1.1, x=1, xanchor='right'),
        hovermode="x unified"
    )
    st.plotly_chart(fig_trends, width='stretch')

with tab2:
    st.markdown("<h3 style='color: #8b5cf6; font-size: 20px; font-weight: 700;'>توزيع هيكل التكاليف والمصاريف الإجمالي</h3>", unsafe_allow_html=True)
    cost_labels = ['تكلفة المبيعات', 'المصاريف التشغيلية', 'المصاريف التسويقية']
    cost_values = [total_cogs, total_opex, total_marketing]
    
    fig_pie = go.Figure(data=[go.Pie(
        labels=cost_labels, 
        values=cost_values, 
        hole=.4,
        marker=dict(colors=['#f43f5e', '#fb923c', '#3b82f6']),
        hovertemplate='%{label}: %{value:,.0f} ج.م (%{percent})<extra></extra>'
    )])
    fig_pie.update_layout(
        template='plotly_dark',
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=20, r=20, t=30, b=20),
        legend=dict(orientation="h", y=1.1, x=1, xanchor='right')
    )
    st.plotly_chart(fig_pie, width='stretch')

st.markdown("<hr>", unsafe_allow_html=True)

# 6. التنبؤ بالذكاء الاصطناعي لعام 2026 (AI Forecasting Engine)
st.markdown("<h2 style='color: #8b5cf6;'>🔮 التنبؤ المالي الذكي لعام 2026</h2>", unsafe_allow_html=True)
st.markdown("<p style='color: #9e9aa9;'>يستخدم النظام نموذج الانحدار الخطي (Linear Regression) لتحليل الاتجاه التاريخي للمبيعات وبناء خطة التنبؤ لـ 12 شهراً مستقبلياً.</p>", unsafe_allow_html=True)

# أ) تجهيز البيانات لنموذج الذكاء الاصطناعي
# نستخدم رقم الشهر كمتغير مستقل (X) لتوقع قيمة الإيرادات (y)
df['Month_Index'] = np.arange(len(df))

X_train = df[['Month_Index']].values
y_revenue = df['الإيرادات'].values

# تدريب النموذج على الإيرادات
model_revenue = LinearRegression()
model_revenue.fit(X_train, y_revenue)

# ب) توليد تواريخ 12 شهراً القادمة (عام 2026)
future_months = pd.date_range(start="2026-01-01", end="2026-12-01", freq="MS")
future_indices = np.arange(len(df), len(df) + 12).reshape(-1, 1)

# توقع الإيرادات لعام 2026
predicted_revenue = model_revenue.predict(future_indices)

# ج) احتساب المصاريف المتوقعة لعام 2026 بناءً على النسب التاريخية وتعديل محاكاة المستخدم (What-If Simulation)
# تكلفة المبيعات التاريخية كنسبة من الإيرادات (المتوسط)
avg_cogs_ratio = (df['تكلفة المبيعات'] / df['الإيرادات']).mean()

# احتساب المصاريف لعام 2026 مع تطبيق مدخلات شريط التحكم الجانبي
predicted_cogs = predicted_revenue * avg_cogs_ratio

# المصاريف التشغيلية المتوقعة مع تطبيق الخصم المستهدف من المستخدم
avg_opex_ratio = (df['المصاريف التشغيلية'] / df['الإيرادات']).mean()
predicted_opex = (predicted_revenue * avg_opex_ratio) * (1 - opex_savings)

# المصاريف التسويقية المتوقعة مع تطبيق الزيادة المستهدفة من المستخدم
avg_mkt_ratio = (df['المصاريف التسويقية'] / df['الإيرادات']).mean()
predicted_marketing = (predicted_revenue * avg_mkt_ratio) * (1 + marketing_multiplier)

# صافي الربح المتوقع لعام 2026 = الإيرادات المتوقعة - إجمالي المصاريف المتوقعة
predicted_net_profit = predicted_revenue - predicted_cogs - predicted_opex - predicted_marketing

# د) بناء جدول التوقعات للعام الجديد
forecast_data = pd.DataFrame({
    'التاريخ': future_months,
    'الإيرادات المتوقعة': predicted_revenue.astype(int),
    'تكلفة المبيعات المتوقعة': predicted_cogs.astype(int),
    'المصاريف التشغيلية المتوقعة': predicted_opex.astype(int),
    'المصاريف التسويقية المتوقعة': predicted_marketing.astype(int),
    'صافي الربح المتوقع': predicted_net_profit.astype(int)
})

# هـ) رسم منحنى التنبؤ التفاعلي ومقارنته بالبيانات التاريخية
fig_forecast = go.Figure()

# البيانات التاريخية (إيرادات وصافي ربح)
fig_forecast.add_trace(go.Scatter(
    x=df['التاريخ'], y=df['الإيرادات'],
    mode='lines', name='الإيرادات التاريخية',
    line=dict(color='#8b5cf6', width=2, shape='spline'),
    hovertemplate='الإيرادات التاريخية: %{y:,.0f} ج.م<extra></extra>'
))
fig_forecast.add_trace(go.Scatter(
    x=df['التاريخ'], y=df['صافي الربح'],
    mode='lines', name='صافي الربح التاريخي',
    line=dict(color='#10b981', width=2, shape='spline'),
    hovertemplate='صافي الربح التاريخي: %{y:,.0f} ج.م<extra></extra>'
))

# البيانات المتوقعة لعام 2026 (خط متقطع dashed للإشارة للتوقع)
fig_forecast.add_trace(go.Scatter(
    x=forecast_data['التاريخ'], y=forecast_data['الإيرادات المتوقعة'],
    mode='lines+markers', name='الإيرادات المتوقعة (2026)',
    line=dict(color='#c084fc', width=3, dash='dash', shape='spline'),
    marker=dict(size=5),
    hovertemplate='الإيرادات المتوقعة: %{y:,.0f} ج.م<extra></extra>'
))
fig_forecast.add_trace(go.Scatter(
    x=forecast_data['التاريخ'], y=forecast_data['صافي الربح المتوقع'],
    mode='lines+markers', name='صافي الربح المتوقع (2026)',
    line=dict(color='#34d399', width=3, dash='dash', shape='spline'),
    marker=dict(size=5),
    hovertemplate='صافي الربح المتوقع: %{y:,.0f} ج.م<extra></extra>'
))

fig_forecast.update_layout(
    template='plotly_dark',
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(0,0,0,0)',
    xaxis=dict(showgrid=False, title='التاريخ'),
    yaxis=dict(showgrid=True, gridcolor='rgba(255,255,255,0.05)', title='القيمة (ج.م)'),
    margin=dict(l=20, r=20, t=30, b=20),
    legend=dict(orientation="h", y=1.1, x=1, xanchor='right'),
    hovermode="x unified"
)

st.plotly_chart(fig_forecast, width='stretch')

# و) عرض الأرقام المتوقعة في جدول منسق
st.markdown("<h3 style='color: #06b6d4;'>📋 جدول التنبؤات والتحليلات المتوقعة لعام 2026</h3>", unsafe_allow_html=True)
display_forecast = forecast_data.copy()
display_forecast['التاريخ'] = display_forecast['التاريخ'].dt.strftime('%Y-%m-%d')
st.dataframe(
    display_forecast.style.format({
        'الإيرادات المتوقعة': '{:,.0f} ج.م',
        'تكلفة المبيعات المتوقعة': '{:,.0f} ج.م',
        'المصاريف التشغيلية المتوقعة': '{:,.0f} ج.م',
        'المصاريف التسويقية المتوقعة': '{:,.0f} ج.م',
        'صافي الربح المتوقع': '{:,.0f} ج.م'
    }),
    width='stretch'
)

st.markdown("<br>", unsafe_allow_html=True)
st.info("💡 نصيحة: يمكنك استخدام لوحة التحكم في القائمة الجانبية لتعديل زيادة التسويق أو خفض المصاريف ورؤية كيف يتغير صافي الربح المتوقع لعام 2026 فورياً في المخطط والجدول!")
