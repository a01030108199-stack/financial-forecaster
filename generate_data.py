import pandas as pd
import numpy as np
from datetime import datetime

# دالة لتوليد البيانات المالية التاريخية
def generate_financial_dataset():
    # 1. إعداد النطاق الزمني: 5 سنوات (من يناير 2021 إلى ديسمبر 2025 = 60 شهراً)
    date_range = pd.date_range(start="2021-01-01", end="2025-12-01", freq="MS")
    
    # تهيئة قوائم فارغة لتخزين البيانات
    data = []
    
    # تحديد القيمة الابتدائية للإيرادات في يناير 2021 (مثلاً 100,000 جنيه)
    base_revenue = 100000
    
    # مولد أرقام عشوائية ثابت للحصول على نفس النتائج عند تشغيل الكود
    np.random.seed(42)
    
    for i, date in enumerate(date_range):
        # أ) احتساب النمو السنوي التدريجي (حوالي 1% نمو شهري تراكمي)
        growth_factor = 1 + (i * 0.012)
        
        # ب) احتساب الموسمية: زيادة المبيعات في الصيف (شهور 6 و 7 و 8) ونهاية السنة (شهر 12)
        # نستخدم دالة جيبية (Sine Wave) لمحاكاة تغيرات الفصول السنوية
        month = date.month
        seasonality = 1 + 0.15 * np.sin(2 * np.pi * month / 12)
        
        # ج) إضافة تقلبات عشوائية طبيعية (Noise) بنسبة +/- 5% لجعل البيانات واقعية وليست خطاً مستقيماً
        random_noise = np.random.uniform(0.95, 1.05)
        
        # احتساب الإيرادات النهائية للشهر
        revenue = int(base_revenue * growth_factor * seasonality * random_noise)
        
        # د) احتساب تكلفة المبيعات (COGS): تمثل عادة حوالي 50% من الإيرادات مع تغير طفيف عشوائي
        cogs_percentage = np.random.uniform(0.48, 0.52)
        cogs = int(revenue * cogs_percentage)
        
        # هـ) احتساب المصاريف التشغيلية (Operating Expenses): جزء ثابت (20,000) وجزء متغير (5% من الإيرادات)
        op_expenses = int(20000 + (revenue * 0.05) + np.random.uniform(-1000, 1000))
        
        # و) احتساب المصاريف التسويقية (Marketing Expenses): جزء ثابت (8,000) وجزء متغير (4% من الإيرادات)
        marketing_expenses = int(8000 + (revenue * 0.04) + np.random.uniform(-500, 500))
        
        # ز) احتساب صافي الربح (Net Profit) = الإيرادات - تكلفة المبيعات - المصاريف التشغيلية - المصاريف التسويقية
        net_profit = revenue - cogs - op_expenses - marketing_expenses
        
        # إضافة السجل المالي لهذا الشهر إلى القائمة
        data.append({
            "التاريخ": date.strftime("%Y-%m-%d"),
            "الإيرادات": revenue,
            "تكلفة المبيعات": cogs,
            "المصاريف التشغيلية": op_expenses,
            "المصاريف التسويقية": marketing_expenses,
            "صافي الربح": net_profit
        })
    
    # 2. تحويل البيانات إلى جدول Pandas DataFrame
    df = pd.DataFrame(data)
    
    # 3. حفظ الجدول في ملف Excel باسم financial_data.xlsx
    output_filename = "financial_data.xlsx"
    df.to_excel(output_filename, index=False)
    
    print(f"Success: Data generated and saved to: {output_filename}")
    print(f"Total months: {len(df)} (from {df['التاريخ'].iloc[0]} to {df['التاريخ'].iloc[-1]})")

if __name__ == "__main__":
    generate_financial_dataset()
