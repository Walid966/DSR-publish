import pandas as pd
import os
from pandas import ExcelWriter
import calendar

# تحميل الملف
file_path = "تقرير الفواتير.xlsx"
sheet_name = 0  # استخدام أول شيت في الملف تلقائيًا

# اقرأ البيانات
print("جاري قراءة البيانات...")
df = pd.read_excel(file_path, sheet_name=sheet_name)

# إنشاء الجدول المحوري المجمع
pivot = pd.pivot_table(
    df,
    index=[
        "MandoubCode",   # رقم المندوب
        "MandoubName",   # اسم المندوب
        "UserCode",      # رقم التاجر
        "ShopName",      # اسم التاجر
    ],
    columns="CategoryName",  # نوع الخدمة كأعمدة
    values="BillTotal",   # المبلغ
    aggfunc="sum",
    fill_value=0
).reset_index()

# استخراج آخر تاريخ من البيانات الأصلية
last_date_str = str(df["Date"].iloc[-1])
try:
    # معالجة التاريخ العربي مثل "الأربعاء, 24 سبتمبر 2025 00:00:14"
    # استخراج اليوم من النص
    import re
    match = re.search(r"(\d{1,2})\s+\w+\s+\d{4}", last_date_str)
    if match:
        day = int(match.group(1))
        # استخراج اسم الشهر والسنة
        month_name = last_date_str.split()[2]
        year = int(last_date_str.split()[3])
        arabic_months = {
            "يناير": 1, "فبراير": 2, "مارس": 3, "ابريل": 4, "أبريل": 4, "مايو": 5, "يونيو": 6,
            "يوليو": 7, "اغسطس": 8, "أغسطس": 8, "سبتمبر": 9, "اكتوبر": 10, "أكتوبر": 10,
            "نوفمبر": 11, "ديسمبر": 12
        }
        month = arabic_months.get(month_name, 0)
        if month == 0:
            raise ValueError("اسم الشهر غير معروف: " + month_name)
    else:
        raise ValueError("لم يتم العثور على اليوم في نص التاريخ: " + last_date_str)
except Exception as e:
    print("خطأ في استخراج التاريخ: ", e)
    day = 1
    month = 1
    year = 2020

# حساب عدد أيام الشهر الكامل
num_days_in_month = calendar.monthrange(year, month)[1]

# Identify service columns (CategoryName values)
service_columns_pivot = [col for col in pivot.columns if col not in ["MandoubCode", "MandoubName", "UserCode", "ShopName"]]

# Calculate BillTotal for each row by summing across service columns
pivot["BillTotal"] = pivot[service_columns_pivot].sum(axis=1)

# Calculate Daily Expected and Monthly Expected for each row
pivot["Daily Expected"] = pivot["BillTotal"] / day if day > 0 else 0
pivot["Monthly Expected"] = pivot["Daily Expected"] * num_days_in_month

# حساب الإجمالي الكلي
bill_total_sum = pivot["BillTotal"].sum()

# جدول ملخص إجمالي كل مندوب
mandoub_summary = pivot.groupby(["MandoubCode", "MandoubName"]).agg(
    Total_Mandoub=("BillTotal", "sum"),
    Daily_Expected_Mandoub=("Daily Expected", "sum"),
    Monthly_Expected_Mandoub=("Monthly Expected", "sum")
).reset_index()



# جدول ملخص إجمالي كل تاجر عند كل مندوب (بدون تفاصيل الخدمات)
user_summary = pivot.groupby(["MandoubCode", "MandoubName", "UserCode", "ShopName"])["BillTotal"].sum().reset_index()
user_summary = user_summary.rename(columns={"BillTotal": "Total_User_Mandoub"})

# جدول ملخص إجمالي كل خدمة عند كل مندوب
service_summary = df.groupby(["MandoubCode", "MandoubName", "CategoryName"])["BillTotal"].sum().reset_index()
service_summary = service_summary.rename(columns={"BillTotal": "Total_Service_Mandoub"})

user_service_summary = pd.pivot_table(df, values='BillTotal', index=["MandoubCode", "MandoubName", "UserCode", "ShopName"], columns=["CategoryName"], aggfunc='sum', fill_value=0).reset_index()
service_total = service_summary["Total_Service_Mandoub"].sum()
service_total_row = {col: "" for col in service_summary.columns}
service_total_row[service_summary.columns[2]] = "الإجمالي الكلي"
service_columns = [col for col in user_service_summary.columns if col not in ["MandoubCode", "MandoubName", "UserCode", "ShopName"]]
user_service_totals_by_category = user_service_summary[service_columns].sum()
user_service_total_row = {col: "" for col in user_service_summary.columns}
user_service_total_row["ShopName"] = "الإجمالي الكلي"
for col in service_columns:
    user_service_total_row[col] = user_service_totals_by_category[col]
user_service_summary = pd.concat([user_service_summary, pd.DataFrame([user_service_total_row])], ignore_index=True)

# إضافة صف الإجمالي في نهاية ملخص المندوبين في التقرير المجمع
mandoub_total = mandoub_summary["Total_Mandoub"].sum()
mandoub_daily_expected_total = mandoub_summary["Daily_Expected_Mandoub"].sum()
mandoub_monthly_expected_total = mandoub_summary["Monthly_Expected_Mandoub"].sum()

mandoub_total_row = {col: "" for col in mandoub_summary.columns}
mandoub_total_row[mandoub_summary.columns[1]] = "الإجمالي الكلي"
mandoub_total_row["Total_Mandoub"] = mandoub_total
mandoub_total_row["Daily_Expected_Mandoub"] = mandoub_daily_expected_total
mandoub_total_row["Monthly_Expected_Mandoub"] = mandoub_monthly_expected_total
mandoub_summary = pd.concat([mandoub_summary, pd.DataFrame([mandoub_total_row])], ignore_index=True)

# إضافة صف الإجمالي في نهاية ملخص التجار في التقرير المجمع
user_total = user_summary["Total_User_Mandoub"].sum()
user_total_row = {col: "" for col in user_summary.columns}
user_total_row[user_summary.columns[3]] = "الإجمالي الكلي"
user_total_row["Total_User_Mandoub"] = user_total
user_summary = pd.concat([user_summary, pd.DataFrame([user_total_row])], ignore_index=True)

# حفظ الناتج المجمع مع شيت الملخصات المطلوبة فقط
output_path = "Pivot_DSR_Report.xlsx"

# Calculate Grand Total, Daily Average, and Expected ACH for the consolidated report
grand_total_bill = pivot["BillTotal"].sum()
overall_daily_average = grand_total_bill / day if day > 0 else 0
overall_expected_ach = overall_daily_average * num_days_in_month

# Create a summary row for the consolidated report
summary_row_consolidated = pd.DataFrame({
    'MandoubCode': [''],
    'MandoubName': ['الإجمالي الكلي'],
    'UserCode': [''],
    'ShopName': [''],
    'BillTotal': [grand_total_bill],
    'Daily Expected': [overall_daily_average],
    'Monthly Expected': [overall_expected_ach]
})

# Concatenate the summary row to the pivot DataFrame
pivot_with_summary = pd.concat([pivot, summary_row_consolidated], ignore_index=True)

with ExcelWriter(output_path) as writer:
    pivot_with_summary.to_excel(writer, sheet_name="تفاصيل", index=False)
    mandoub_summary.to_excel(writer, sheet_name="ملخص المندوبين", index=False)
    user_summary.to_excel(writer, sheet_name="ملخص التجار", index=False)
    service_summary.to_excel(writer, sheet_name="ملخص الخدمات", index=False)
    user_service_summary.to_excel(writer, sheet_name="ملخص الخدمات والتجار", index=False)
print(f"تم إنشاء التقرير المجمع مع الملخصات: {output_path}")

# تقسيم كل مندوب في شيت منفصل مع ملخص التجار
mandoubs = pivot["MandoubCode"].unique()
for code in mandoubs:
    mandoub_df = pivot[pivot["MandoubCode"] == code]
    mandoub_name = mandoub_df["MandoubName"].iloc[0]
    file_name = f"Pivot_DSR_{mandoub_name}_{code}.xlsx"
    file_name = "".join(c if c.isalnum() or c in "_-. " else "_" for c in file_name)
    mandoub_user_summary = mandoub_df.groupby(["UserCode", "ShopName"])["BillTotal"].sum().reset_index()
    mandoub_user_summary = mandoub_user_summary.rename(columns={"BillTotal": "Total_User_Mandoub"})
    # إضافة صف الإجمالي في نهاية ملخص التجار لكل مندوب منفصل
    mandoub_user_total = mandoub_user_summary["Total_User_Mandoub"].sum()
    mandoub_user_total_row = {col: "" for col in mandoub_user_summary.columns}
    mandoub_user_total_row[mandoub_user_summary.columns[1]] = "الإجمالي الكلي"
    mandoub_user_total_row["Total_User_Mandoub"] = mandoub_user_total
    mandoub_user_summary = pd.concat([mandoub_user_summary, pd.DataFrame([mandoub_user_total_row])], ignore_index=True)

    # حساب Daily Average و Expected ACH للمندوب
    mandoub_df_total = mandoub_df["BillTotal"].sum()
    mandoub_daily_average = mandoub_df_total / day if day > 0 else 0
    mandoub_expected_ach = mandoub_daily_average * num_days_in_month

    # إنشاء صف ملخص للمندوب
    summary_row_mandoub = pd.DataFrame({
        'MandoubCode': [''],
        'MandoubName': ['الإجمالي الكلي'],
        'UserCode': [''],
        'ShopName': [''],
        'BillTotal': [mandoub_df_total],
        'Daily Expected': [mandoub_daily_average],
        'Monthly Expected': [mandoub_expected_ach]
    })

    # دمج صف الملخص مع DataFrame الخاص بالمندوب
    mandoub_df_with_summary = pd.concat([mandoub_df.drop(columns=service_columns_pivot), summary_row_mandoub], ignore_index=True)

    with ExcelWriter(file_name) as writer:
        mandoub_df_with_summary.to_excel(writer, sheet_name="تفاصيل", index=False)
        mandoub_user_summary.to_excel(writer, sheet_name="ملخص التجار", index=False)
    print(f"تم إنشاء تقرير للمندوب: {mandoub_name} ({code}) -> {file_name}")

print("تم الانتهاء من إنشاء جميع التقارير.")