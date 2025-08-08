# -*- coding: utf-8 -*-
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.worksheet.views import SheetView
from datetime import datetime
import jdatetime

# TODO: use better method for last-night duplicate deletion
# TODO: speed up search.xlsx loading time
# TODO: ستون پله درآمد

# تاریخ شمسی برای افزودن به نام فایل‌ها
today_jalali = jdatetime.date.today().strftime("%y%m%d")

# مسیر پوشه takhsis روی دسکتاپ هر کاربر
user_desktop = os.path.join(os.path.expanduser("~"), "Desktop", "takhsis")

in_wait_path = os.path.join(user_desktop, "in-wait.xlsx")
last_night_path = os.path.join(user_desktop, "last-night.xlsx")
search_path = os.path.join(user_desktop, "search.xlsx")
report_day_path = os.path.join(user_desktop, "takhsisReport.xlsx")
report_month_path = os.path.join(user_desktop, "takhsisReport-m.xlsx")

# --- بارگذاری فایل‌ها ---
print("📂 Reading file in-wait in:", in_wait_path)
in_wait = pd.read_excel(in_wait_path,  dtype={"کد پذیرنده": str, "سریال پوز تخصیص یافته": str})

print("📂 Reading file last-night in:", last_night_path)
last_night = pd.read_excel(last_night_path, dtype={"کد پذیرنده": str})

print("📂 Reading file search in:", search_path)
search = pd.read_excel(search_path, dtype={"کد پذیرنده": str, "سریال پایانه": str})

print("📂 Reading file takhsisReport in:", report_day_path)
report_day = pd.read_excel(report_day_path)

print("📂 Reading file takhsisReport-m in:", report_month_path)
report_month = pd.read_excel(report_month_path)

# حذف ستون‌های اضافی از in-wait
columns_to_keep = [
    "کد پذیرنده", "کد درخواست", "کد پیگیری", "مدل پوز", "گروه پروژه",
    "تاریخ ایجاد", "آخرین تاریخ ویرایش", "شماره حساب"
]
in_wait = in_wait[[col for col in columns_to_keep if col in in_wait.columns]]

# Helper: POS type based on model
def get_pos_type(model):
    if pd.isna(model): return ""
    if model.strip().upper() == "GPRS":
        return "بیسیم"
    elif model.strip().upper() in ["LAN", "DIALUP", "PCPOSLAN"]:
        return "ثابت"
    return "نامشخص"

# Merge step 1: in-wait with search (on کد پذیرنده)
merged = pd.merge(
    in_wait,
    search[["کد پذیرنده", "نام فروشگاه", "شهر", "آدرس"]],
    on="کد پذیرنده",
    how="left"
)

# Merge step 2: with last-night for پشتیبان و توضیح (تنها یک ردیف به‌ازای هر کد پذیرنده)
last_night_info = last_night[["کد پذیرنده", "نام پشتیبان", "نام خانوادگی پشتیبان", "توضیح"]].copy()
last_night_info["نام و نام خانوادگی پشتیبان"] = last_night_info["نام پشتیبان"].fillna("") + " " + last_night_info["نام خانوادگی پشتیبان"].fillna("")
last_night_info = last_night_info.drop(columns=["نام پشتیبان", "نام خانوادگی پشتیبان"])
last_night_info = last_night_info.rename(columns={"توضیح": "توضیحات"})
last_night_info = last_night_info.drop_duplicates(subset="کد پذیرنده")

merged = pd.merge(
    merged,
    last_night_info,
    on="کد پذیرنده",
    how="left"
)

# توضیحات نهایی با در نظر گرفتن شرایط خاص
merged["توضیحات"] = merged["توضیحات"].fillna("نصب اولیه")
merged.loc[merged["توضیحات"] == "--", "توضیحات"] = ""
merged["توضیحات"] = merged["توضیحات"].apply(lambda x: x.split(" - ")[0].strip() if isinstance(x, str) else x)

# گروه پایانه بر اساس مدل پوز
merged["گروه پایانه"] = merged["مدل پوز"].apply(get_pos_type)

# انتخاب و ساخت جدول نهایی
result_cols = [
    "کد پذیرنده", "نام فروشگاه", "توضیحات", "شهر", "آدرس", "نام و نام خانوادگی پشتیبان",
    "گروه پایانه", "کد درخواست", "کد پیگیری", "مدل پوز", "گروه پروژه",
    "تاریخ ایجاد", "آخرین تاریخ ویرایش", "شماره حساب"
]

final_result = merged[[col for col in result_cols if col in merged.columns]].copy()

# ساخت فایل نصب اولیه و فیلتر پروژه فروش
initial_installs = final_result[final_result["توضیحات"] == "نصب اولیه"].copy()
filtered_result = final_result[final_result["گروه پروژه"] != "پروژه فروش"].copy()

# ---- ساخت گزارش تخصیص ----
report_sections = []

# بخش 1: در انتظار تخصیص
waiting_total = len(filtered_result)
waiting_fixed = (filtered_result["گروه پایانه"] == "ثابت").sum()
waiting_wireless = (filtered_result["گروه پایانه"] == "بیسیم").sum()
report_sections.append(pd.DataFrame({
    "نوع": ["کل درخواست‌ها", "درخواست‌های ثابت", "درخواست‌های سیار"],
    "تعداد": [waiting_total, waiting_fixed, waiting_wireless]
}))

# Helper برای گزارش پروژه‌ها
def project_counts(df):
    persian_switch = df[df["پروژه"].str.contains("پرشين", na=False)].shape[0]
    sales_project = df[df["پروژه"].str.contains("فروش", na=False)].shape[0]
    bank_project = len(df) - persian_switch - sales_project
    return pd.DataFrame({
        "نوع": ["پروژه پرشین سوییچ", "پروژه فروش", "پروژه بانکی", "مجموع"],
        "تعداد": [persian_switch, sales_project, bank_project, persian_switch + sales_project + bank_project]
    })

# بخش 2: تخصیص‌های روز قبل
report_sections.append(project_counts(report_day))

# بخش 3: تخصیص‌های از اول ماه
report_sections.append(project_counts(report_month))

# ترکیب بخش‌ها
allocation_report = pd.concat(report_sections, keys=["در انتظار تخصیص", "تخصیص روز قبل", "تخصیص از اول ماه"], names=["بخش", "ردیف"])

# ذخیره فایل تخصیص با تنظیم راست‌به‌چپ
output_path = os.path.join(user_desktop, f"takhsis{today_jalali}.xlsx")
filtered_result.to_excel(output_path, index=False, sheet_name="نتیجه")
wb = load_workbook(output_path)
ws = wb["نتیجه"]
ws.sheet_view.rightToLeft = True
wb.save(output_path)

# ذخیره فایل نصب اولیه با تنظیم راست‌به‌چپ
initial_path = os.path.join(user_desktop, f"نصب اولیه{today_jalali}.xlsx")
initial_installs.to_excel(initial_path, index=False, sheet_name="نصب اولیه")
wb2 = load_workbook(initial_path)
ws2 = wb2["نصب اولیه"]
ws2.sheet_view.rightToLeft = True
wb2.save(initial_path)

# ذخیره گزارش تخصیص با تنظیم راست‌به‌چپ
allocation_path = os.path.join(user_desktop, f"گزارش تخصیص{today_jalali}.xlsx")
allocation_report.to_excel(allocation_path, sheet_name="گزارش")
wb3 = load_workbook(allocation_path)
ws3 = wb3["گزارش"]
ws3.sheet_view.rightToLeft = True
wb3.save(allocation_path)

print("\n✅ فایل‌ها ذخیره شدند!\n📁", output_path, "\n📁", initial_path, "\n📁", allocation_path)
