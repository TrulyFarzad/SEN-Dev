# build_rating_from_support.py
# -*- coding: utf-8 -*-
import os
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string

# مسیرها
home = os.path.expanduser("~")
desktop = os.path.join(home, "Desktop")
takhsis_dir = os.path.join(desktop, "takhsis")
support_path = os.path.join(desktop, "فایل پشتیبانی.xlsx")  # فایل سنگین روی دسکتاپ
rating_out = os.path.join(takhsis_dir, "rating.xlsx")       # خروجی داخل پوشه takhsis

SHEET_NAME = "File"
AY_INDEX_0 = column_index_from_string("AY") - 1  # ایندکس صفر-پایه ستون AY

# تبدیل ارقام فارسی/عربی به لاتین
TRANS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")

WORDS_MAP = {
    "پله اول": 1, "پله یکم": 1, "پله یک": 1, "اول": 1, "یکم": 1, "یک": 1, "۱": 1, "1": 1,
    "پله دوم": 2, "دوم": 2, "دو": 2, "۲": 2, "2": 2,
    "پله سوم": 3, "سوم": 3, "سه": 3, "۳": 3, "3": 3,
    "پله چهارم": 4, "چهارم": 4, "چهار": 4, "۴": 4, "4": 4,
    "پله پنجم": 5, "پنجم": 5, "پنج": 5, "۵": 5, "5": 5,
    "پله ششم": 6, "ششم": 6, "شش": 6, "۶": 6, "6": 6,
}

def parse_rank(val):
    if pd.isna(val):
        return pd.NA
    s = str(val).strip()
    s = s.translate(TRANS)
    s = re.sub(r"\\s+", " ", s)
    m = re.search(r"\\b([1-6])\\b", s)
    if m:
        return int(m.group(1))
    for k, v in WORDS_MAP.items():
        if k in s:
            return v
    return pd.NA

print("📂 Reading support file:", support_path)
df = pd.read_excel(support_path, sheet_name=SHEET_NAME, dtype={"کد پذیرنده": str})

# چک طول ستون‌ها
if AY_INDEX_0 >= len(df.columns):
    raise IndexError("ستون AY در شیت File پیدا نشد یا تعداد ستون‌ها کمتر از AY است.")

income_col_name = df.columns[AY_INDEX_0]  # نام واقعی ستون «پله درآمد» دوره جاری (مثلاً پله درآمد تیر)

needed = ["کد پذیرنده", income_col_name]
for c in needed:
    if c not in df.columns:
        raise ValueError(f"ستون لازم پیدا نشد: {c}")

rating = df[needed].copy().rename(columns={income_col_name: "پله درآمد"})
rating["پله درآمد"] = rating["پله درآمد"].apply(parse_rank).astype("Int64")

# بهترین پله برای هر پذیرنده (بالاترین)
rating_clean = (
    rating.sort_values(["کد پذیرنده", "پله درآمد"], ascending=[True, False])
          .drop_duplicates(subset=["کد پذیرنده"], keep="first")
          .reset_index(drop=True)
)

# ذخیره خروجی
os.makedirs(takhsis_dir, exist_ok=True)
rating_clean.to_excel(rating_out, index=False)

# راست‌به‌چپ کردن
wb = load_workbook(rating_out)
ws = wb.active
try:
    ws.sheet_view.rightToLeft = True
except Exception:
    pass
wb.save(rating_out)

print("✅ rating.xlsx ساخته شد:", rating_out)
print("   ردیف‌ها (کل/یکتا):", len(rating), "/", len(rating_clean))
