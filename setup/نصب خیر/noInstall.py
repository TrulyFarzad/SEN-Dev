# -*- coding: utf-8 -*-
"""
noInstall.py
==============================

هدف:
-----
این اسکریپت برای مدیریت و اتوماسیون «پیگیری نصب‌خیر» طراحی شده است. «نصب‌خیر»
به لیست تجهیزاتی گفته می‌شود که تخصیص خورده‌اند اما هنوز نصب نشده‌اند و باید
روزانه با پشتیبان‌ها پیگیری شوند. خروجی اسکریپت یک فایل اکسل است با چند شیت
که وضعیت «در انتظار نصب»، «نصب‌شده‌های تازه کشف‌شده» و «آرشیو نصب‌شده‌ها» را
نگهداری می‌کند؛ همچنین مواردی که بعد از تخصیص «غیرفعال» شده‌اند را لاگ می‌کند.

ورودی‌ها (در مسیر Desktop/noInstall/input):
--------------------------------------------
- install.xlsx  : گزارش شب گذشته (یا کلی) از تجهیزات فعال در بستر POS.
                  «وضعیت نصب» یا «تاریخ نصب» در این فایل مشخص است.
                  (برای نصب‌خیر، ردیف‌هایی که «وضعیت نصب = خیر» دارند ملاک‌اند.)
- 1025.xlsx     : لیست تراکنش‌های نوع 1025 (تست بعد از تخصیص).
- خروج.xlsx     : ثبت خروج تجهیز از شرکت (تحویل به پشتیبان / ارسال پست / ...).
                  ستون «توضیحات» اگر حاوی «نزد پشتیبان» باشد یعنی تخصیص از نزد پشتیبان.
- disable.xlsx  : دستگاه‌هایی که در یک ماه اخیر «پایان تخصیص/غیرفعال» داشته‌اند
                  (ستون کلیدی: «تاریخ پایان تخصیص»).

خروجی‌ها (در مسیر Desktop/noInstall):
--------------------------------------
- install_kheir_output.xlsx با شیت‌های:
  1) Pending                : در انتظار نصب‌های فعلی (از install.xlsx با «وضعیت نصب = خیر»)
      - ستون‌های اصلی شامل «تاریخ تخصیص تجهیز»، «تاریخ تراکنش 1025»، «خروج» (ممکن است «- نزد پشتیبان» داشته باشد)،
        پرچم «از_نزد_پشتیبان»، ستون «توضیح» (یادداشت‌های پیگیری که بین اجراها حفظ می‌شود)، ...
  2) Installed_Candidates   : مواردی که از Pending قدیمی حذف شده‌اند (احتمالاً نصب‌شده)،
      برایشان «تاریخ نصب» از install جدید جستجو می‌شود؛ «تاخیر روز» با منطق SLA محاسبه می‌شود،
      «Fraud detection» چک می‌شود. موارد هشداردار در همین شیت باقی می‌مانند و آرشیو نمی‌شوند.
  3) Archive                : آرشیوِ همه نصب‌شده‌هایی که هشدار تقلب ندارند (کپی از شیت 2 در همان اجرا).
  4) Disabled_Log           : مواردی که بعد از تخصیص، قبل از نصب، در فایل disable آمده و حذف شده‌اند.

منطق کلیدی:
-----------
1) فیلتر «نصب‌نشده‌ها»:
   - از install.xlsx فقط ردیف‌هایی که «وضعیت نصب = خیر» دارند، به‌عنوان Pending گرفته می‌شود.
   - پروژه «پروژه فروش» حذف می‌شود (در این پروژه‌ها اساساً در نصب‌خیر پیگیری نمی‌کنیم).

2) استانداردسازی تاریخ‌ها (همه به سطح روز جلالی):
   - تاریخ‌ها به صورت کلید عددی YYYYMMDD استخراج می‌شوند؛ برای نمایش «YYYY/MM/DD».

3) انتخاب تاریخ‌ها:
   - تخصیص: از ستون «تاریخ تخصیص تجهیز» در install.
   - 1025 : اولین تاریخ 1025 که «روز ≥ تخصیص» باشد (اگر یافت نشد، خالی).
   - خروج: فقط با تخصیص مقایسه می‌شود (نه با 1025). اگر «نزد پشتیبان» موجود باشد و «روز ≥ تخصیص»، همان اولویت دارد.
           در غیر اینصورت اولین «خروج» بعد از «تخصیص» انتخاب می‌شود. اگر توضیح «نزد پشتیبان» داشت، در خروج
           به صورت «YYYY/MM/DD - نزد پشتیبان» ثبت می‌شود.

4) از_نزد_پشتیبان:
   - اگر خروج «نزد پشتیبان» پس از تخصیص وجود داشت → پرچم True (وگرنه False).
   - این پرچم تعیین می‌کند «پایه_تاخیر» چه باشد:
       - True  → base = تاریخ خروج
       - False → base = تاریخ 1025

5) محاسبهٔ تاخیر و SLA:
   - SLA شهر: مشهد = ۲ روز، سایر شهرها = ۵ روز.
   - تاخیر = max(0, (تاریخ نصب - base) - SLA).
   - اگر base موجود نبود، یا Fraud هشدار داد، «تاخیر روز» NA می‌شود.

6) Fraud detection:
   - اگر «تاریخ 1025 > تاریخ خروج» باشد (به سطح روز)، پرچم هشدار «هشدار_احتمال_تقلب=True».
   - موارد هشداردار به آرشیو نمی‌روند و در Installed_Candidates می‌مانند تا بررسی دستی شوند.

7) نگهداری یادداشت‌های پیگیری («توضیح»):
   - Pending جدید با Pending قبلی بر اساس «سریال پایانه» left-merge می‌شود تا اگر «توضیح» جدید خالی بود،
     از مقدار قدیمی پر شود. این باعث می‌شود یادداشت‌های پیگیری قبلی از بین نروند.
   - وقتی ردیفی از Pending قدیمی به Installed_Candidates منتقل می‌شود، «توضیح» را همراه خود می‌برد؛
     و در صورت آرشیو، همان «توضیح» نیز همراهش به Archive می‌رود.

8) حذف خودکار موارد disable:
   - اگر ردیفی در Pending یا Installed_Candidates (بدون تاریخ نصب) باشد و برای همان سریال (و ترجیحاً همان کد پذیرنده)
     در disable.xlsx «تاریخ پایان تخصیص ≥ تاریخ تخصیص» یافت شود، آن ردیف از چرخه حذف و در Disabled_Log ثبت می‌شود.

9) پاکسازی شیت 2 قبل از هر اجرا:
   - ابتدای هر اجرا ردیف‌های دارای «تاریخ نصب» و «هشدار=False» از Installed_Candidates حذف می‌شوند.
     (موارد هشداردار حتی اگر تاریخ نصب داشته باشند باقی می‌مانند تا دستی تصمیم‌گیری شود.)

10) استایل‌ها و جهت صفحه:
    - شیت‌ها Right-to-Left.
    - در Installed_Candidates: سطرهای «هشدار=True» قرمز کم‌رنگ؛
      سلول‌های «تاخیر روز > 0» نارنجی ملایم.

نحوه اجرا:
----------
- پیش‌نیاز: نصب xlsxwriter →  pip install xlsxwriter
- پوشه‌ها: Desktop/noInstall/input باید شامل چهار فایل ورودی باشد.
- اجرای مستقیم: python noInstall.py
- خروجی: Desktop/noInstall/install_kheir_output.xlsx

محدودیت‌ها و نکات:
-------------------
- مقایسهٔ تاریخ‌ها در سطح روز جلالی انجام می‌شود؛ ساعت/دقیقه لحاظ نمی‌شود.
- کلید همسان‌سازی در اکثر جاها «سریال پایانه» است؛ در بعضی کنترل‌ها «کد پذیرنده» نیز لحاظ می‌شود.
- اگر نام ستون‌ها در ورودی‌ها متفاوت باشد، ممکن است نیاز به تنظیم کوچک باشد.
- این نسخه برای اجرای تک‌فایلی بهینه شده است؛ برای مهاجرت به SQL، می‌توان I/O را جدا کرد.

"""

import sys, os, shutil, re
from datetime import date as _date, date
from pathlib import Path
import pandas as pd

# تلاش برای وارد کردن xlsxwriter (برای نوشتن اکسل با استایل)
try:
    import xlsxwriter
except Exception:
    print("❌ xlsxwriter نصب نیست. اجرا: pip install xlsxwriter")
    sys.exit(1)

# -------------------- مسیرها و مقدمات --------------------
def get_desktop():
    """
    تلاش امن برای یافتن مسیر دسکتاپ کاربر در ویندوز/لینوکس/مک.
    اگر Desktop پیدا نشود، از home استفاده می‌کند.
    """
    home = Path.home()
    for p in [Path(os.environ.get("USERPROFILE",""))/"Desktop", home/"Desktop", home]:
        if p.exists(): return p
    return home

DESKTOP   = get_desktop()
BASE_DIR  = DESKTOP / "noInstall"
INPUT_DIR = BASE_DIR / "input"
OUTPUT    = BASE_DIR / "install_kheir_output.xlsx"
BASE_DIR.mkdir(parents=True, exist_ok=True)
INPUT_DIR.mkdir(parents=True, exist_ok=True)

# -------------------- توابع کمکی عمومی --------------------
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    یکسان‌سازی نام ستون‌ها: جایگزینی حروف عربی با فارسی، حذف فاصله‌های اضافه.
    """
    df = df.copy()
    df.columns = df.columns.astype(str).str.replace("ي","ی").str.replace("ك","ک").str.strip()
    return df

def normalize_text(v) -> str:
    """
    نرمال‌سازی متن سلولی: یکسان‌سازی ی/ک، حذف نیم‌فاصله، فشرده‌سازی فاصله.
    برای مقایسه‌های متنی قابل اعتمادتر.
    """
    if pd.isna(v): return ""
    s = str(v).replace("ي","ی").replace("ك","ک").replace("\u200c","")
    return re.sub(r"\s+"," ", s).strip()

def extract_day_key(v) -> int|None:
    """
    استخراج کلید روز جلالی به صورت عددی YYYYMMDD از یک رشتهٔ تاریخ/تاریخ-زمان.
    اگر کمتر از 8 رقم یافت شود، None برمی‌گرداند.
    """
    if pd.isna(v): return None
    digits = "".join(ch for ch in str(v) if ch.isdigit())
    if len(digits) < 8: return None
    return int(digits[:8])

def pretty_jalali(v) -> str|None:
    """
    تولید نمایش استاندارد YYYY/MM/DD از مقدار ورودی (با استفاده از extract_day_key).
    """
    k = extract_day_key(v)
    if k is None: return None
    y,m,d = k//10000, (k//100)%100, k%100
    return f"{y:04d}/{m:02d}/{d:02d}"

# --- تبدیل جلالی→میلادی برای اختلاف روز (محاسبه تاخیر) ---
def jalali_to_gregorian(jy, jm, jd):
    """
    تبدیل تاریخ جلالی به میلادی (محاسبات روزمحور) برای محاسبه اختلاف روزها.
    """
    jy += 1595
    days = -355668 + 365*jy + (jy//33)*8 + ((jy%33)+3)//4 + jd
    days += (jm-1)*31 if jm<7 else ((jm-7)*30 + 186)
    gy = 400*(days//146097); days%=146097
    if days>36524:
        gy += 100*((days-1)//36524); days=(days-1)%36524
        if days>=365: days+=1
    gy += 4*(days//1461); days%=1461
    if days>365:
        gy += (days-1)//365; days=(days-1)%365
    gd = days+1
    leap = (days==0)
    gmd = [0,31,29 if leap else 28,31,30,31,30,31,31,30,31,30,31]
    gm=1
    while gm<=12 and gd>gmd[gm]:
        gd-=gmd[gm]; gm+=1
    return gy,gm,gd

def jalali_key_to_ordinal(key:int) -> int|None:
    """
    تبدیل کلید YYYYMMDD جلالی به ordinal میلادی برای محاسبه اختلاف روزها.
    """
    y=key//10000; m=(key//100)%100; d=key%100
    try:
        gy,gm,gd = jalali_to_gregorian(y,m,d)
        from datetime import date as _d
        return _d(gy,gm,gd).toordinal()
    except: return None

def days_diff_jalali(start_key:int|None, end_key:int|None) -> int|None:
    """
    اختلاف روز بین دو کلید جلالی (end - start).
    اگر هر کدام None باشد، None برمی‌گرداند.
    """
    if start_key is None or end_key is None: return None
    s = jalali_key_to_ordinal(start_key); e = jalali_key_to_ordinal(end_key)
    if s is None or e is None: return None
    return e - s

def sla_days(city:str) -> int:
    """
    SLA شهر: مشهد=۲ روز، سایر شهرها=۵ روز.
    """
    return 2 if normalize_text(city) == "مشهد" else 5

def backup_prev(path: Path) -> Path|None:
    """
    اگر خروجی قبلی وجود دارد، یک بک‌آپ با پسوند تاریخ روز می‌سازد.
    """
    if not path.exists(): return None
    b = path.with_name(path.stem + _date.today().strftime("_prev_%Y%m%d") + path.suffix)
    shutil.copy2(path, b); return b

def read_prev_triplet(prev_path: Path):
    """
    خواندن سه شیت خروجی قبلی (اگر باشد). اگر نبود، دیتافریم‌های خالی برمی‌گرداند.
    Pending (ساده‌تر)، Sheet2 و Archive (ستون‌های افزوده) را هم‌تراز می‌کند.
    """
    cols1 = [
        "کد پذیرنده","نام فروشگاه","شهر","آدرس","مدل پایانه","کد پایانه","سریال پایانه",
        "نام خانوادگی پشتیبان","پروژه",
        "تاریخ تخصیص تجهیز","تاریخ تراکنش 1025","خروج","از_نزد_پشتیبان",
        "توضیح","مهلت","تاریخ نصب"
    ]
    ext  = cols1 + ["پایه_تاخیر","تحویل پست","تاخیر روز","هشدار_احتمال_تقلب"]
    if not prev_path or not prev_path.exists():
        return pd.DataFrame(columns=cols1), pd.DataFrame(columns=ext), pd.DataFrame(columns=ext)
    xls = pd.ExcelFile(prev_path)
    def safe(idx, cols):
        try:
            df = normalize_columns(xls.parse(idx))
            for c in cols:
                if c not in df.columns: df[c]=pd.NA
            return df[cols]
        except: return pd.DataFrame(columns=cols)
    return safe(0,cols1), safe(1,ext), safe(2,ext)

def load_inputs():
    """
    بارگذاری چهار ورودی اصلی از noInstall/input:
      - install.xlsx, 1025.xlsx, خروج.xlsx, disable.xlsx
    اگر هر کدام نبود، خطا می‌دهد.
    """
    f_install = INPUT_DIR/"install.xlsx"
    f_1025    = INPUT_DIR/"1025.xlsx"
    f_exit    = INPUT_DIR/"خروج.xlsx"
    f_disable = INPUT_DIR/"disable.xlsx"
    missing   = [p.name for p in (f_install,f_1025,f_exit,f_disable) if not p.exists()]
    if missing:
        raise FileNotFoundError("فایل‌های ورودی در noInstall/input نیستند: " + ", ".join(missing))
    return (normalize_columns(pd.read_excel(f_install)),
            normalize_columns(pd.read_excel(f_1025)),
            normalize_columns(pd.read_excel(f_exit)),
            normalize_columns(pd.read_excel(f_disable)))

# -------------------- ایندکس‌سازها برای جستجوی سریع تاریخ‌ها --------------------
def build_1025_index(df_1025, serial_col, date_col):
    """
    اندیس‌سازی فایل 1025: برای هر سریال لیستی از (day_key, pretty_date) با ترتیب نزولی تاریخ.
    """
    tmp = df_1025[[serial_col, date_col]].copy()
    tmp["_day"]    = tmp[date_col].apply(extract_day_key)
    tmp["_pretty"] = tmp[date_col].apply(pretty_jalali)
    tmp = tmp.dropna(subset=["_day"]).sort_values("_day", ascending=False)
    d={}
    for s,grp in tmp.groupby(serial_col):
        d[str(s)] = list(zip(grp["_day"].tolist(), grp["_pretty"].tolist()))
    return d

def build_exit_index_with_flag(df_exit, serial_col, date_col):
    """
    اندیس‌سازی خروج: برای هر سریال لیستی از (day_key, pretty_date, is_nazdPoshtiban)
    - اگر در «توضیحات» عبارت «نزد پشتیبان» وجود داشته باشد، is_nazd=True است.
    - در صورت is_nazd، pretty_date با « - نزد پشتیبان» تزئین می‌شود.
    """
    note_col = "توضیحات" if "توضیحات" in df_exit.columns else None
    cols = [serial_col, date_col] + ([note_col] if note_col else [])
    tmp = df_exit[cols].copy()
    tmp["_day"]    = tmp[date_col].apply(extract_day_key)
    tmp["_pretty"] = tmp[date_col].apply(pretty_jalali)

    def make_tuple(row):
        day = row["_day"]
        if day is None: return None
        pretty = row["_pretty"]
        is_nazd = False
        if note_col:
            is_nazd = "نزد پشتیبان" in normalize_text(row[note_col])
            if pretty is not None and is_nazd:
                pretty = pretty + " - نزد پشتیبان"
        return (day, pretty, is_nazd)

    tmp["_t"] = tmp.apply(make_tuple, axis=1)
    tmp = tmp.dropna(subset=["_day"]).sort_values("_day", ascending=False)

    d={}
    for s,grp in tmp.groupby(serial_col):
        d[str(s)] = [t for t in grp["_t"].tolist() if t is not None]
    return d

def build_disable_index(df_disable, serial_col):
    """
    اندیس disable بر اساس ستون «تاریخ پایان تخصیص» (اگر نبود: fallback به ستونی که «پایان تخصیص» در نام دارد،
    یا در نهایت اولین ستونی که «تاریخ» دارد).
    خروجی: dict[serial] = [(day_key, pretty_str, merchant_code_str), ...]  (جدیدترین در اول)
    """
    date_col = "تاریخ پایان تخصیص"
    if date_col not in df_disable.columns:
        cand = [c for c in df_disable.columns if "پایان تخصیص" in c]
        if cand:
            date_col = cand[0]
        else:
            cand = [c for c in df_disable.columns if "تاریخ" in c]
            if cand:
                date_col = cand[0]
            else:
                return {}

    merch_col = "کد پذیرنده" if "کد پذیرنده" in df_disable.columns else None
    cols = [serial_col, date_col] + ([merch_col] if merch_col else [])
    tmp = df_disable[cols].copy()
    tmp["_day"]    = tmp[date_col].apply(extract_day_key)
    tmp["_pretty"] = tmp[date_col].apply(pretty_jalali)
    if merch_col:
        tmp["_merch"] = tmp[merch_col].astype(str)
    else:
        tmp["_merch"] = ""
    tmp = tmp.dropna(subset=["_day"]).sort_values("_day", ascending=False)

    d={}
    for s, grp in tmp.groupby(serial_col):
        d[str(s)] = list(zip(grp["_day"].tolist(), grp["_pretty"].tolist(), grp["_merch"].tolist()))
    return d

# -------------------- انتخاب تاریخ‌ها با قواعد تعریف‌شده --------------------
def pick_exit_after_alloc(exit_idx:dict, serial:str, alloc_day:int|None):
    """
    انتخاب خروج پس از تخصیص (فقط با تخصیص مقایسه می‌شود):
      1) اگر خروج «نزد پشتیبان» با day >= تخصیص وجود داشت → همان را برگردان
      2) در غیر اینصورت، اولین خروج با day >= تخصیص
    خروجی: (exit_day_key, exit_pretty, is_nazdPoshtiban)
    """
    if alloc_day is None: return None, None, False
    items = exit_idx.get(str(serial))
    if not items: return None, None, False
    # اولویت با نزد پشتیبان
    for day, pretty, is_nazd in items:
        if day >= alloc_day and is_nazd:
            return day, pretty, True
    for day, pretty, is_nazd in items:
        if day >= alloc_day:
            return day, pretty, False
    return None, None, False

def pick_1025_after_alloc(idx_1025:dict, serial:str, alloc_day:int|None):
    """
    انتخاب اولین 1025 پس از تخصیص (day >= تخصیص). اگر نیافت، None.
    خروجی: (test_day_key, test_pretty)
    """
    if alloc_day is None: return None, None
    items = idx_1025.get(str(serial))
    if not items: return None, None
    for day, pretty in items:
        if day >= alloc_day:
            return day, pretty
    return None, None

# -------------------- ابزار کمکی خروجی اکسل --------------------
def col_letter(idx_zero_based:int) -> str:
    """
    تبدیل شماره ستون صفر-مبنا به حروف اکسل (A, B, ..., AA, AB, ...)
    """
    s = ""
    n = idx_zero_based + 1
    while n:
        n, rem = divmod(n-1, 26)
        s = chr(65+rem) + s
    return s

def coalesce_text(a, b):
    """
    انتخاب مقدار متن غیرخالی: اگر a خالی بود، b؛ در غیر اینصورت a.
    برای حفظ «توضیح» قبلی وقتی جدید خالی است.
    """
    a_ = normalize_text(a)
    b_ = normalize_text(b)
    return a if a_ != "" else (b if b_ != "" else a)

# -------------------- اجرای اصلی Pipeline --------------------
def main():
    # 1) ورودی‌ها: چهار فایل
    df_install_full, df_1025, df_exit, df_disable = load_inputs()

    # ستون‌های کلیدی
    serial_col = "سریال پایانه"
    alloc_col  = "تاریخ تخصیص تجهیز"
    proj_col   = "پروژه"
    status_col = "وضعیت نصب"

    # صحت وجود ستون‌ها در install
    for col in [serial_col, alloc_col, proj_col, status_col]:
        if col not in df_install_full.columns:
            raise KeyError(f"ستون «{col}» در install.xlsx نیست.")

    # 2) حذف پروژه فروش از install کامل
    df_install_full = df_install_full[df_install_full[proj_col].apply(lambda x: normalize_text(x)!="پروژه فروش")].copy()

    # 3) Pending = نصب‌نشده‌ها (وضعیت نصب = خیر)
    df_install = df_install_full[df_install_full[status_col].apply(lambda x: normalize_text(x)=="خیر")].copy()

    # 4) استانداردسازی و استخراج تاریخ تخصیص
    df_install["__alloc_day"]    = df_install[alloc_col].apply(extract_day_key)
    df_install["__alloc_pretty"] = df_install[alloc_col].apply(pretty_jalali)

    # 5) ساخت ایندکس‌ها برای جستجوی سریع
    #    - ستون تاریخ در 1025/خروج را با اولین ستونی که «تاریخ» در نام دارد می‌یابیم
    date_col_1025 = next(c for c in df_1025.columns if "تاریخ" in c)
    if serial_col not in df_exit.columns and "سریال" in df_exit.columns:
        df_exit.rename(columns={"سریال": serial_col}, inplace=True)
    exit_date_col = next(c for c in df_exit.columns if "تاریخ" in c)

    idx_1025    = build_1025_index(df_1025, serial_col, date_col_1025)
    idx_exit    = build_exit_index_with_flag(df_exit, serial_col, exit_date_col)
    idx_disable = build_disable_index(df_disable, serial_col)

    # 6) ساخت Pending جدید با پر کردن تاریخ‌های نمایش و پرچم نزد پشتیبان
    rows=[]
    for _, r in df_install.iterrows():
        serial    = str(r.get(serial_col,""))
        alloc_day = r["__alloc_day"]
        alloc_pre = r["__alloc_pretty"]

        t1025_day, t1025_pre = pick_1025_after_alloc(idx_1025, serial, alloc_day)
        exit_day, exit_pre, is_nazd = pick_exit_after_alloc(idx_exit, serial, alloc_day)

        out = dict(r)
        out["تاریخ تخصیص تجهیز"] = alloc_pre
        out["تاریخ تراکنش 1025"] = t1025_pre
        out["خروج"]              = exit_pre
        out["از_نزد_پشتیبان"]   = bool(is_nazd)
        rows.append(out)

    df_pending = pd.DataFrame(rows)
    df_pending = normalize_columns(df_pending)

    # ستون‌های نهایی Pending (سازگار با خروجی قدیم)
    s1_cols = ["کد پذیرنده","نام فروشگاه","شهر","آدرس","مدل پایانه","کد پایانه","سریال پایانه",
               "نام خانوادگی پشتیبان","پروژه",
               "تاریخ تخصیص تجهیز","تاریخ تراکنش 1025","خروج","از_نزد_پشتیبان",
               "توضیح","مهلت","تاریخ نصب"]
    for c in s1_cols:
        if c not in df_pending.columns: df_pending[c]=pd.NA
    df_pending = df_pending[s1_cols]

    # 7) خروجی قبلی را بخوان و از آن برای حفظ «توضیح» استفاده کن
    prev_backup = backup_prev(OUTPUT)
    prev_pending, prev_sheet2, prev_archive = read_prev_triplet(prev_backup if prev_backup else OUTPUT)

    # نگهداری توضیحات قبلی: merge روی «سریال پایانه»، و coalesce روی ستون «توضیح»
    if not prev_pending.empty and not df_pending.empty:
        df_pending = df_pending.merge(
            prev_pending[["سریال پایانه","توضیح"]],
            on="سریال پایانه", how="left", suffixes=("", "_old")
        )
        df_pending["توضیح"] = df_pending.apply(
            lambda r: coalesce_text(r.get("توضیح"), r.get("توضیح_old")), axis=1
        )
        if "توضیح_old" in df_pending.columns:
            df_pending.drop(columns=["توضیح_old"], inplace=True)

    # 8) حذف از Pending بر اساس disable (غیرفعال‌شده پس از تخصیص)
    disabled_log_rows = []
    if not df_pending.empty:
        keep_mask = []
        for _, row in df_pending.iterrows():
            serial = str(row["سریال پایانه"]).strip()
            merch  = str(row.get("کد پذیرنده","")).strip()
            alloc_day = extract_day_key(row.get("تاریخ تخصیص تجهیز"))
            dis_items = idx_disable.get(serial, [])
            picked = None
            # جدیدترین disable پس از تخصیص، با ترجیح match کد پذیرنده
            for dday, dpretty, dmerch in dis_items:
                if alloc_day is not None and dday >= alloc_day and (merch=="" or dmerch==merch):
                    picked = (dday, dpretty); break
            if picked is None:
                keep_mask.append(True)
            else:
                # حذف از Pending و ثبت در Disabled_Log
                log = dict(row)
                log["تاریخ غیر فعال"] = picked[1]  # نمایش استاندارد از «تاریخ پایان تخصیص»
                disabled_log_rows.append(log)
                keep_mask.append(False)
        df_pending = df_pending[keep_mask].copy()

    # 9) ابتدای هر اجرا: پاکسازی شیت2 قبلی از موارد نصب‌شده بدون هشدار
    sheet2 = prev_sheet2.copy()
    if not sheet2.empty:
        warn_col = "هشدار_احتمال_تقلب" if "هشدار_احتمال_تقلب" in sheet2.columns else None
        if warn_col:
            mask_keep = sheet2["تاریخ نصب"].isna() | (sheet2[warn_col]==True)
        else:
            mask_keep = sheet2["تاریخ نصب"].isna()
        sheet2 = sheet2[mask_keep].copy()

    for c in ["پایه_تاخیر","تحویل پست","تاخیر روز","هشدار_احتمال_تقلب"]:
        if c not in sheet2.columns: sheet2[c]=pd.NA

    # 10) انتقال «نصب‌شده‌های جدید» از prev_pending به sheet2:
    #     newly_installed_serials = prev_pending - curr_pending (بر اساس سریال)
    prev_serials = set(prev_pending["سریال پایانه"].astype(str).fillna("")) if not prev_pending.empty else set()
    curr_serials = set(df_pending["سریال پایانه"].astype(str).fillna(""))   if not df_pending.empty else set()
    newly_installed_serials = prev_serials - curr_serials
    if newly_installed_serials:
        new_cands = prev_pending[prev_pending["سریال پایانه"].astype(str).isin(newly_installed_serials)].copy()
        sheet2 = pd.concat([sheet2, new_cands], ignore_index=True)

    # 11) حذف از Sheet2 بر اساس disable (برای مواردی که هنوز تاریخ نصب ندارند)
    if not sheet2.empty:
        keep_mask2 = []
        for _, row in sheet2.iterrows():
            if pd.notna(row.get("تاریخ نصب")):
                keep_mask2.append(True)
                continue
            serial = str(row["سریال پایانه"]).strip()
            merch  = str(row.get("کد پذیرنده","")).strip()
            alloc_day = extract_day_key(row.get("تاریخ تخصیص تجهیز"))
            dis_items = idx_disable.get(serial, [])
            picked = None
            for dday, dpretty, dmerch in dis_items:
                if alloc_day is not None and dday >= alloc_day and (merch=="" or dmerch==merch):
                    picked = (dday, dpretty); break
            if picked is None:
                keep_mask2.append(True)
            else:
                # حذف از Sheet2 و ثبت در Disabled_Log
                log = dict(row)
                log["تاریخ غیر فعال"] = picked[1]
                disabled_log_rows.append(log)
                keep_mask2.append(False)
        sheet2 = sheet2[keep_mask2].copy()

    # 12) تکمیل «تاریخ نصب» و محاسبهٔ «تاخیر» + «پایه_تاخیر» + Fraud روی Sheet2
    df_lu = df_install_full.copy()
    if "تاریخ نصب" not in df_lu.columns:
        df_lu["تاریخ نصب"] = pd.NA
    df_lu["__install_day"]    = df_lu["تاریخ نصب"].apply(extract_day_key)
    df_lu["__install_pretty"] = df_lu["تاریخ نصب"].apply(pretty_jalali)

    install_days = []
    delays = []
    bases  = []
    frauds = []

    for _, row in sheet2.iterrows():
        serial = str(row.get("سریال پایانه","")).strip()
        merch  = str(row.get("کد پذیرنده","")).strip()
        alloc_day = extract_day_key(row.get("تاریخ تخصیص تجهیز"))
        test_day  = extract_day_key(row.get("تاریخ تراکنش 1025"))
        exit_day  = extract_day_key(row.get("خروج"))
        # پرچم نزد پشتیبان: تعیین «پایه_تاخیر»
        is_nazd   = str(row.get("از_نزد_پشتیبان","")).strip().lower() in ("true","1","بله","yes")

        # از install کامل: جدیدترین تاریخ نصب معتبر (≥ تخصیص) برای همین سریال+کد پذیرنده
        sub = df_lu[
            (df_lu["سریال پایانه"].astype(str).str.strip()==serial) &
            (df_lu["کد پذیرنده"].astype(str).str.strip()==merch) &
            (df_lu["__install_day"].notna())
        ].copy()
        if alloc_day is not None:
            sub = sub[sub["__install_day"] >= alloc_day]
        sub = sub.sort_values("__install_day", ascending=False)

        if not sub.empty:
            inst_day   = int(sub["__install_day"].iloc[0])
            inst_prett = sub["__install_pretty"].iloc[0]
            install_days.append(inst_prett)

            # Fraud: اگر 1025 > خروج (هر دو موجود)، هشدار True
            is_fraud = (test_day is not None and exit_day is not None and test_day > exit_day)
            frauds.append(True if is_fraud else False)

            # پایه تاخیر: نزد پشتیبان → خروج | غیرنزد → 1025
            if is_nazd:
                base = exit_day; bases.append("خروج")
            else:
                base = test_day; bases.append("1025")

            # اگر هشدار یا base ناموجود → تاخیر NA
            if is_fraud or base is None:
                delays.append(pd.NA)
            else:
                diff = days_diff_jalali(base, inst_day)
                if diff is None:
                    delays.append(pd.NA)
                else:
                    late = diff - sla_days(row.get("شهر"))
                    delays.append(int(late) if late>0 else 0)
        else:
            # هنوز تاریخ نصب در install دیده نشده
            install_days.append(pd.NA)
            delays.append(pd.NA)
            bases.append(pd.NA)
            frauds.append(False)

    if not sheet2.empty:
        mask_fill = sheet2["تاریخ نصب"].isna()
        sheet2.loc[mask_fill, "تاریخ نصب"]       = pd.Series(install_days, index=sheet2.index)[mask_fill]
        sheet2["تاخیر روز"]                      = pd.Series(delays, index=sheet2.index)
        sheet2["پایه_تاخیر"]                     = pd.Series(bases, index=sheet2.index)
        sheet2["هشدار_احتمال_تقلب"]              = pd.Series(frauds, index=sheet2.index)

    # 13) آرشیو: نصب‌شده‌های همین اجرا که هشدار=False
    archive = prev_archive.copy()
    installed_now = sheet2[(sheet2["تاریخ نصب"].notna()) & (~sheet2["هشدار_احتمال_تقلب"].fillna(False))].copy()
    if not installed_now.empty:
        archive = pd.concat([archive, installed_now], ignore_index=True)

    # 14) Disabled_Log: جمع‌آوری موارد حذف‌شده به دلیل disable
    disabled_log = pd.DataFrame(disabled_log_rows) if disabled_log_rows else pd.DataFrame(columns=list(df_pending.columns)+["تاریخ غیر فعال"])
    disabled_log = normalize_columns(disabled_log)

    # 15) یکتاسازی Sheet2 بر اساس سریال (آخرین رکورد نگه‌داشته می‌شود)
    if not sheet2.empty:
        sheet2 = sheet2.reset_index(drop=True)
        sheet2["_row"] = sheet2.index
        sheet2 = sheet2.sort_values("_row").drop_duplicates(subset=["سریال پایانه"], keep="last").drop(columns=["_row"])

    # 16) ذخیره خروجی + استایل‌های اکسل + Right-to-Left
    with pd.ExcelWriter(OUTPUT, engine="xlsxwriter") as w:
        df_pending.to_excel(w, index=False, sheet_name="Pending")
        sheet2.to_excel(w, index=False, sheet_name="Installed_Candidates")
        archive.to_excel(w, index=False, sheet_name="Archive")
        disabled_log.to_excel(w, index=False, sheet_name="Disabled_Log")

        # راست‌چین کردن شیت‌ها (مطابق تنظیم sheet-right-to-left در اکسل)
        for sh in ["Pending","Installed_Candidates","Archive","Disabled_Log"]:
            w.sheets[sh].right_to_left()

        # استایل‌های های‌لایت روی شیت 2
        ws2 = w.sheets["Installed_Candidates"]
        cols2 = list(sheet2.columns)
        try:
            warn_idx  = cols2.index("هشدار_احتمال_تقلب")
            delay_idx = cols2.index("تاخیر روز")
        except ValueError:
            warn_idx, delay_idx = None, None

        warn_format  = w.book.add_format({"bg_color": "#F8D7DA", "bold": True})  # قرمز کم‌رنگ برای هشدار
        delay_format = w.book.add_format({"bg_color": "#FFE5B4"})                 # نارنجی ملایم برای تاخیر>0

        nrows = len(sheet2) + 1  # به اضافهٔ هدر
        ncols = len(cols2)

        # سطرهایی که هشدار=True → کل ردیف قرمز ملایم
        if warn_idx is not None and nrows > 1:
            warn_col_letter = col_letter(warn_idx)
            ws2.conditional_format(f"A2:{col_letter(ncols-1)}{nrows}", {
                "type": "formula",
                "criteria": f'=${warn_col_letter}2=TRUE',
                "format": warn_format
            })

        # سلول‌های «تاخیر روز» که >0 هستند → نارنجی ملایم
        if delay_idx is not None and nrows > 1:
            delay_col_letter = col_letter(delay_idx)
            ws2.conditional_format(f"{delay_col_letter}2:{delay_col_letter}{nrows}", {
                "type": "cell",
                "criteria": ">",
                "value": 0,
                "format": delay_format
            })

    print("✅ Done")
    print(f"📄 Output: {OUTPUT}")
    if prev_backup:
        print(f"💾 Backup: {prev_backup}")

# نقطهٔ ورود استاندارد پایتون برای اجرای مستقیم فایل:
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("❌ Error:", e)
        sys.exit(1)
