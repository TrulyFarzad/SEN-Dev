# -*- coding: utf-8 -*-
"""
noInstall.py — رفتار موردنظر:
- Pending از install با «وضعیت نصب = خیر»
- شروع هر اجرا: شیت۲ قبلی را پاکسازی کن و هر ردیفی که «تاریخ نصب» دارد را حذف کن (آرشیو نکن)
- کشف نصب‌های جدید همین اجرا و محاسبه تاریخ نصب/تاخیر → این ردیف‌ها را
  1) در شیت۲ نگه دار  2) به شیت۳ نیز کپی کن
- تاریخ‌ها استاندارد YYYY/MM/DD؛ منطق: تخصیص ≤ 1025 ≤ خروج؛ « - نزد پشتیبان» در خروج
- شیت‌ها Right-to-Left
"""

import sys
import os
import shutil
from datetime import date as _date, date
from pathlib import Path
import pandas as pd
import re

try:
    import xlsxwriter
except Exception:
    print("❌ xlsxwriter نصب نیست. اجرا: pip install xlsxwriter")
    sys.exit(1)

# -------------------- مسیرها --------------------
def get_desktop():
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

# -------------------- Helper ها --------------------
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.replace("ي","ی")
        .str.replace("ك","ک")
        .str.strip()
    )
    return df

def normalize_text(v) -> str:
    if pd.isna(v): return ""
    s = str(v).replace("ي","ی").replace("ك","ک").replace("\u200c","")
    return re.sub(r"\s+"," ", s).strip()

def extract_day_key(v) -> int|None:
    if pd.isna(v): return None
    digits = "".join(ch for ch in str(v) if ch.isdigit())
    if len(digits) < 8: return None
    return int(digits[:8])  # YYYYMMDD (جلالی)

def pretty_jalali(v) -> str|None:
    k = extract_day_key(v)
    if k is None: return None
    y,m,d = k//10000, (k//100)%100, k%100
    return f"{y:04d}/{m:02d}/{d:02d}"

# --- تبدیل جلالی→میلادی برای اختلاف روز ---
def jalali_to_gregorian(jy, jm, jd):
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

from datetime import date as _d
def jalali_key_to_ordinal(key:int) -> int|None:
    y=key//10000; m=(key//100)%100; d=key%100
    try:
        gy,gm,gd = jalali_to_gregorian(y,m,d)
        return _d(gy,gm,gd).toordinal()
    except: return None

def days_diff_jalali(start_key:int|None, end_key:int|None) -> int|None:
    if start_key is None or end_key is None: return None
    s = jalali_key_to_ordinal(start_key); e = jalali_key_to_ordinal(end_key)
    if s is None or e is None: return None
    return e - s

def sla_days(city:str) -> int:
    return 2 if normalize_text(city) == "مشهد" else 5

def backup_prev(path: Path) -> Path|None:
    if not path.exists(): return None
    b = path.with_name(path.stem + _date.today().strftime("_prev_%Y%m%d") + path.suffix)
    shutil.copy2(path, b); return b

def read_prev_triplet(prev_path: Path):
    cols1 = [
        "کد پذیرنده","نام فروشگاه","شهر","آدرس","مدل پایانه","کد پایانه","سریال پایانه",
        "نام خانوادگی پشتیبان","پروژه",
        "تاریخ تخصیص تجهیز","تاریخ تراکنش 1025","خروج",
        "توضیح","مهلت","تاریخ نصب"
    ]
    ext  = cols1 + ["تحویل پست","تاخیر روز"]
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

# -------------------- ورودی‌ها --------------------
def load_inputs():
    f_install = INPUT_DIR/"install.xlsx"
    f_1025    = INPUT_DIR/"1025.xlsx"
    f_exit    = INPUT_DIR/"خروج.xlsx"
    missing   = [p.name for p in (f_install,f_1025,f_exit) if not p.exists()]
    if missing:
        raise FileNotFoundError("فایل‌های ورودی در noInstall/input نیستند: " + ", ".join(missing))
    return (normalize_columns(pd.read_excel(f_install)),
            normalize_columns(pd.read_excel(f_1025)),
            normalize_columns(pd.read_excel(f_exit)))

# -------------------- اندیس‌ها --------------------
def build_1025_index(df_1025, serial_col, date_col):
    tmp = df_1025[[serial_col, date_col]].copy()
    tmp["_day"]    = tmp[date_col].apply(extract_day_key)
    tmp["_pretty"] = tmp[date_col].apply(pretty_jalali)
    tmp = tmp.dropna(subset=["_day"]).sort_values("_day", ascending=False)
    d={}
    for s,grp in tmp.groupby(serial_col):
        d[str(s)] = list(zip(grp["_day"].tolist(), grp["_pretty"].tolist()))
    return d

def build_exit_index(df_exit, serial_col, date_col):
    note_col = "توضیحات" if "توضیحات" in df_exit.columns else None
    cols = [serial_col, date_col] + ([note_col] if note_col else [])
    tmp = df_exit[cols].copy()
    tmp["_day"]    = tmp[date_col].apply(extract_day_key)
    tmp["_pretty"] = tmp[date_col].apply(pretty_jalali)
    if note_col:
        tmp["_pretty_out"] = tmp.apply(
            lambda r: (None if r["_pretty"] is None else (r["_pretty"] + " - نزد پشتیبان"
                        if "نزد پشتیبان" in normalize_text(r[note_col]) else r["_pretty"])), axis=1)
    else:
        tmp["_pretty_out"] = tmp["_pretty"]
    tmp = tmp.dropna(subset=["_day"]).sort_values("_day", ascending=False)
    d={}
    for s,grp in tmp.groupby(serial_col):
        d[str(s)] = list(zip(grp["_day"].tolist(), grp["_pretty_out"].tolist()))
    return d

def pick_after(index_dict, serial:str, min_day:int|None):
    if min_day is None: return None, None
    items = index_dict.get(str(serial))
    if not items: return None, None
    for day, pretty in items:
        if day >= min_day:
            return day, pretty
    return None, None

# -------------------- اجرای اصلی --------------------
def main():
    # 1) ورودی‌ها
    df_install_full, df_1025, df_exit = load_inputs()

    serial_col = "سریال پایانه"
    alloc_col  = "تاریخ تخصیص تجهیز"
    proj_col   = "پروژه"
    status_col = "وضعیت نصب"  # بله/خیر

    for col in [serial_col, alloc_col, proj_col, status_col]:
        if col not in df_install_full.columns:
            raise KeyError(f"ستون «{col}» در install.xlsx یافت نشد.")

    # حذف پروژه فروش
    df_install_full = df_install_full[
        df_install_full[proj_col].apply(lambda x: normalize_text(x) != "پروژه فروش")
    ].copy()

    # Pending فقط وضعیت نصب = خیر
    df_install = df_install_full[
        df_install_full[status_col].apply(lambda x: normalize_text(x) == "خیر")
    ].copy()

    # استاندارد تخصیص
    df_install["__alloc_day"]    = df_install[alloc_col].apply(extract_day_key)
    df_install["__alloc_pretty"] = df_install[alloc_col].apply(pretty_jalali)

    # ستون تاریخ‌ها در 1025 و خروج
    date_col_1025 = next(c for c in df_1025.columns if "تاریخ" in c)
    if serial_col not in df_exit.columns and "سریال" in df_exit.columns:
        df_exit.rename(columns={"سریال": serial_col}, inplace=True)
    exit_date_col = next(c for c in df_exit.columns if "تاریخ" in c)

    # ایندکس‌ها
    idx_1025 = build_1025_index(df_1025, serial_col, date_col_1025)
    idx_exit = build_exit_index(df_exit, serial_col, exit_date_col)

    # 2) ساخت Pending
    rows=[]
    for _, r in df_install.iterrows():
        serial    = str(r.get(serial_col,""))
        alloc_day = r["__alloc_day"]
        alloc_pre = r["__alloc_pretty"]

        t1025_day, t1025_pre = pick_after(idx_1025, serial, alloc_day)
        exit_day,  exit_pre  = pick_after(idx_exit,  serial, t1025_day)

        out = dict(r)
        out["تاریخ تخصیص تجهیز"] = alloc_pre
        out["تاریخ تراکنش 1025"] = t1025_pre
        out["خروج"]              = exit_pre
        rows.append(out)

    df_pending = pd.DataFrame(rows)
    df_pending = normalize_columns(df_pending)

    s1_cols = [
        "کد پذیرنده","نام فروشگاه","شهر","آدرس","مدل پایانه","کد پایانه","سریال پایانه",
        "نام خانوادگی پشتیبان","پروژه",
        "تاریخ تخصیص تجهیز","تاریخ تراکنش 1025","خروج",
        "توضیح","مهلت","تاریخ نصب"
    ]
    for c in s1_cols:
        if c not in df_pending.columns: df_pending[c]=pd.NA
    df_pending = df_pending[s1_cols]

    # 3) نسخه قبلی: بکاپ + خواندن
    prev_backup = backup_prev(OUTPUT)
    prev_pending, prev_sheet2, prev_archive = read_prev_triplet(prev_backup if prev_backup else OUTPUT)

    # ----------- مرحله A: پاکسازی Sheet2 قبلی از نصب‌شده‌ها (حذف کامل؛ آرشیو نکن) -----------
    sheet2 = prev_sheet2.copy()
    if not sheet2.empty:
        sheet2 = sheet2[sheet2["تاریخ نصب"].isna()].copy()

    for c in ["تحویل پست","تاخیر روز"]:
        if c not in sheet2.columns: sheet2[c]=pd.NA

    # ----------- مرحله B: افزودن نصب‌شده‌های جدید (prev_pending - curr_pending) به Sheet2 -----------
    prev_serials = set(prev_pending["سریال پایانه"].astype(str).fillna("")) if not prev_pending.empty else set()
    curr_serials = set(df_pending["سریال پایانه"].astype(str).fillna(""))   if not df_pending.empty else set()
    newly_installed_serials = prev_serials - curr_serials
    if newly_installed_serials:
        new_cands = prev_pending[prev_pending["سریال پایانه"].astype(str).isin(newly_installed_serials)].copy()
        sheet2 = pd.concat([sheet2, new_cands], ignore_index=True)

    # ----------- مرحله C: تکمیل تاریخ نصب و تاخیر برای Sheet2 فعلی -----------
    df_lu = df_install_full.copy()
    if "تاریخ نصب" not in df_lu.columns:
        df_lu["تاریخ نصب"] = pd.NA
    df_lu["__install_day"]    = df_lu["تاریخ نصب"].apply(extract_day_key)
    df_lu["__install_pretty"] = df_lu["تاریخ نصب"].apply(pretty_jalali)

    install_days = []
    delays = []
    for _, row in sheet2.iterrows():
        serial = str(row.get("سریال پایانه","")).strip()
        merch  = str(row.get("کد پذیرنده","")).strip()
        alloc_day = extract_day_key(row.get("تاریخ تخصیص تجهیز"))
        test_day  = extract_day_key(row.get("تاریخ تراکنش 1025"))

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
            diff = days_diff_jalali(test_day, inst_day)
            if diff is None:
                delays.append(pd.NA)
            else:
                late = diff - sla_days(row.get("شهر"))
                delays.append(int(late) if late>0 else 0)
        else:
            install_days.append(pd.NA)
            delays.append(pd.NA)

    if not sheet2.empty:
        mask_fill = sheet2["تاریخ نصب"].isna()
        sheet2.loc[mask_fill, "تاریخ نصب"] = pd.Series(install_days, index=sheet2.index)[mask_fill]
        sheet2["تاخیر روز"] = pd.Series(delays, index=sheet2.index)

    # ----------- مرحله D: آرشیو = prev_archive + نصب‌شده‌های همین اجرا (بدون حذف از شیت۲) -----------
    archive = prev_archive.copy()
    installed_now = sheet2[sheet2["تاریخ نصب"].notna()].copy()
    if not installed_now.empty:
        archive = pd.concat([archive, installed_now], ignore_index=True)

    # یکتاسازی بر اساس سریال برای Sheet2 (آخرین ردیف بماند)
    if not sheet2.empty:
        sheet2 = sheet2.reset_index(drop=True)
        sheet2["_row"] = sheet2.index
        sheet2 = sheet2.sort_values("_row").drop_duplicates(subset=["سریال پایانه"], keep="last").drop(columns=["_row"])

    # 4) ذخیره
    with pd.ExcelWriter(OUTPUT, engine="xlsxwriter") as w:
        df_pending.to_excel(w, index=False, sheet_name="Pending")
        sheet2.to_excel(w, index=False, sheet_name="Installed_Candidates")
        archive.to_excel(w, index=False, sheet_name="Archive")
        for sh in ["Pending","Installed_Candidates","Archive"]:
            w.sheets[sh].right_to_left()

    print("✅ Done")
    print(f"📄 Output: {OUTPUT}")
    if prev_backup:
        print(f"💾 Backup: {prev_backup}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("❌ Error:", e)
        sys.exit(1)
