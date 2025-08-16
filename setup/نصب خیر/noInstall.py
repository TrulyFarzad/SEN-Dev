# -*- coding: utf-8 -*-
"""
noInstall.py — نصب‌خیر با:
- نگهداری «توضیح» بین اجراها
- پرچم «از_نزد_پشتیبان»، «پایه_تاخیر»
- Fraud detection: اگر 1025 > خروج → هشدار و عدم آرشیو
- حذف خودکار «غیرفعال‌شده‌ها» (disable.xlsx) از Pending و Sheet2 (بدون تاریخ نصب) + ثبت در Disabled_Log
- استایل رنگی برای هشدار و تاخیر
- شیت‌ها Right-to-Left
"""

import sys, os, shutil, re
from datetime import date as _date, date
from pathlib import Path
import pandas as pd

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
    df.columns = df.columns.astype(str).str.replace("ي","ی").str.replace("ك","ک").str.strip()
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

def jalali_key_to_ordinal(key:int) -> int|None:
    y=key//10000; m=(key//100)%100; d=key%100
    try:
        gy,gm,gd = jalali_to_gregorian(y,m,d)
        from datetime import date as _d
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

def build_1025_index(df_1025, serial_col, date_col):
    tmp = df_1025[[serial_col, date_col]].copy()
    tmp["_day"]    = tmp[date_col].apply(extract_day_key)
    tmp["_pretty"] = tmp[date_col].apply(pretty_jalali)
    tmp = tmp.dropna(subset=["_day"]).sort_values("_day", ascending=False)
    d={}
    for s,grp in tmp.groupby(serial_col):
        d[str(s)] = list(zip(grp["_day"].tolist(), grp["_pretty"].tolist()))
    return d

def build_exit_index_with_flag(df_exit, serial_col, date_col):
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
    اندیس disable بر اساس «تاریخ پایان تخصیص» (در صورت نبود، fallback هوشمند)
    خروجی: dict[serial] = [(day_key, pretty_str, merchant_code_str), ...]  (جدیدترین اول)
    """
    # ستون تاریخ هدف
    date_col = "تاریخ پایان تخصیص"
    if date_col not in df_disable.columns:
        # fallback: هر ستونی که «پایان تخصیص» در نام دارد
        cand = [c for c in df_disable.columns if "پایان تخصیص" in c]
        if cand:
            date_col = cand[0]
        else:
            # در نهایت اولین ستونی که «تاریخ» دارد
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

def pick_exit_after_alloc(exit_idx:dict, serial:str, alloc_day:int|None):
    if alloc_day is None: return None, None, False
    items = exit_idx.get(str(serial))
    if not items: return None, None, False
    # اول نزد پشتیبان
    for day, pretty, is_nazd in items:
        if day >= alloc_day and is_nazd:
            return day, pretty, True
    # سپس اولین خروج بعد از تخصیص
    for day, pretty, is_nazd in items:
        if day >= alloc_day:
            return day, pretty, False
    return None, None, False

def pick_1025_after_alloc(idx_1025:dict, serial:str, alloc_day:int|None):
    if alloc_day is None: return None, None
    items = idx_1025.get(str(serial))
    if not items: return None, None
    for day, pretty in items:
        if day >= alloc_day:
            return day, pretty
    return None, None

# Excel column letter helper
def col_letter(idx_zero_based:int) -> str:
    s = ""
    n = idx_zero_based + 1
    while n:
        n, rem = divmod(n-1, 26)
        s = chr(65+rem) + s
    return s

def coalesce_text(a, b):
    a_ = normalize_text(a)
    b_ = normalize_text(b)
    return a if a_ != "" else (b if b_ != "" else a)

# -------------------- اجرای اصلی --------------------
def main():
    df_install_full, df_1025, df_exit, df_disable = load_inputs()

    serial_col = "سریال پایانه"
    alloc_col  = "تاریخ تخصیص تجهیز"
    proj_col   = "پروژه"
    status_col = "وضعیت نصب"

    for col in [serial_col, alloc_col, proj_col, status_col]:
        if col not in df_install_full.columns:
            raise KeyError(f"ستون «{col}» در install.xlsx نیست.")

    # حذف پروژه فروش
    df_install_full = df_install_full[df_install_full[proj_col].apply(lambda x: normalize_text(x)!="پروژه فروش")].copy()

    # Pending فقط وضعیت نصب = خیر
    df_install = df_install_full[df_install_full[status_col].apply(lambda x: normalize_text(x)=="خیر")].copy()

    # استاندارد تخصیص
    df_install["__alloc_day"]    = df_install[alloc_col].apply(extract_day_key)
    df_install["__alloc_pretty"] = df_install[alloc_col].apply(pretty_jalali)

    # ایندکس‌ها
    date_col_1025 = next(c for c in df_1025.columns if "تاریخ" in c)
    if serial_col not in df_exit.columns and "سریال" in df_exit.columns:
        df_exit.rename(columns={"سریال": serial_col}, inplace=True)
    exit_date_col = next(c for c in df_exit.columns if "تاریخ" in c)

    idx_1025    = build_1025_index(df_1025, serial_col, date_col_1025)
    idx_exit    = build_exit_index_with_flag(df_exit, serial_col, exit_date_col)
    idx_disable = build_disable_index(df_disable, serial_col)

    # Pending
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

    s1_cols = ["کد پذیرنده","نام فروشگاه","شهر","آدرس","مدل پایانه","کد پایانه","سریال پایانه",
               "نام خانوادگی پشتیبان","پروژه",
               "تاریخ تخصیص تجهیز","تاریخ تراکنش 1025","خروج","از_نزد_پشتیبان",
               "توضیح","مهلت","تاریخ نصب"]
    for c in s1_cols:
        if c not in df_pending.columns: df_pending[c]=pd.NA
    df_pending = df_pending[s1_cols]

    # نسخه قبلی
    prev_backup = backup_prev(OUTPUT)
    prev_pending, prev_sheet2, prev_archive = read_prev_triplet(prev_backup if prev_backup else OUTPUT)

    # --- نگهداری توضیحات: merge روی سریال پایانه ---
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

    # ---------------- حذف از Pending بر اساس disable ----------------
    disabled_log_rows = []
    if not df_pending.empty:
        keep_mask = []
        for _, row in df_pending.iterrows():
            serial = str(row["سریال پایانه"]).strip()
            merch  = str(row.get("کد پذیرنده","")).strip()
            alloc_day = extract_day_key(row.get("تاریخ تخصیص تجهیز"))
            dis_items = idx_disable.get(serial, [])
            picked = None
            # جدیدترین غیرفعال بعد از تخصیص، با ترجیح match کد پذیرنده
            for dday, dpretty, dmerch in dis_items:
                if alloc_day is not None and dday >= alloc_day and (merch=="" or dmerch==merch):
                    picked = (dday, dpretty); break
            if picked is None:
                keep_mask.append(True)
            else:
                log = dict(row)
                log["تاریخ غیر فعال"] = picked[1]  # pretty از «تاریخ پایان تخصیص»
                disabled_log_rows.append(log)
                keep_mask.append(False)
        df_pending = df_pending[keep_mask].copy()

    # شروع اجرا: شیت۲ قبلی را از نصب‌شده‌های بدون هشدار پاک کن
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

    # افزودن نصب‌شده‌های جدید (prev_pending - curr_pending)
    prev_serials = set(prev_pending["سریال پایانه"].astype(str).fillna("")) if not prev_pending.empty else set()
    curr_serials = set(df_pending["سریال پایانه"].astype(str).fillna(""))   if not df_pending.empty else set()
    newly_installed_serials = prev_serials - curr_serials
    if newly_installed_serials:
        new_cands = prev_pending[prev_pending["سریال پایانه"].astype(str).isin(newly_installed_serials)].copy()
        sheet2 = pd.concat([sheet2, new_cands], ignore_index=True)

    # ---------------- حذف از Sheet2 بر اساس disable (فقط بدون تاریخ نصب) ----------------
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
                log = dict(row)
                log["تاریخ غیر فعال"] = picked[1]
                disabled_log_rows.append(log)
                keep_mask2.append(False)
        sheet2 = sheet2[keep_mask2].copy()

    # تکمیل «تاریخ نصب» و «تاخیر» و «پایه_تاخیر» + Fraud
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
        is_nazd   = str(row.get("از_نزد_پشتیبان","")).strip().lower() in ("true","1","بله","yes")

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

            is_fraud = (test_day is not None and exit_day is not None and test_day > exit_day)
            frauds.append(True if is_fraud else False)

            if is_nazd:
                base = exit_day; bases.append("خروج")
            else:
                base = test_day; bases.append("1025")

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
            install_days.append(pd.NA)
            delays.append(pd.NA)
            bases.append(pd.NA)
            frauds.append(False)

    if not sheet2.empty:
        mask_fill = sheet2["تاریخ نصب"].isna()
        sheet2.loc[mask_fill, "تاریخ نصب"] = pd.Series(install_days, index=sheet2.index)[mask_fill]
        sheet2["تاخیر روز"] = pd.Series(delays, index=sheet2.index)
        sheet2["پایه_تاخیر"] = pd.Series(bases, index=sheet2.index)
        sheet2["هشدار_احتمال_تقلب"] = pd.Series(frauds, index=sheet2.index)

    # آرشیو: فقط نصب‌شده‌های همین اجرا که هشدار=False
    archive = prev_archive.copy()
    installed_now = sheet2[(sheet2["تاریخ نصب"].notna()) & (~sheet2["هشدار_احتمال_تقلب"].fillna(False))].copy()
    if not installed_now.empty:
        archive = pd.concat([archive, installed_now], ignore_index=True)

    # Disabled_Log شیت چهارم
    disabled_log = pd.DataFrame(disabled_log_rows) if disabled_log_rows else pd.DataFrame(columns=list(df_pending.columns)+["تاریخ غیر فعال"])
    disabled_log = normalize_columns(disabled_log)

    # یکتاسازی Sheet2
    if not sheet2.empty:
        sheet2 = sheet2.reset_index(drop=True)
        sheet2["_row"] = sheet2.index
        sheet2 = sheet2.sort_values("_row").drop_duplicates(subset=["سریال پایانه"], keep="last").drop(columns=["_row"])

    # -------------------- ذخیره + استایل اکسل --------------------
    with pd.ExcelWriter(OUTPUT, engine="xlsxwriter") as w:
        df_pending.to_excel(w, index=False, sheet_name="Pending")
        sheet2.to_excel(w, index=False, sheet_name="Installed_Candidates")
        archive.to_excel(w, index=False, sheet_name="Archive")
        disabled_log.to_excel(w, index=False, sheet_name="Disabled_Log")

        for sh in ["Pending","Installed_Candidates","Archive","Disabled_Log"]:
            w.sheets[sh].right_to_left()

        # های‌لایت‌ها روی Sheet2
        ws2 = w.sheets["Installed_Candidates"]
        cols2 = list(sheet2.columns)
        try:
            warn_idx = cols2.index("هشدار_احتمال_تقلب")
            delay_idx = cols2.index("تاخیر روز")
        except ValueError:
            warn_idx, delay_idx = None, None

        warn_format  = w.book.add_format({"bg_color": "#F8D7DA", "bold": True})
        delay_format = w.book.add_format({"bg_color": "#FFE5B4"})

        nrows = len(sheet2) + 1
        ncols = len(cols2)

        if warn_idx is not None and nrows > 1:
            warn_col_letter = col_letter(warn_idx)
            ws2.conditional_format(f"A2:{col_letter(ncols-1)}{nrows}", {
                "type": "formula",
                "criteria": f'=${warn_col_letter}2=TRUE',
                "format": warn_format
            })
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

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("❌ Error:", e)
        sys.exit(1)
