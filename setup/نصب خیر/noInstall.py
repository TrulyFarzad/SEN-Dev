# -*- coding: utf-8 -*-
"""
noInstall.py — نسخه با «استانداردسازی تاریخ‌ها»
سه‌شیتی: Pending / Installed_Candidates / Archive

به‌روزرسانی‌ها:
- همه‌ی تاریخ‌های کلیدی در خروجی «استاندارد» می‌شوند به قالب YYYY/MM/DD:
  * تاریخ تخصیص تجهیز  → YYYY/MM/DD
  * تاریخ تراکنش 1025  → YYYY/MM/DD
  * خروج               → YYYY/MM/DD  (و اگر «نزد پشتیبان» بود:  YYYY/MM/DD - نزد پشتیبان)
- ترتیب تاریخ‌ها: تخصیص ≤ 1025 ≤ خروج (در سطح «روز»)
- فیلتر پروژه: ردیف‌های «پروژه فروش» از install حذف می‌شوند.
- شیت‌ها Right-to-Left هستند.
- ورودی‌ها: سه فایل داخل Desktop/noInstall/input  (install.xlsx, 1025.xlsx, خروج.xlsx)
"""

import sys
import os
import shutil
from datetime import datetime
from pathlib import Path
import pandas as pd
import re

# برای Right-to-Left
try:
    import xlsxwriter  # noqa: F401
except Exception:
    print("❌ کتابخانه xlsxwriter نصب نیست. اجرا:  pip install xlsxwriter")
    sys.exit(1)


# -------------------- مسیرها --------------------
def get_desktop():
    home = Path.home()
    candidates = [
        Path(os.environ.get("USERPROFILE", "")) / "Desktop",
        home / "Desktop",
        home,
    ]
    for c in candidates:
        if c.exists():
            return c
    return home

DESKTOP = get_desktop()
BASE_DIR = DESKTOP / "noInstall"
INPUT_DIR = BASE_DIR / "input"
OUTPUT_FILE = BASE_DIR / "install_kheir_output.xlsx"

BASE_DIR.mkdir(parents=True, exist_ok=True)
INPUT_DIR.mkdir(parents=True, exist_ok=True)


# -------------------- Helper ها --------------------
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.replace("ي", "ی")
        .str.replace("ك", "ک")
        .str.strip()
    )
    return df

def normalize_text(val) -> str:
    if pd.isna(val):
        return ""
    s = str(val)
    s = s.replace("ي", "ی").replace("ك", "ک").replace("\u200c", "")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def extract_day_key(val) -> int | None:
    """
    از هر ورودی تاریخ، «فقط روز» به‌صورت عدد YYYYMMDD استخراج می‌کند.
    مثال‌ها:
      14040516                → 14040516
      1404/05/16 07:13:03     → 14040516
      1404-05-16T11:02        → 14040516
    """
    if pd.isna(val):
        return None
    s = str(val)
    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) < 8:
        return None
    try:
        return int(digits[:8])
    except Exception:
        return None

def pretty_jalali_day(val) -> str | None:
    """
    خروجی استاندارد برای نمایش روز: YYYY/MM/DD
    """
    k = extract_day_key(val)
    if k is None:
        return None
    y = k // 10000
    m = (k // 100) % 100
    d = k % 100
    return f"{y:04d}/{m:02d}/{d:02d}"

def load_inputs():
    f_install = INPUT_DIR / "install.xlsx"
    f_1025    = INPUT_DIR / "1025.xlsx"
    f_exit    = INPUT_DIR / "خروج.xlsx"

    missing = [p.name for p in (f_install, f_1025, f_exit) if not p.exists()]
    if missing:
        raise FileNotFoundError(
            "ورودی‌ها یافت نشدند. فایل‌ها را در Desktop/noInstall/input قرار بده:\n"
            "- install.xlsx\n- 1025.xlsx\n- خروج.xlsx\n"
            f"فایل‌های مفقود: {', '.join(missing)}"
        )

    return (
        normalize_columns(pd.read_excel(f_install)),
        normalize_columns(pd.read_excel(f_1025)),
        normalize_columns(pd.read_excel(f_exit)),
    )

def backup_prev(path: Path) -> Path | None:
    if not path.exists():
        return None
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bpath = path.with_name(path.stem + f"_prev_{stamp}" + path.suffix)
    shutil.copy2(path, bpath)
    return bpath

def read_prev_triplet(prev_path: Path):
    cols_s1 = [
        "کد پذیرنده","نام فروشگاه","شهر","آدرس","مدل پایانه","کد پایانه","سریال پایانه",
        "نام خانوادگی پشتیبان","پروژه","تاریخ تخصیص تجهیز","تاریخ تراکنش 1025","خروج","توضیح","مهلت","تاریخ نصب"
    ]
    ext_cols = cols_s1 + ["تحویل پست","تاخیر روز"]

    if not prev_path or not prev_path.exists():
        return pd.DataFrame(columns=cols_s1), pd.DataFrame(columns=ext_cols), pd.DataFrame(columns=ext_cols)

    xls = pd.ExcelFile(prev_path)

    def safe_parse(idx_or_name, cols):
        try:
            df = normalize_columns(xls.parse(idx_or_name))
            for c in cols:
                if c not in df.columns:
                    df[c] = pd.NA
            return df[cols]
        except Exception:
            return pd.DataFrame(columns=cols)

    return safe_parse(0, cols_s1), safe_parse(1, ext_cols), safe_parse(2, ext_cols)

# ایندکس‌های روز-محور
def build_1025_index(df_1025: pd.DataFrame, serial_col: str, date_col: str):
    """
    {serial: [(day_key_desc, pretty_day_str), ...]} نزولی
    """
    tmp = df_1025[[serial_col, date_col]].copy()
    tmp["_day_key"] = tmp[date_col].apply(extract_day_key)
    tmp["_pretty"]  = tmp[date_col].apply(pretty_jalali_day)
    tmp = tmp.dropna(subset=["_day_key"]).sort_values("_day_key", ascending=False)
    idx = {}
    for s, sub in tmp.groupby(serial_col):
        idx[str(s)] = list(zip(sub["_day_key"].tolist(), sub["_pretty"].tolist()))
    return idx

def build_exit_index(df_exit: pd.DataFrame, serial_col: str, date_col: str):
    """
    {serial: [(day_key_desc, pretty_day_str_or_with_note), ...]} نزولی
    اگر در «توضیحات» عبارت «نزد پشتیبان» بود، به انتهای تاریخ « - نزد پشتیبان» افزوده می‌شود.
    """
    note_col = "توضیحات" if "توضیحات" in df_exit.columns else None
    cols = [serial_col, date_col] + ([note_col] if note_col else [])
    tmp = df_exit[cols].copy()
    tmp["_day_key"] = tmp[date_col].apply(extract_day_key)
    tmp["_pretty"]  = tmp[date_col].apply(pretty_jalali_day)

    if note_col:
        def with_note(row):
            b = row["_pretty"]
            if b is None:
                return None
            note = normalize_text(row[note_col])
            return b + " - نزد پشتیبان" if "نزد پشتیبان" in note else b
        tmp["_pretty_out"] = tmp.apply(with_note, axis=1)
    else:
        tmp["_pretty_out"] = tmp["_pretty"]

    tmp = tmp.dropna(subset=["_day_key"]).sort_values("_day_key", ascending=False)
    idx = {}
    for s, sub in tmp.groupby(serial_col):
        idx[str(s)] = list(zip(sub["_day_key"].tolist(), sub["_pretty_out"].tolist()))
    return idx

def pick_after_day(index_dict: dict, serial: str, min_day: int | None):
    """
    اولین رکوردی که day_key >= min_day باشد (به‌دلیل نزولی بودن، «آخرین مطابق شرط» است).
    """
    if min_day is None:
        return None, None
    items = index_dict.get(str(serial))
    if not items:
        return None, None
    for day_key, pretty in items:
        if day_key >= min_day:
            return day_key, pretty
    return None, None


# -------------------- اجرای اصلی --------------------
def main():
    df_install, df_1025, df_exit = load_inputs()

    serial_col = "سریال پایانه"
    alloc_col  = "تاریخ تخصیص تجهیز"
    proj_col   = "پروژه"

    if serial_col not in df_install.columns:
        raise KeyError("ستون «سریال پایانه» در install.xlsx یافت نشد.")
    if alloc_col not in df_install.columns:
        raise KeyError("ستون «تاریخ تخصیص تجهیز» در install.xlsx یافت نشد.")
    if proj_col not in df_install.columns:
        raise KeyError("ستون «پروژه» در install.xlsx یافت نشد.")

    # حذف «پروژه فروش»
    df_install = df_install[df_install[proj_col].apply(lambda x: normalize_text(x) != "پروژه فروش")].copy()

    # ستون تاریخ‌ها در 1025 و خروج
    # (اولین ستونی که شامل «تاریخ» باشد را می‌گیریم)
    date_col_1025 = next(c for c in df_1025.columns if "تاریخ" in c)
    if serial_col not in df_exit.columns and "سریال" in df_exit.columns:
        df_exit.rename(columns={"سریال": serial_col}, inplace=True)
    exit_date_col = next(c for c in df_exit.columns if "تاریخ" in c)

    # ایندکس‌ها
    idx_1025 = build_1025_index(df_1025, serial_col, date_col_1025)
    idx_exit = build_exit_index(df_exit, serial_col, exit_date_col)

    # استخراج و استانداردسازی
    rows = []
    for _, r in df_install.iterrows():
        serial = str(r.get(serial_col, ""))
        alloc_day_key = extract_day_key(r.get(alloc_col))
        alloc_pretty  = pretty_jalali_day(r.get(alloc_col))  # استاندارد خروجی تخصیص

        # 1025 پس از تخصیص
        test_day_key, test_pretty = pick_after_day(idx_1025, serial, alloc_day_key)

        # خروج پس از 1025 (اگر 1025 نبود، خروج را ست نمی‌کنیم)
        exit_day_key, exit_pretty = pick_after_day(idx_exit, serial, test_day_key)

        out = dict(r)
        out["تاریخ تخصیص تجهیز"] = alloc_pretty             # استاندارد‌شده
        out["تاریخ تراکنش 1025"] = test_pretty              # استاندارد‌شده
        out["خروج"]              = exit_pretty              # استاندارد‌شده (+ « - نزد پشتیبان» در صورت نیاز)
        rows.append(out)

    df_pending = pd.DataFrame(rows)
    df_pending = normalize_columns(df_pending)

    # چینش و تکمیل ستون‌ها
    sheet1_cols = [
        "کد پذیرنده","نام فروشگاه","شهر","آدرس","مدل پایانه","کد پایانه","سریال پایانه",
        "نام خانوادگی پشتیبان","پروژه",
        "تاریخ تخصیص تجهیز","تاریخ تراکنش 1025","خروج",   # ← ترتیب جدید
        "توضیح","مهلت","تاریخ نصب"
    ]
    for c in sheet1_cols:
        if c not in df_pending.columns:
            df_pending[c] = pd.NA
    df_pending = df_pending[sheet1_cols]

    # نسخه قبلی
    prev_backup = backup_prev(OUTPUT_FILE)
    prev_pending, prev_installed_candidates, prev_archive = read_prev_triplet(prev_backup if prev_backup else OUTPUT_FILE)

    # شیت ۲: کسانی که از Pending قبلی حذف شده‌اند
    new_candidates = pd.DataFrame(columns=prev_pending.columns)
    if not prev_pending.empty:
        prev_serials = set(prev_pending["سریال پایانه"].astype(str).fillna(""))
        curr_serials = set(df_pending["سریال پایانه"].astype(str).fillna(""))
        newly_installed_serials = prev_serials - curr_serials
        if newly_installed_serials:
            new_candidates = prev_pending[prev_pending["سریال پایانه"].astype(str).isin(newly_installed_serials)].copy()

    sheet2 = pd.concat([prev_installed_candidates, new_candidates], ignore_index=True)
    for col in ["تحویل پست","تاخیر روز"]:
        if col not in sheet2.columns:
            sheet2[col] = pd.NA
    if not sheet2.empty:
        sheet2 = sheet2.reset_index(drop=True)
        sheet2["_ROW"] = sheet2.index
        sheet2 = sheet2.sort_values("_ROW").drop_duplicates(subset=["سریال پایانه"], keep="last").drop(columns=["_ROW"])

    # شیت ۳: آرشیو
    finalized_from_prev = pd.DataFrame(columns=sheet2.columns)
    if not prev_installed_candidates.empty and "تاریخ نصب" in prev_installed_candidates.columns:
        finalized_from_prev = prev_installed_candidates[prev_installed_candidates["تاریخ نصب"].notna()].copy()
        if not finalized_from_prev.empty:
            done_serials = set(finalized_from_prev["سریال پایانه"].astype(str))
            sheet2 = sheet2[~sheet2["سریال پایانه"].astype(str).isin(done_serials)].copy()

    sheet3 = pd.concat([prev_archive, finalized_from_prev], ignore_index=True)

    # ذخیره + Right-to-Left
    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        df_pending.to_excel(writer, index=False, sheet_name="Pending")
        sheet2.to_excel(writer, index=False, sheet_name="Installed_Candidates")
        sheet3.to_excel(writer, index=False, sheet_name="Archive")

        for sh in ["Pending", "Installed_Candidates", "Archive"]:
            writer.sheets[sh].right_to_left()

    print("✅ Done")
    print(f"📄 Output: {OUTPUT_FILE}")
    if prev_backup:
        print(f"💾 Backup: {prev_backup}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("❌ Error:", e)
        sys.exit(1)
