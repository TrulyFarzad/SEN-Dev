# -*- coding: utf-8 -*-
"""
noInstall.py — نسخه با جابجایی ستون‌های تاریخ تراکنش 1025 و خروج
"""

import sys
import os
import shutil
from datetime import datetime
from pathlib import Path
import pandas as pd

try:
    import xlsxwriter  # noqa: F401
except Exception:
    print("❌ کتابخانه xlsxwriter نصب نیست. اجرا: pip install xlsxwriter")
    sys.exit(1)


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


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.replace("ي", "ی")
        .str.replace("ك", "ک")
        .str.strip()
    )
    return df


def digits_date(val) -> int | None:
    if pd.isna(val):
        return None
    s = str(val)
    digits = "".join(ch for ch in s if ch.isdigit())
    if not digits:
        return None
    digits = digits[:8]
    try:
        return int(digits)
    except Exception:
        return None


def load_inputs():
    f_install = INPUT_DIR / "install.xlsx"
    f_1025 = INPUT_DIR / "1025.xlsx"
    f_exit = INPUT_DIR / "خروج.xlsx"

    missing = [p.name for p in [f_install, f_1025, f_exit] if not p.exists()]
    if missing:
        raise FileNotFoundError(
            "❌ ورودی‌ها یافت نشدند.\n"
            "- install.xlsx\n- 1025.xlsx\n- خروج.xlsx\n"
            f"مفقود: {', '.join(missing)}"
        )

    return (
        normalize_columns(pd.read_excel(f_install)),
        normalize_columns(pd.read_excel(f_1025)),
        normalize_columns(pd.read_excel(f_exit)),
    )


def backup_prev(path: Path):
    if not path.exists():
        return None
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bpath = path.with_name(path.stem + f"_prev_{stamp}" + path.suffix)
    shutil.copy2(path, bpath)
    return bpath


def read_prev_triplet(prev_path: Path):
    cols_s1 = [
        "کد پذیرنده", "نام فروشگاه", "شهر", "آدرس", "مدل پایانه", "کد پایانه", "سریال پایانه",
        "نام خانوادگی پشتیبان", "پروژه", "تاریخ تخصیص تجهیز", "تاریخ تراکنش 1025", "خروج", "توضیح", "مهلت", "تاریخ نصب"
    ]
    ext_cols = cols_s1 + ["تحویل پست", "تاخیر روز"]

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

    return (
        safe_parse(0, cols_s1),
        safe_parse(1, ext_cols),
        safe_parse(2, ext_cols)
    )


def build_index(df: pd.DataFrame, serial_col: str, date_col: str):
    tmp = df[[serial_col, date_col]].copy()
    tmp["_ts"] = tmp[date_col].apply(digits_date)
    tmp = tmp.dropna(subset=["_ts"]).sort_values(by="_ts", ascending=False)
    dct = {}
    for serial, sub in tmp.groupby(serial_col):
        dct[str(serial)] = list(zip(sub["_ts"], sub[date_col].tolist()))
    return dct


def pick_after(base_ts: int | None, idx_dict: dict, serial: str):
    if base_ts is None:
        return None, None
    items = idx_dict.get(str(serial))
    if not items:
        return None, None
    for ts, raw in items:
        if ts >= base_ts:
            return ts, raw
    return None, None


def main():
    df_install, df_1025, df_exit = load_inputs()

    serial_col = "سریال پایانه"
    alloc_col = "تاریخ تخصیص تجهیز"
    proj_col = "پروژه"

    if proj_col not in df_install.columns:
        raise KeyError("ستون «پروژه» در install.xlsx نیست.")

    # فیلتر پروژه فروش
    df_install = df_install[df_install[proj_col] != "پروژه فروش"].copy()

    date_col_1025 = [c for c in df_1025.columns if "تاریخ" in c][0]
    if serial_col not in df_exit.columns and "سریال" in df_exit.columns:
        df_exit.rename(columns={"سریال": serial_col}, inplace=True)
    exit_date_col = [c for c in df_exit.columns if "تاریخ" in c][0]

    idx_1025 = build_index(df_1025, serial_col, date_col_1025)
    idx_exit = build_index(df_exit, serial_col, exit_date_col)

    rows = []
    for _, row in df_install.iterrows():
        serial = str(row[serial_col])
        alloc_num = digits_date(row[alloc_col])

        t1025_num, t1025_raw = pick_after(alloc_num, idx_1025, serial)
        exit_num, exit_raw = pick_after(t1025_num, idx_exit, serial)

        # اگر خروج توضیح "نزد پشتیبان" داشت
        if exit_raw:
            sub_df = df_exit[(df_exit[serial_col].astype(str) == serial) &
                             (df_exit[exit_date_col] == exit_raw)]
            if not sub_df.empty and "توضیحات" in sub_df.columns:
                if "نزد پشتیبان" in str(sub_df["توضیحات"].iloc[0]):
                    exit_raw = f"{exit_raw} - نزد پشتیبان"

        out = dict(row)
        out["تاریخ تراکنش 1025"] = t1025_raw
        out["خروج"] = exit_raw
        rows.append(out)

    df_pending = pd.DataFrame(rows)
    sheet1_cols = [
        "کد پذیرنده", "نام فروشگاه", "شهر", "آدرس", "مدل پایانه", "کد پایانه", "سریال پایانه",
        "نام خانوادگی پشتیبان", "پروژه", "تاریخ تخصیص تجهیز", "تاریخ تراکنش 1025", "خروج", "توضیح", "مهلت", "تاریخ نصب"
    ]
    for c in sheet1_cols:
        if c not in df_pending.columns:
            df_pending[c] = pd.NA
    df_pending = df_pending[sheet1_cols]

    prev_backup = backup_prev(OUTPUT_FILE)
    prev_pending, prev_installed_candidates, prev_archive = read_prev_triplet(prev_backup if prev_backup else OUTPUT_FILE)

    new_candidates = pd.DataFrame(columns=prev_pending.columns)
    if not prev_pending.empty:
        prev_serials = set(prev_pending[serial_col].astype(str))
        curr_serials = set(df_pending[serial_col].astype(str))
        newly_installed_serials = prev_serials - curr_serials
        if newly_installed_serials:
            new_candidates = prev_pending[prev_pending[serial_col].astype(str).isin(newly_installed_serials)].copy()

    sheet2 = pd.concat([prev_installed_candidates, new_candidates], ignore_index=True)
    for col in ["تحویل پست", "تاخیر روز"]:
        if col not in sheet2.columns:
            sheet2[col] = pd.NA

    finalized_from_prev = pd.DataFrame(columns=sheet2.columns)
    if not prev_installed_candidates.empty and "تاریخ نصب" in prev_installed_candidates.columns:
        finalized_from_prev = prev_installed_candidates[prev_installed_candidates["تاریخ نصب"].notna()].copy()
        if not finalized_from_prev.empty:
            done_serials = set(finalized_from_prev[serial_col].astype(str))
            sheet2 = sheet2[~sheet2[serial_col].astype(str).isin(done_serials)].copy()

    sheet3 = pd.concat([prev_archive, finalized_from_prev], ignore_index=True)

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
    main()
