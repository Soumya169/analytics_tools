import pandas as pd
import numpy as np
import io
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────
# FILE LOADER
# ─────────────────────────────────────────────
def load_file(uploaded):
    file_bytes = uploaded.read()
    df = pd.read_excel(io.BytesIO(file_bytes), keep_default_na=False,
                       na_filter=False, engine='openpyxl')
    df.columns = df.columns.str.strip()
    return df


# ─────────────────────────────────────────────
# DATETIME PARSER
# ─────────────────────────────────────────────
def to_dt(series):
    try:
        s = series.replace("", pd.NaT).replace("NaT", pd.NaT)
        return pd.to_datetime(s, dayfirst=True, errors='coerce',
                              infer_datetime_format=True)
    except Exception:
        return None


# ─────────────────────────────────────────────
# SECONDS → HH:MM  (NO seconds — vectorized)
# ─────────────────────────────────────────────
def sec_to_hms_series(sec_series):
    s = pd.to_numeric(sec_series, errors='coerce').fillna(-1)
    mask  = s >= 0
    total = s.where(mask, 0).astype(np.int64)
    h  = (total // 3600).astype(str).str.zfill(2)
    m  = ((total % 3600) // 60).astype(str).str.zfill(2)
    result = h + ":" + m          # HH:MM only — no seconds
    return result.where(mask, "")


def sec_to_hms(sec):
    if pd.isna(sec) or sec < 0:
        return ""
    sec = int(sec)
    return f"{sec//3600:02d}:{(sec%3600)//60:02d}"


# ─────────────────────────────────────────────
# DIFF → HH:MM
# ─────────────────────────────────────────────
def diff_hms(dt_a, dt_b, label=""):
    try:
        diff_sec = (dt_b - dt_a).dt.total_seconds()
        neg = int((diff_sec < 0).sum())
        return sec_to_hms_series(diff_sec), neg
    except Exception:
        return None, 0


# ─────────────────────────────────────────────
# AUTO COLUMN INDEX
# ─────────────────────────────────────────────
def auto_index(all_cols, name):
    if name in all_cols:
        return all_cols.index(name)
    lower_map = {c.lower().strip(): i for i, c in enumerate(all_cols)}
    return lower_map.get(name.lower(), 0)


# ─────────────────────────────────────────────
# HH:MM → MINUTES  (vectorized)
# ─────────────────────────────────────────────
def hms_to_min_series(series):
    s     = series.astype(str).str.strip()
    split = s.str.split(":", expand=True)
    try:
        h = pd.to_numeric(split[0], errors='coerce')
        m = pd.to_numeric(split[1], errors='coerce')
        minutes = h * 60 + m
        minutes[s.isin(["", "nan", "None", "–", "NaT"])] = np.nan
        return minutes
    except Exception:
        return pd.Series(np.nan, index=series.index, dtype=float)


def hms_to_min(val):
    try:
        parts = str(val).strip().split(":")
        if len(parts) >= 2:
            return int(parts[0]) * 60 + int(parts[1])
        return None
    except Exception:
        return None


# ─────────────────────────────────────────────
# MINUTES → HH:MM
# ─────────────────────────────────────────────
def min_to_hms(m):
    if pd.isna(m):
        return "–"
    total_sec = int(m * 60)
    h  = total_sec // 3600
    mn = (total_sec % 3600) // 60
    return f"{h:02d}:{mn:02d}"


# ─────────────────────────────────────────────
# HH:MM → Excel day fraction  (vectorized)
# ─────────────────────────────────────────────
def hms_to_excel_fraction_series(series):
    mins = hms_to_min_series(series)
    return (mins * 60) / 86400


# ─────────────────────────────────────────────
# BUILD MULTI-SHEET EXCEL  (xlsxwriter — fast)
# ─────────────────────────────────────────────
def build_excel_multi(sheets_dict, tat_cols_set=None, dt_col_names=None):
    tat_cols_set = tat_cols_set or set()
    dt_col_names = dt_col_names or []
    buf = io.BytesIO()

    # Pre-convert Full Data
    if "Full Data" in sheets_dict:
        df_fd = sheets_dict["Full Data"].copy()
        for col in tat_cols_set:
            if col in df_fd.columns:
                df_fd[col] = hms_to_excel_fraction_series(df_fd[col])
        for col in dt_col_names:
            if col in df_fd.columns:
                df_fd[col] = pd.to_datetime(
                    df_fd[col].replace("", pd.NaT),
                    dayfirst=True, errors='coerce'
                ).dt.to_pydatetime()
        sheets_dict["Full Data"] = df_fd

    TIME_STAT_COLS = {"Average", "Median", "Min", "Max"}

    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        wb = writer.book

        for sheet_name, df in sheets_dict.items():
            if df is None or df.empty:
                continue
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]

            # Single header colour — dark blue for all sheets
            fmt_hdr = wb.add_format({
                'bold': True, 'font_name': 'Arial', 'font_size': 10,
                'font_color': '#FFFFFF', 'bg_color': '#1F4E79',
                'align': 'center', 'valign': 'vcenter',
                'text_wrap': True, 'border': 1
            })
            fmt_hdr_tat = wb.add_format({
                'bold': True, 'font_name': 'Arial', 'font_size': 10,
                'font_color': '#FFFFFF', 'bg_color': '#1F4E79',
                'align': 'center', 'valign': 'vcenter',
                'text_wrap': True, 'border': 1
            })
            fmt_even = wb.add_format({
                'font_name': 'Arial', 'font_size': 9,
                'bg_color': '#EBF3FB', 'align': 'center',
                'valign': 'vcenter', 'border': 1
            })
            fmt_odd = wb.add_format({
                'font_name': 'Arial', 'font_size': 9,
                'bg_color': '#FFFFFF', 'align': 'center',
                'valign': 'vcenter', 'border': 1
            })
            fmt_tat_even = wb.add_format({
                'font_name': 'Arial', 'font_size': 9,
                'bg_color': '#E2EFDA', 'align': 'center',
                'valign': 'vcenter', 'border': 1,
                'num_format': '[hh]:mm'
            })
            fmt_tat_odd = wb.add_format({
                'font_name': 'Arial', 'font_size': 9,
                'bg_color': '#E2EFDA', 'align': 'center',
                'valign': 'vcenter', 'border': 1,
                'num_format': '[hh]:mm'
            })
            fmt_time_val = wb.add_format({
                'font_name': 'Arial', 'font_size': 9,
                'bg_color': '#EBF3FB', 'align': 'center',
                'valign': 'vcenter', 'border': 1,
                'num_format': '[hh]:mm'
            })
            fmt_dt = wb.add_format({
                'font_name': 'Arial', 'font_size': 9,
                'bg_color': '#EBF3FB', 'align': 'center',
                'valign': 'vcenter', 'border': 1,
                'num_format': 'dd-mm-yyyy hh:mm:ss'
            })

            col_names = df.columns.tolist()
            is_stats  = sheet_name in ("Overall Stats", "Category Stats",
                                        "Date Wise", "Day of Week",
                                        "Week Wise", "Month Wise")

            # Write headers
            for ci, cn in enumerate(col_names):
                fmt = fmt_hdr_tat if cn in tat_cols_set else fmt_hdr
                ws.write(0, ci, cn, fmt)
            ws.set_row(0, 28)

            # Column widths
            for ci, cn in enumerate(col_names):
                ws.set_column(ci, ci,
                    14 if cn in tat_cols_set else
                    22 if cn in dt_col_names else 20)

            # Data rows
            data_values = df.values.tolist()
            for ri, row_data in enumerate(data_values):
                rn = ri + 1
                base = fmt_even if rn % 2 == 0 else fmt_odd
                for ci, cn in enumerate(col_names):
                    val = row_data[ci]
                    if cn in tat_cols_set:
                        # Convert HH:MM string to fraction if needed
                        if isinstance(val, str) and ":" in str(val):
                            mins = hms_to_min(val)
                            val  = (mins * 60 / 86400) if mins is not None else None
                        ws.write(rn, ci, val, fmt_tat_even)
                    elif is_stats and cn in TIME_STAT_COLS:
                        if isinstance(val, str) and ":" in str(val):
                            mins = hms_to_min(val)
                            val  = (mins * 60 / 86400) if mins is not None else None
                        ws.write(rn, ci, val, fmt_time_val)
                    elif cn in dt_col_names:
                        ws.write(rn, ci, val, fmt_dt)
                    else:
                        ws.write(rn, ci, val, base)

            ws.freeze_panes(1, 0)

    buf.seek(0)
    return buf


# ─────────────────────────────────────────────
# build_excel — single sheet (used by other modules)
# ─────────────────────────────────────────────
def build_excel(result, dt_col_names, tat_cols_set, sheet_name="Result"):
    return build_excel_multi(
        {sheet_name: result},
        tat_cols_set=tat_cols_set,
        dt_col_names=dt_col_names
    )


# ─────────────────────────────────────────────
# PARSE SUMMARY
# ─────────────────────────────────────────────
def parse_summary(cols_info, st_cols):
    for widget, (lbl, dts, cn) in zip(st_cols, cols_info):
        if dts is not None:
            widget.metric(lbl, f"{int(dts.notna().sum())}/{len(dts)}", f"← {cn}")
        else:
            widget.metric(lbl, "Not mapped", "")


# ─────────────────────────────────────────────
# CALCULATE STAGES
# ─────────────────────────────────────────────
def calculate_stages(result, stages, st):
    tat_cols_set = set()
    errors = []
    for col_name, dt_a, dt_b, from_l, to_l in stages:
        if dt_a is not None and dt_b is not None:
            val, neg = diff_hms(dt_a, dt_b, col_name)
            if val is not None:
                result[col_name] = val
                tat_cols_set.add(col_name)
                filled = int((val != "").sum())
                sample = val[val != ""].iloc[0] if filled > 0 else "N/A"
                st.success(
                    f"✅ **{col_name}** ({to_l} − {from_l}) → "
                    f"{filled}/{len(result)} rows | Sample: `{sample}`"
                )
                if neg > 0:
                    st.warning(f"⚠️ {col_name}: {neg} negative rows → blank")
        else:
            errors.append(f"{col_name}: {from_l} or {to_l} not mapped")
    for e in errors:
        st.warning(f"⚠️ {e}")
    return result, tat_cols_set
