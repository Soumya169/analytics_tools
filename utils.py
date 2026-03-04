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
    df = pd.read_excel(io.BytesIO(file_bytes), keep_default_na=False, na_filter=False)
    df.columns = df.columns.str.strip()
    return df


# ─────────────────────────────────────────────
# DATETIME PARSER
# ─────────────────────────────────────────────
def to_dt(series):
    try:
        s = series.replace("", pd.NaT)
        return pd.to_datetime(s, dayfirst=True, errors='coerce')
    except:
        return None


# ─────────────────────────────────────────────
# SECONDS → HH:MM:SS  (VECTORIZED)
# Old: .apply(sec_to_hms) = 4000 Python calls
# New: numpy array ops   = 1 operation
# ─────────────────────────────────────────────
def sec_to_hms_series(sec_series):
    s = pd.to_numeric(sec_series, errors='coerce').fillna(-1)
    mask = s >= 0
    total = s.where(mask, 0).astype(np.int64)
    h   = total // 3600
    m   = (total % 3600) // 60
    sc  = total % 60
    result = (h.astype(str).str.zfill(2) + ":" +
              m.astype(str).str.zfill(2) + ":" +
              sc.astype(str).str.zfill(2))
    result = result.where(mask, "")
    return result


# scalar fallback
def sec_to_hms(sec):
    if pd.isna(sec) or sec < 0:
        return ""
    sec = int(sec)
    return f"{sec//3600:02d}:{(sec%3600)//60:02d}:{sec%60:02d}"


# ─────────────────────────────────────────────
# DIFF → HH:MM:SS  (uses vectorized version)
# ─────────────────────────────────────────────
def diff_hms(dt_a, dt_b, label=""):
    try:
        diff_sec = (dt_b - dt_a).dt.total_seconds()
        neg = int((diff_sec < 0).sum())
        result = sec_to_hms_series(diff_sec)
        return result, neg
    except:
        return None, 0


# ─────────────────────────────────────────────
# AUTO COLUMN INDEX FOR SELECTBOX
# ─────────────────────────────────────────────
def auto_index(all_cols, name):
    if name in all_cols:
        return all_cols.index(name)
    lower_map = {c.lower().strip(): i for i, c in enumerate(all_cols)}
    return lower_map.get(name.lower(), 0)


# ─────────────────────────────────────────────
# HH:MM:SS → MINUTES  (VECTORIZED)
# Old: .apply(hms_to_min) = 4000 calls × 7 stages = 28000 calls
# New: vectorized str.split = 1 operation per column
# ─────────────────────────────────────────────
def hms_to_min_series(series):
    s = series.astype(str).str.strip()
    split = s.str.split(":", expand=True)
    if split.shape[1] < 3:
        return pd.Series(np.nan, index=series.index)
    try:
        h  = pd.to_numeric(split[0], errors='coerce')
        m  = pd.to_numeric(split[1], errors='coerce')
        sc = pd.to_numeric(split[2], errors='coerce')
        minutes = h * 60 + m + sc / 60
        minutes[s.isin(["", "nan", "None", "–", "NaT"])] = np.nan
        return minutes
    except:
        return pd.Series(np.nan, index=series.index)


# scalar fallback
def hms_to_min(val):
    try:
        parts = str(val).split(":")
        if len(parts) == 3:
            return int(parts[0]) * 60 + int(parts[1]) + int(parts[2]) / 60
        return None
    except:
        return None


# ─────────────────────────────────────────────
# MINUTES → HH:MM:SS
# ─────────────────────────────────────────────
def min_to_hms(m):
    if pd.isna(m):
        return "–"
    total_sec = int(m * 60)
    h  = total_sec // 3600
    mn = (total_sec % 3600) // 60
    s  = total_sec % 60
    return f"{h:02d}:{mn:02d}:{s:02d}"


# ─────────────────────────────────────────────
# HH:MM:SS → Excel day fraction  (VECTORIZED)
# ─────────────────────────────────────────────
def hms_to_excel_fraction_series(series):
    mins = hms_to_min_series(series)
    return (mins * 60) / 86400


# ─────────────────────────────────────────────
# BUILD FORMATTED EXCEL  (OPTIMIZED)
# Key: pre-convert data BEFORE writing → avoid cell-by-cell loops
# ─────────────────────────────────────────────
def build_excel(result, dt_col_names, tat_cols_set, sheet_name="Result"):

    # ── Step 1: Pre-convert all values before writing to Excel ──
    result_excel = result.copy()

    # TAT cols: HH:MM:SS string → Excel day fraction (float)
    for col in tat_cols_set:
        if col in result_excel.columns:
            result_excel[col] = hms_to_excel_fraction_series(result_excel[col])

    # Datetime cols: string → Python datetime
    for col in dt_col_names:
        if col in result_excel.columns:
            result_excel[col] = pd.to_datetime(
                result_excel[col].replace("", pd.NaT),
                dayfirst=True, errors='coerce'
            ).dt.to_pydatetime()

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        result_excel.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]

        n_rows = len(result_excel)
        n_cols = len(result_excel.columns)

        # ── Step 2: Create style objects ONCE (not per cell) ──
        thin        = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'),  bottom=Side(style='thin'))
        center      = Alignment(horizontal="center", vertical="center")
        center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)

        fill_blue  = PatternFill("solid", fgColor="1F4E79")
        fill_green = PatternFill("solid", fgColor="375623")
        fill_even  = PatternFill("solid", fgColor="EBF3FB")
        fill_odd   = PatternFill("solid", fgColor="FFFFFF")
        fill_tat   = PatternFill("solid", fgColor="E2EFDA")

        font_hdr  = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        font_data = Font(name="Arial", size=9,  color="000000")
        font_tat  = Font(name="Arial", size=9,  bold=True, color="375623")

        # ── Step 3: Header row ──
        for idx, col_name in enumerate(result_excel.columns, 1):
            cell           = ws.cell(row=1, column=idx)
            cell.font      = font_hdr
            cell.fill      = fill_green if col_name in tat_cols_set else fill_blue
            cell.alignment = center_wrap
            cell.border    = thin
        ws.row_dimensions[1].height = 30

        # ── Step 4: Build column index sets once ──
        col_pos     = {col: idx for idx, col in enumerate(result_excel.columns, 1)}
        tat_indices = {col_pos[c] for c in tat_cols_set if c in col_pos}
        dt_indices  = {col_pos[c] for c in dt_col_names if c in col_pos}

        # ── Step 5: Style data — iterate COLUMN by column (cache-friendly) ──
        for col_idx in range(1, n_cols + 1):
            is_tat = col_idx in tat_indices
            is_dt  = col_idx in dt_indices
            f_data = font_tat  if is_tat else font_data
            f_fill = fill_tat  if is_tat else None
            nfmt   = "[HH]:MM:SS"   if is_tat else (
                     "DD-MM-YYYY HH:MM:SS" if is_dt else None)

            for row_num in range(2, n_rows + 2):
                cell           = ws.cell(row=row_num, column=col_idx)
                cell.font      = f_data
                cell.alignment = center
                cell.border    = thin
                cell.fill      = f_fill if f_fill else (
                                  fill_even if row_num % 2 == 0 else fill_odd)
                if nfmt and cell.value not in (None, ""):
                    cell.number_format = nfmt

        # ── Step 6: Column widths ──
        for idx, col_name in enumerate(result_excel.columns, 1):
            cl = get_column_letter(idx)
            ws.column_dimensions[cl].width = (
                14 if col_name in tat_cols_set else
                22 if col_name in dt_col_names else 18)

        ws.freeze_panes = "A2"

    buf.seek(0)
    return buf


# ─────────────────────────────────────────────
# PARSE SUMMARY METRICS
# ─────────────────────────────────────────────
def parse_summary(cols_info, st_cols):
    for widget, (lbl, dts, cn) in zip(st_cols, cols_info):
        if dts is not None:
            widget.metric(lbl, f"{int(dts.notna().sum())}/{len(dts)}", f"← {cn}")
        else:
            widget.metric(lbl, "Not mapped", "")


# ─────────────────────────────────────────────
# CALCULATE & SHOW STAGES
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
                    st.warning(f"⚠️ {col_name}: {neg} rows had negative time — set to blank")
        else:
            errors.append(f"{col_name}: {from_l} or {to_l} not mapped")
    for e in errors:
        st.warning(f"⚠️ {e}")
    return result, tat_cols_set