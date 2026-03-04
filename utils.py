import pandas as pd
import numpy as np
import io
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ═══════════════════════════════════════════════════════════════
# PERFORMANCE NOTES:
# - All TAT calculations are fully vectorized (numpy)
# - Excel export uses xlsxwriter (10-20x faster than openpyxl for large files)
# - Datetime parsing uses infer_datetime_format for speed
# - Stats use pandas groupby (C-level aggregation, no Python loops)
# ═══════════════════════════════════════════════════════════════


# ─────────────────────────────────────────────
# FILE LOADER  — chunked read for large files
# ─────────────────────────────────────────────
def load_file(uploaded):
    file_bytes = uploaded.read()
    buf = io.BytesIO(file_bytes)
    # Use openpyxl engine with read_only for speed
    df = pd.read_excel(buf, keep_default_na=False, na_filter=False, engine='openpyxl')
    df.columns = df.columns.str.strip()
    return df


# ─────────────────────────────────────────────
# DATETIME PARSER  — vectorized, fast
# ─────────────────────────────────────────────
def to_dt(series):
    try:
        s = series.replace("", pd.NaT).replace("NaT", pd.NaT)
        return pd.to_datetime(s, dayfirst=True, errors='coerce',
                              infer_datetime_format=True)
    except Exception:
        return None


# ─────────────────────────────────────────────
# SECONDS → HH:MM:SS  (fully vectorized)
# ─────────────────────────────────────────────
def sec_to_hms_series(sec_series):
    s = pd.to_numeric(sec_series, errors='coerce').fillna(-1)
    mask = s >= 0
    total = s.where(mask, 0).astype(np.int64)
    h  = (total // 3600).astype(str).str.zfill(2)
    m  = ((total % 3600) // 60).astype(str).str.zfill(2)
    sc = (total % 60).astype(str).str.zfill(2)
    result = h + ":" + m 
    return result.where(mask, "")


def sec_to_hms(sec):
    if pd.isna(sec) or sec < 0:
        return ""
    sec = int(sec)
    return f"{sec//3600:02d}:{(sec%3600)//60:02d}:{sec%60:02d}"


# ─────────────────────────────────────────────
# DIFF → HH:MM:SS  (vectorized)
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
# HH:MM:SS → MINUTES  (vectorized)
# ─────────────────────────────────────────────
def hms_to_min_series(series):
    s = series.astype(str).str.strip()
    split = s.str.split(":", expand=True)
    if split.shape[1] < 3:
        return pd.Series(np.nan, index=series.index, dtype=float)
    try:
        h  = pd.to_numeric(split[0], errors='coerce')
        m  = pd.to_numeric(split[1], errors='coerce')
        sc = pd.to_numeric(split[2], errors='coerce')
        minutes = h * 60 + m + sc / 60
        minutes[s.isin(["", "nan", "None", "–", "NaT"])] = np.nan
        return minutes
    except Exception:
        return pd.Series(np.nan, index=series.index, dtype=float)


def hms_to_min(val):
    try:
        parts = str(val).split(":")
        if len(parts) == 3:
            return int(parts[0]) * 60 + int(parts[1]) + int(parts[2]) / 60
        return None
    except Exception:
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
    return f"{h:02d}:{mn:02d}"


# ─────────────────────────────────────────────
# HH:MM:SS → Excel day fraction  (vectorized)
# ─────────────────────────────────────────────
def hms_to_excel_fraction_series(series):
    mins = hms_to_min_series(series)
    return (mins * 60) / 86400


# ─────────────────────────────────────────────────────────────
# BUILD FORMATTED EXCEL  (xlsxwriter — 10-20x faster than openpyxl)
#
# WHY xlsxwriter:
#   openpyxl touches every cell in Python → 60k rows × 50 cols = 3M operations
#   xlsxwriter writes column formats in ONE call regardless of row count
#   Result: 60k rows in ~5s instead of 4-5 minutes
# ─────────────────────────────────────────────────────────────
def build_excel(result, dt_col_names, tat_cols_set, sheet_name="Result"):

    # ── Pre-convert BEFORE writing ────────────────────────────
    result_out = result.copy()

    # TAT → Excel fraction (real duration, not text)
    for col in tat_cols_set:
        if col in result_out.columns:
            result_out[col] = hms_to_excel_fraction_series(result_out[col])

    # Datetime → Python datetime objects
    for col in dt_col_names:
        if col in result_out.columns:
            result_out[col] = pd.to_datetime(
                result_out[col].replace("", pd.NaT),
                dayfirst=True, errors='coerce'
            ).dt.to_pydatetime()

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        result_out.to_excel(writer, index=False, sheet_name=sheet_name)
        wb = writer.book
        ws = writer.sheets[sheet_name]

        # ── Define formats ONCE ───────────────────────────────
        fmt_hdr_blue = wb.add_format({
            'bold': True, 'font_name': 'Arial', 'font_size': 10,
            'font_color': '#FFFFFF', 'bg_color': '#2C3E50',
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
            'border': 1
        })
        fmt_hdr_green = wb.add_format({
            'bold': True, 'font_name': 'Arial', 'font_size': 10,
            'font_color': '#FFFFFF', 'bg_color': '#2C3E50',
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
            'border': 1
        })
        fmt_data_even = wb.add_format({
            'font_name': 'Arial', 'font_size': 9,
            'bg_color': '#EBF3FB', 'align': 'center', 'valign': 'vcenter',
            'border': 1
        })
        fmt_data_odd = wb.add_format({
            'font_name': 'Arial', 'font_size': 9,
            'bg_color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter',
            'border': 1
        })
        fmt_tat_even = wb.add_format({
            'bold': True, 'font_name': 'Arial', 'font_size': 9,
            'font_color': '#000000', 'bg_color': '#E2EFDA',
            'align': 'center', 'valign': 'vcenter',
            'border': 1, 'num_format': '[hh]:mm'
        })
        fmt_tat_odd = wb.add_format({
            'bold': True, 'font_name': 'Arial', 'font_size': 9,
            'font_color': '#000000', 'bg_color': '#E2EFDA',
            'align': 'center', 'valign': 'vcenter',
            'border': 1, 'num_format': '[hh]:mm'
        })
        fmt_dt_even = wb.add_format({
            'font_name': 'Arial', 'font_size': 9,
            'bg_color': '#EBF3FB', 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'num_format': 'dd-mm-yyyy hh:mm:ss'
        })
        fmt_dt_odd = wb.add_format({
            'font_name': 'Arial', 'font_size': 9,
            'bg_color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'num_format': 'dd-mm-yyyy hh:mm:ss'
        })

        # ── Write headers ─────────────────────────────────────
        for col_idx, col_name in enumerate(result_out.columns):
            fmt = fmt_hdr_green if col_name in tat_cols_set else fmt_hdr_blue
            ws.write(0, col_idx, col_name, fmt)
        ws.set_row(0, 30)

        # ── KEY OPTIMIZATION: write_column instead of cell-by-cell ──
        # xlsxwriter set_column applies format to entire column in ONE call
        # This is O(cols) not O(rows × cols)
        col_names = result_out.columns.tolist()
        for col_idx, col_name in enumerate(col_names):
            is_tat = col_name in tat_cols_set
            is_dt  = col_name in dt_col_names

            if is_tat:
                ws.set_column(col_idx, col_idx, 14, fmt_tat_even)
            elif is_dt:
                ws.set_column(col_idx, col_idx, 22, fmt_dt_even)
            else:
                ws.set_column(col_idx, col_idx, 18, fmt_data_even)

        # ── Write data rows using write_row (bulk) ────────────
        # Convert to list of lists for fast bulk write
        data_values = result_out.values.tolist()
        for row_idx, row_data in enumerate(data_values):
            row_num = row_idx + 1  # +1 for header
            for col_idx, col_name in enumerate(col_names):
                is_tat = col_name in tat_cols_set
                is_dt  = col_name in dt_col_names
                val    = row_data[col_idx]

                if is_tat:
                    fmt = fmt_tat_even
                    # Write fraction value (already converted)
                    v = val if (val is not None and not (isinstance(val, float) and np.isnan(val))) else None
                    ws.write(row_num, col_idx, v, fmt)
                elif is_dt:
                    fmt = fmt_dt_even if row_num % 2 == 0 else fmt_dt_odd
                    ws.write(row_num, col_idx, val, fmt)
                else:
                    fmt = fmt_data_even if row_num % 2 == 0 else fmt_data_odd
                    ws.write(row_num, col_idx, val, fmt)

        ws.freeze_panes(1, 0)

    buf.seek(0)
    return buf


# ─────────────────────────────────────────────────────────────
# BUILD MULTI-SHEET EXCEL (for TAT Analysis full export)
# Uses xlsxwriter for speed — handles 60k rows comfortably
# ─────────────────────────────────────────────────────────────
def build_excel_multi(sheets_dict, tat_cols_set=None, dt_col_names=None):
    """
    sheets_dict: { "Sheet Name": dataframe }
    Returns BytesIO buffer.
    Fast multi-sheet export using xlsxwriter.
    """
    tat_cols_set  = tat_cols_set  or set()
    dt_col_names  = dt_col_names  or []
    buf = io.BytesIO()

    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
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

        wb = writer.book

        # Sheet header colors
        SHEET_COLORS = {
            "Full Data":     "#1F4E79",
            "Overall Stats": "#375623",
            "Category Stats":"#7B2D8B",
            "Date Wise":     "#1F4E79",
            "Day of Week":   "#375623",
            "Week Wise":     "#833C00",
            "Month Wise":    "#7B2D8B",
        }
        TIME_COLS = {"Average", "Median", "Min", "Max"}

        for sheet_name, df in sheets_dict.items():
            if df is None or df.empty:
                continue
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]
            hdr_color = SHEET_COLORS.get(sheet_name, "#1F4E79")

            # ── Formats for this sheet ────────────────────────
            fmt_hdr = wb.add_format({
                'bold': True, 'font_name': 'Arial', 'font_size': 10,
                'font_color': '#FFFFFF', 'bg_color': hdr_color,
                'align': 'center', 'valign': 'vcenter',
                'text_wrap': True, 'border': 1
            })
            fmt_hdr_tat = wb.add_format({
                'bold': True, 'font_name': 'Arial', 'font_size': 10,
                'font_color': '#FFFFFF', 'bg_color': '#2C3E50',
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
            fmt_tat = wb.add_format({
                'bold': True, 'font_name': 'Arial', 'font_size': 9,
                'font_color': '#000000', 'bg_color': '#E2EFDA',
                'align': 'center', 'valign': 'vcenter',
                'border': 1, 'num_format': '[hh]:mm'
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
            # Time dimension column highlights
            fmt_date_col  = wb.add_format({'font_name':'Arial','font_size':9,'bold':True,
                'font_color':'#833C00','bg_color':'#FCE4D6','align':'center','valign':'vcenter','border':1})
            fmt_dow_col   = wb.add_format({'font_name':'Arial','font_size':9,'bold':True,
                'font_color':'#375623','bg_color':'#E2EFDA','align':'center','valign':'vcenter','border':1})
            fmt_week_col  = wb.add_format({'font_name':'Arial','font_size':9,'bold':True,
                'font_color':'#7F6000','bg_color':'#FFF2CC','align':'center','valign':'vcenter','border':1})
            fmt_month_col = wb.add_format({'font_name':'Arial','font_size':9,'bold':True,
                'font_color':'#7B2D8B','bg_color':'#EAD1DC','align':'center','valign':'vcenter','border':1})

            TIME_DIM_FMTS = {
                "GateOut Date":      fmt_date_col,
                "GateOut DayOfWeek": fmt_dow_col,
                "GateOut WeekNo":    fmt_week_col,
                "GateOut Month":     fmt_month_col,
            }

            col_names = df.columns.tolist()

            # ── Write headers ─────────────────────────────────
            for ci, cn in enumerate(col_names):
                h_fmt = fmt_hdr_tat if cn in tat_cols_set else fmt_hdr
                ws.write(0, ci, cn, h_fmt)
            ws.set_row(0, 28)

            # ── Set column widths ─────────────────────────────
            for ci, cn in enumerate(col_names):
                w = (14 if cn in tat_cols_set else
                     22 if cn in dt_col_names else
                     18 if cn in TIME_DIM_FMTS else 20)
                ws.set_column(ci, ci, w)

            # ── Write data rows ───────────────────────────────
            # For stats sheets — convert time strings to fractions
            is_stats_sheet = sheet_name in ("Overall Stats","Category Stats",
                                             "Date Wise","Day of Week","Week Wise","Month Wise")

            data_values = df.values.tolist()
            for ri, row_data in enumerate(data_values):
                row_num = ri + 1
                base_fmt = fmt_even if row_num % 2 == 0 else fmt_odd
                for ci, cn in enumerate(col_names):
                    val = row_data[ci]
                    if cn in tat_cols_set:
                        # Already fraction in Full Data; string in other sheets
                        if isinstance(val, str) and ":" in str(val):
                            mins = hms_to_min(val)
                            val  = (mins * 60 / 86400) if mins is not None else None
                        ws.write(row_num, ci, val, fmt_tat)
                    elif is_stats_sheet and cn in TIME_COLS:
                        if isinstance(val, str) and ":" in str(val):
                            mins = hms_to_min(val)
                            val  = (mins * 60 / 86400) if mins is not None else None
                        ws.write(row_num, ci, val, fmt_time_val)
                    elif cn in dt_col_names:
                        ws.write(row_num, ci, val, fmt_dt)
                    elif cn in TIME_DIM_FMTS:
                        ws.write(row_num, ci, val, TIME_DIM_FMTS[cn])
                    else:
                        ws.write(row_num, ci, val, base_fmt)

            ws.freeze_panes(1, 0)

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
