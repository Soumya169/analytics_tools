import pandas as pd
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
# SECONDS → HH:MM:SS
# ─────────────────────────────────────────────
def sec_to_hms(sec):
    if pd.isna(sec) or sec < 0:
        return ""
    sec = int(sec)
    h = sec // 3600
    m = (sec % 3600) // 60
    s = sec % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


# ─────────────────────────────────────────────
# DIFF → HH:MM:SS COLUMN
# ─────────────────────────────────────────────
def diff_hms(dt_a, dt_b, label=""):
    try:
        diff_sec = (dt_b - dt_a).dt.total_seconds()
        neg = int((diff_sec < 0).sum())
        result = diff_sec.apply(sec_to_hms)
        return result, neg
    except Exception as ex:
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
# HH:MM:SS → MINUTES (for analysis)
# ─────────────────────────────────────────────
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
    h = total_sec // 3600
    mn = (total_sec % 3600) // 60
    s = total_sec % 60
    return f"{h:02d}:{mn:02d}:{s:02d}"


# ─────────────────────────────────────────────
# BUILD FORMATTED EXCEL OUTPUT
# ─────────────────────────────────────────────
def build_excel(result, dt_col_names, tat_cols_set, sheet_name="Result"):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        result.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]

        col_letter_map = {col: get_column_letter(idx) for idx, col in enumerate(result.columns, 1)}

        thin = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'),  bottom=Side(style='thin')
        )
        center = Alignment(horizontal="center", vertical="center")

        # Header styling
        for idx, col_name in enumerate(result.columns, 1):
            cell = ws.cell(row=1, column=idx)
            cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            cell.fill = (PatternFill("solid", start_color="375623")
                         if col_name in tat_cols_set
                         else PatternFill("solid", start_color="1F4E79"))
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin
        ws.row_dimensions[1].height = 30

        # Data rows
        for row_num in range(2, len(result) + 2):
            bg = (PatternFill("solid", start_color="EBF3FB")
                  if row_num % 2 == 0
                  else PatternFill("solid", start_color="FFFFFF"))
            for col_idx, col_name in enumerate(result.columns, 1):
                cell = ws.cell(row=row_num, column=col_idx)
                is_tat = col_name in tat_cols_set
                cell.fill = PatternFill("solid", start_color="E2EFDA") if is_tat else bg
                cell.font = Font(name="Arial", size=9,
                                 bold=is_tat,
                                 color="375623" if is_tat else "000000")
                cell.alignment = center
                cell.border = thin
                if is_tat:
                    # Convert HH:MM:SS string → Excel day fraction for real duration format
                    if isinstance(cell.value, str) and cell.value != "":
                        try:
                            parts = cell.value.strip().split(":")
                            if len(parts) == 3:
                                h, m, s = int(parts[0]), int(parts[1]), int(parts[2])
                                cell.value = (h * 3600 + m * 60 + s) / 86400
                        except:
                            pass
                    if cell.value not in (None, ""):
                        cell.number_format = "[HH]:MM:SS"

        # Datetime columns → real Excel datetime
        date_fmt = "DD-MM-YYYY HH:MM:SS"
        for col_name in dt_col_names:
            if col_name in col_letter_map:
                cl = col_letter_map[col_name]
                for row_num in range(2, len(result) + 2):
                    cell = ws[f"{cl}{row_num}"]
                    if isinstance(cell.value, str) and cell.value != "":
                        try:
                            cell.value = pd.to_datetime(cell.value, dayfirst=True).to_pydatetime()
                        except:
                            pass
                    if cell.value and not isinstance(cell.value, str):
                        cell.number_format = date_fmt

        # Column widths
        for idx, col_name in enumerate(result.columns, 1):
            cl = get_column_letter(idx)
            if col_name in tat_cols_set:
                ws.column_dimensions[cl].width = 14
            elif col_name in dt_col_names:
                ws.column_dimensions[cl].width = 22
            else:
                ws.column_dimensions[cl].width = 18

        ws.freeze_panes = "A2"

    buf.seek(0)
    return buf


# ─────────────────────────────────────────────
# PARSE SUMMARY METRICS (for display)
# ─────────────────────────────────────────────
def parse_summary(cols_info, st_cols):
    for widget, (lbl, dts, cn) in zip(st_cols, cols_info):
        if dts is not None:
            widget.metric(lbl,
                          f"{int(dts.notna().sum())}/{len(dts)}",
                          f"← {cn}")
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

    if errors:
        for e in errors:
            st.warning(f"⚠️ {e}")

    return result, tat_cols_set