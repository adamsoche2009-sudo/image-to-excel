import io
import json
import re

import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageOps

MODELS = ["gemini-3.5-flash", "gemini-3.7-flash", "gemini-3.8-flash"]
PROMPT = """You are an expert at turning images into clean Excel data.
Look carefully at the image and extract ALL tables, lists and data in it.
Rules:
- Copy text and numbers exactly. Do not invent, guess or skip anything.
- Keep the original column order and row order.
- Write numbers as plain JSON numbers (no currency symbols or thousands separators).
  If there is a currency or unit, put it in the column header.
- Keep dates and codes as text exactly as shown.
- If the image has several tables, make one sheet per table.
- If there is no clear table, organise the content into sensible columns.
Reply ONLY with JSON in this shape:
{"sheets":[{"name":"Sheet name","headers":["Col1","Col2"],"rows":[["a",1],["b",2]]}]}"""


def extract(image):
    client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"], http_options=types.HttpOptions(timeout=90000))
    last_error = None
    for model in MODELS:
        try:
            response = client.models.generate_content(
                model=model,
                contents=[image, PROMPT],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json", temperature=0
                ),
            )
            data = json.loads(response.text)
            if not data.get("sheets"):
                raise ValueError("No data found in the image.")
            return data
        except Exception as e:
            last_error = e
    raise last_error


def unique_headers(headers):
    seen, out = {}, []
    for n, h in enumerate(headers, 1):
        h = str(h).strip() or f"Column {n}"
        if h in seen:
            seen[h] += 1
            h = f"{h} ({seen[h]})"
        else:
            seen[h] = 1
        out.append(h)
    return out


def make_df(sheet):
    headers = list(sheet.get("headers") or [])
    rows = [list(r) for r in (sheet.get("rows") or [])]
    width = max([len(headers)] + [len(r) for r in rows])
    headers = unique_headers(headers + [""] * (width - len(headers)))
    fixed = []
    for r in rows:
        r = r + [None] * (width - len(r))
        fixed.append(["" if v is None else str(v) for v in r])
    return pd.DataFrame(fixed, columns=headers)


def to_cell(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    v = str(v).strip()
    if v == "":
        return None
    digits = v.lstrip("-")
    if re.fullmatch(r"-?\d+", v) and not (len(digits) > 1 and digits.startswith("0")):
        return int(v)
    if re.fullmatch(r"-?\d+\.\d+", v):
        return float(v)
    return v


def build_excel(sheets):
    wb = Workbook()
    wb.remove(wb.active)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", start_color="1F4E78")
    for i, s in enumerate(sheets, 1):
        name = "".join(c for c in str(s["name"] or "") if c not in '[]:*?/\\')[:31]
        ws = wb.create_sheet(name or f"Sheet{i}")
        df = s["df"]
        ws.append(list(df.columns))
        for row in df.itertuples(index=False):
            ws.append([to_cell(v) for v in row])
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = head_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
        for col in ws.columns:
            longest = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(longest + 2, 10), 50)
        ws.freeze_panes = None
        if ws.max_row > 1:
            ws.auto_filter.ref = ws.dimensions
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

def build_excel(sheets):
    wb = Workbook()
    wb.remove(wb.active)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", start_color="1F4E78")
    total_fill = PatternFill("solid", start_color="DDEBF7")
    for i, s in enumerate(sheets, 1):
        name = "".join(c for c in str(s["name"] or "") if c not in '[]:*?/\\')[:31]
        ws = wb.create_sheet(name or f"Sheet{i}")
        df = s["df"]
        ws.append(list(df.columns))
        for row in df.itertuples(index=False):
            ws.append([to_cell(v) for v in row])
        last = ws.max_row
        ncols = ws.max_column
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = head_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for col in ws.columns:
            longest = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(longest + 2, 10), 50)
        if last > 1:
            sum_cols = []
            for c in range(1, ncols + 1):
                vals = [ws.cell(r, c).value for r in range(2, last + 1)]
                nums = [v for v in vals if v is not None]
                is_num = len(nums) > 0 and all(isinstance(v, (int, float)) for v in nums)
                is_serial = vals == list(range(1, len(vals) + 1))
                if is_num and not is_serial:
                    sum_cols.append(c)
            if sum_cols:
                t = last + 1
                others = [c for c in range(1, ncols + 1) if c not in sum_cols]
                if others:
                    ws.cell(t, others[0], "TOTAL")
                for c in sum_cols:
                    letter = get_column_letter(c)
                    ws.cell(t, c, f"=SUM({letter}2:{letter}{last})")
                    has_float = any(isinstance(ws.cell(r, c).value, float) for r in range(2, last + 1))
                    fmt = "#,##0.00" if has_float else "#,##0"
                    for r in range(2, t + 1):
                        ws.cell(r, c).number_format = fmt
                    width = ws.column_dimensions[letter].width or 10
                    ws.column_dimensions[letter].width = max(width, 16)
                for c in range(1, ncols + 1):
                    cell = ws.cell(t, c)
                    cell.font = Font(bold=True)
                    cell.fill = total_fill
        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
        ws.freeze_panes = None
        if last > 1:
            ws.auto_filter.ref = f"A1:{get_column_letter(ncols)}{last}"
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


st.set_page_config(page_title="Image to Excel", page_icon="📊")
st.title("📊 Image to Excel")
st.write("Upload a photo or screenshot of a table, list, receipt or notes.")

file = st.file_uploader("Choose an image", type=["png", "jpg", "jpeg", "webp"])

if file:
    key = f"{file.name}-{file.size}"
    if st.session_state.get("file_key") != key:
        st.session_state.pop("result", None)
        st.session_state["file_key"] = key

    image = ImageOps.exif_transpose(Image.open(file)).convert("RGB")
    image.thumbnail((2400, 2400))
    st.image(image)

    if st.button("Convert to Excel", type="primary"):
        with st.spinner("Reading your image..."):
            try:
                st.session_state["result"] = extract(image)
                st.session_state["run"] = st.session_state.get("run", 0) + 1
            except Exception as e:
                st.error(f"Could not read the image: {e}")

    if "result" in st.session_state:
        st.success("Check the table below. You can fix any cell, then download.")
        edited = []
        for i, s in enumerate(st.session_state["result"]["sheets"]):
            st.subheader(s.get("name") or f"Sheet {i + 1}")
            df = st.data_editor(
                make_df(s),
                num_rows="dynamic",
                key=f"sheet_{st.session_state['run']}_{i}",
            )
            edited.append({"name": s.get("name"), "df": df})
        st.download_button(
            "⬇️ Download Excel file",
            data=build_excel(edited),
            file_name="converted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
st.divider()
st.caption("© 2026 Created by Daniel Oche Adams")
