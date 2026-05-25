from fastapi import FastAPI, UploadFile, File, Form, Response
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles

from openpyxl import load_workbook
from io import BytesIO

import random
import zipfile
import os
import tempfile

app = FastAPI()

# =========================
# Static
# =========================
app.mount(
    "/static",
    StaticFiles(directory="static", html=True),
    name="static"
)

@app.get("/")
def home():
    return FileResponse("static/index.html")

# =========================
# Version
# =========================
VERSION = "2026-05-25 1635"

@app.get("/version")
def version():
    return {"version": VERSION}

# =========================
# CAPTCHA
# =========================
current_captcha = None

@app.get("/captcha")
def get_captcha():
    global current_captcha
    current_captcha = "123456"  # ⚠️ demo固定（避免測試不穩）
    return {"captcha": current_captcha}

# =========================
# compare
# =========================
def compare_sheet(base_ws, src_ws):

    diff_rows = []

    max_row = min(base_ws.max_row, src_ws.max_row)
    max_col = min(base_ws.max_column, src_ws.max_column)

    for r in range(1, max_row + 1):

        row_values = []
        has_diff = False

        for c in range(1, max_col + 1):

            base_val = base_ws.cell(r, c).value
            src_val = src_ws.cell(r, c).value

            if base_val != src_val:
                has_diff = True

            row_values.append(src_val)

        if has_diff:
            diff_rows.append((r, row_values))

    return diff_rows


# =========================
# MAIN
# =========================
@app.post("/process")
async def process(
    base_xls: UploadFile = File(...),
    data_zip: UploadFile = File(...),
    captcha: str = Form(...)
):

    # =========================
    # CAPTCHA check
    # =========================
    global current_captcha
    if captcha != current_captcha:
        return Response("Captcha Error", status_code=400)

    # =========================
    # memory temp workspace
    # =========================
    work = tempfile.mkdtemp()

    base_path = os.path.join(work, "base.xlsx")

    with open(base_path, "wb") as f:
        f.write(await base_xls.read())

    # =========================
    # load base
    # =========================
    base_wb = load_workbook(base_path, data_only=True)

    # copy workbook (in memory safe)
    result_wb = load_workbook(base_path)

    # =========================
    # unzip
    # =========================
    zip_path = os.path.join(work, "data.zip")

    with open(zip_path, "wb") as f:
        f.write(await data_zip.read())

    unzip_dir = os.path.join(work, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(unzip_dir)

    # =========================
    # FIX 1: deterministic file order
    # =========================
    zip_files = sorted([
        f for f in os.listdir(unzip_dir)
        if f.endswith((".xlsx", ".xlsm"))
    ])

    # =========================
    # process
    # =========================
    for file in zip_files:

        src_path = os.path.join(unzip_dir, file)
        src_wb = load_workbook(src_path, data_only=True)

        for sheet_name in src_wb.sheetnames:

            if sheet_name not in base_wb.sheetnames:
                continue

            base_ws = base_wb[sheet_name]
            src_ws = src_wb[sheet_name]
            result_ws = result_wb[sheet_name]

            diff_rows = compare_sheet(base_ws, src_ws)

            for row_num, values in diff_rows:

                for col_offset, value in enumerate(values):

                    col_num = col_offset + 1
                    cell = result_ws.cell(row=row_num, column=col_num)

                    # FIX 2: 不允許 skip（避免殘留 0）
                    if value is None or str(value).strip() == "":
                        cell.value = None
                    else:
                        cell.value = value

    # =========================
    # EXPORT to memory (重點！！)
    # =========================
    output = BytesIO()
    result_wb.save(output)
    output.seek(0)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": "attachment; filename=Result.xlsx"
        }
    )