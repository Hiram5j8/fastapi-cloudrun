from fastapi import FastAPI, UploadFile, File, Form, Response
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from openpyxl import load_workbook
from io import BytesIO

import random
import zipfile
import os
import shutil
import tempfile

# =====================================================
# FastAPI
# =====================================================

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
VERSION = "2026-05-25 1626"

@app.get("/version")
def version():
    return {"version": VERSION}

# =========================
# CAPTCHA (⚠️ demo版)
# =========================
current_captcha = None

@app.get("/captcha")
def get_captcha():
    global current_captcha
    current_captcha = str(random.randint(100000, 999999))
    return {"captcha": current_captcha}

# =========================
# 比較 sheet
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
# Main API
# =========================
@app.post("/process")
async def process(
    base_xls: UploadFile = File(...),
    data_zip: UploadFile = File(...),
    captcha: str = Form(...)
):

    # =========================
    # 0. CAPTCHA check（先做）
    # =========================
    global current_captcha
    if captcha != current_captcha:
        return Response(content="Captcha Error", status_code=400)

    # =========================
    # 1. 工作資料夾
    # =========================
    work = tempfile.mkdtemp(prefix="excel_")

    # =========================
    # 2. 存 base Excel
    # =========================
    base_path = os.path.join(work, "base.xlsx")
    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)

    # result copy
    result_path = os.path.join(work, "Result.xlsx")
    shutil.copy(base_path, result_path)

    # =========================
    # 3. 存 ZIP
    # =========================
    zip_path = os.path.join(work, "data.zip")
    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # =========================
    # 4. 解壓 ZIP
    # =========================
    unzip_dir = os.path.join(work, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(unzip_dir)

    # =========================
    # 5. Load Excel（穩定模式）
    # =========================
    base_wb = load_workbook(base_path, data_only=True)
    result_wb = load_workbook(result_path)

    # =========================
    # 6. FIX：ZIP順序穩定（重點）
    # =========================
    zip_files = sorted([
        f for f in os.listdir(unzip_dir)
        if f.endswith((".xlsx", ".xlsm"))
    ])

    # =========================
    # 7. 核心處理
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

                    # =========================
                    # FIX：避免舊值殘留（0問題來源）
                    # =========================
                    if value is None:
                        cell.value = None
                    else:
                        value_str = str(value).strip()
                        if value_str == "":
                            cell.value = None
                        else:
                            cell.value = value

    # =========================
    # 8. Save
    # =========================
    result_wb.save(result_path)

    # =========================
    # 9. Return file
    # =========================
    return FileResponse(
        path=result_path,
        filename="Result.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
