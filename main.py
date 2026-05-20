from fastapi import FastAPI, UploadFile, File
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import zipfile, shutil, os
import pandas as pd
from openpyxl import Workbook
import os, shutil, zipfile, tempfile, glob
from datetime import datetime

import pandas as pd
import numpy as np
from fastapi import UploadFile, File
from fastapi.responses import FileResponse
from openpyxl import load_workbook
from copy import copy

app = FastAPI()

# ✅ static 一定只能掛 /static（安全）
app.mount("/static", StaticFiles(directory="static", html=True), name="static")

# ✅ 首頁
@app.get("/")
def home():
    return FileResponse("static/index.html")

#start
# =========================
# 找到真正有資料的 sheet
# =========================
def find_data_sheet(wb):
    for name in wb.sheetnames:
        ws = wb[name]

        # 判斷 C1 是否有資料（你的固定資料起點）1,3	#A3
        if ws.cell(row=3, column=1).value is not None:
            return ws

    # 找不到就回傳第一個（保底）
    return wb.active

# 🔧 你可以改這裡的處理邏輯
# =========================
def process_sheet(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    """
    每一個 sheet 的處理邏輯
    """
    df = df.copy()

    # 範例：新增一欄 sheet 名稱
    #df["sheet_name"] = sheet_name

    return df


def copy_cell(src_cell, dst_cell):
    dst_cell.value = src_cell.value
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.border = copy(src_cell.border)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.number_format = copy(src_cell.number_format)
        dst_cell.protection = copy(src_cell.protection)
        dst_cell.alignment = copy(src_cell.alignment)
    
def copy_range(src_ws, dst_ws,
               min_row, max_row,
               min_col, max_col,
               target_row, target_col):

    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            src = src_ws.cell(r, c)
            dst = dst_ws.cell(
                target_row + (r - min_row),
                target_col + (c - min_col)
            )
            copy_cell(src, dst)

# =========================
# API
# =========================
@app.post("/process")
async def process(
    base_xls: UploadFile = File(...),
    data_zip: UploadFile = File(...)
):
    work = tempfile.mkdtemp(prefix="work_")

    base_path = os.path.join(work, "base.xlsx")
    zip_path = os.path.join(work, "data.zip")

    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)

    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # 解壓
    extract_dir = os.path.join(work, "extract")
    os.makedirs(extract_dir, exist_ok=True)
    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(extract_dir)

    # 找 Excel
    excel_file = None
    for root, _, files in os.walk(extract_dir):
        for f in files:
            if f.endswith((".xlsx", ".xls")):
                excel_file = os.path.join(root, f)
                break

    # 開啟兩個 Excel（重點）
    src_wb = load_workbook(excel_file)
    src_ws = src_wb.active

    dst_wb = load_workbook(base_path)
    dst_ws = dst_wb.active

    # ⭐ 從 ZIP Excel A3:AK36 複製 → 貼到 base A3:AK36
    copy_range(
        src_ws, dst_ws,
        min_row=3, max_row=36,
        min_col=1, max_col=37,  # A~AK
        target_row=3, target_col=1
    )    # =========================
    # 7. 回傳結果檔
    # =========================
    return FileResponse(
        result_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="result.xlsx"
    )