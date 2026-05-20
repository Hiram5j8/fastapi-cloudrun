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

        # 判斷 C1 是否有資料（你的固定資料起點）
        if ws.cell(row=1, column=3).value is not None:
            return ws

    # 找不到就回傳第一個（保底）
    return wb.active


# =========================
# API
# =========================
@app.post("/process")
async def process(
    base_xls: UploadFile = File(...),
    data_zip: UploadFile = File(...)
):
    # 建立暫存工作資料夾
    work = tempfile.mkdtemp(prefix="work_")

    # =========================
    # 1. 存 base.xlsx
    # =========================
    base_path = os.path.join(work, "base.xlsx")
    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)

    # 複製成 result.xlsx
    result_path = os.path.join(work, "result.xlsx")
    shutil.copy(base_path, result_path)

    # =========================
    # 2. 存 zip
    # =========================
    zip_path = os.path.join(work, "data.zip")
    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # 解壓縮
    unzip_dir = os.path.join(work, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(unzip_dir)

    # =========================
    # 3. 開啟目標 Excel
    # =========================
    dst_wb = load_workbook(result_path)
    dst_ws = dst_wb.active

    # 從最後一列往下開始貼
    paste_row = dst_ws.max_row + 2

    # =========================
    # 4. 找所有 Excel
    # =========================
    files = glob.glob(
        os.path.join(unzip_dir, "**", "*.xlsx"),
        recursive=True
    )

    # =========================
    # 5. 合併 Excel
    # =========================
    for file in files:
        src_wb = load_workbook(file, data_only=True)

        # ⭐ 修正重點：抓真正有資料的 sheet
        src_ws = find_data_sheet(src_wb)

        # 固定範圍：C1 ~ AK36
        for r in range(1, 37):
            for c in range(3, 38):
                val = src_ws.cell(row=r, column=c).value

                dst_ws.cell(
                    row=paste_row + r - 1,
                    column=c,
                    value=val
                )

        # 每個檔案往下空 40 列
        paste_row += 40

    # =========================
    # 6. 存檔
    # =========================
    dst_wb.save(result_path)

    # =========================
    # 7. 回傳結果檔
    # =========================
    return FileResponse(
        result_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="result.xlsx"
    )