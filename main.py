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
     # A3:AK38  → row 2~37, col 0~36 (0-based)
    sub_df = df.iloc[2:38, 0:37].copy()

    return sub_df
    #df = df.copy()

    # 範例：新增一欄 sheet 名稱
    #df["sheet_name"] = sheet_name

    #return df

def process_and_paste(df: pd.DataFrame, ws_result):
    """
    核心邏輯：
    ZIP A3 比對 result 每5列一組的 A欄
    相同則把 ZIP B3:AK8 貼到該區塊
    """

    # ZIP 的 A3
    zip_a3 = str(df.iloc[2, 0]).strip()

    # ZIP 要貼的區塊 B3:AK8
    block = df.iloc[2:8, 1:37].values  # 6列 x 36欄

    check_row = 3  # 從 A3 開始

    while True:
        cell_val = ws_result.cell(row=check_row, column=1).value

        if cell_val is None:
            break

        if str(cell_val).strip() == zip_a3:
            # 貼到該區塊 B~AK
            for r in range(6):        # 6列
                for c in range(36):   # 36欄
                    ws_result.cell(
                        row=check_row + r,
                        column=2 + c,
                        value=block[r][c]
                    )

        check_row += 5  # 下一組
        
# =========================
# API
# =========================
@app.post("/process")
async def process(
    base_xls: UploadFile = File(...),
    data_zip: UploadFile = File(...)
):
   # 1. 建立工作資料夾
    work = tempfile.mkdtemp(prefix="work_")

    # 儲存 base.xlsx
    base_path = os.path.join(work, "base.xlsx")
    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)

    # 儲存 zip
    zip_path = os.path.join(work, "data.zip")
    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # 複製 base 成 result（保留所有格式）
    result_path = os.path.join(work, "result.xlsx")
    shutil.copy(base_path, result_path)

    # 開啟 result.xlsx
    wb = load_workbook(result_path)
    ws_result = wb.active

    # 解壓 ZIP
    unzip_dir = os.path.join(work, "unzipped")
    with zipfile.ZipFile(zip_path, 'r') as z:
        z.extractall(unzip_dir)

    # 讀取 ZIP 內所有 Excel
    excel_files = glob.glob(os.path.join(unzip_dir, "*.xls*"))

    for excel in excel_files:
        # 每個 Excel 可能有多個 sheet
        xls = pd.ExcelFile(excel)

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            process_and_paste(df, ws_result)

    # 儲存 result
    wb.save(result_path)

    return FileResponse(
        result_path,
        filename="result.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )