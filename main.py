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
    df = df.copy()

    # 範例：新增一欄 sheet 名稱
    #df["sheet_name"] = sheet_name

    return df

def match(index, v):
    return index.get(key(v))

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

    base_path = os.path.join(work, "base.xlsx")
    zip_path = os.path.join(work, "data.zip")

    # 2. 存 base.xlsx
    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)

    # 3. 存 zip
    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # 4. 解壓 zip
    extract_dir = os.path.join(work, "extract")
    os.makedirs(extract_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(extract_dir)

    # 5. 找 Excel 檔
    excel_file = None
    for root, _, files in os.walk(extract_dir):
        for f in files:
            if f.endswith((".xls", ".xlsx")):
                excel_file = os.path.join(root, f)
                break

    if not excel_file:
        return {"error": "ZIP 內沒有 Excel 檔"}

    # 6. 讀 Excel（所有 sheet）
    xls = pd.ExcelFile(excel_file)

    result_sheets = {}

    # 7. 依序處理 sheet
    for i, sheet_name in enumerate(xls.sheet_names):
        df = pd.read_excel(xls, sheet_name=sheet_name)

        print(f"處理第 {i+1} 張 sheet: {sheet_name}")

        #result_df = process_sheet(df, sheet_name)
        max_c = 37  # AK

        # ZIP 每5列一組：A3, A8, A13...
        for start in range(3, 68, 5):

            zip_r0 = start - 2
            if zip_r0 >= len(df):
                break

            zip_key = norm(df.iloc[zip_r0, 0])  # ZIP 的 A欄
            if zip_key == "":
                continue

            # 去比對 result 同一列的 A欄
            result_val = norm(ws_result.cell(row=start, column=1).value)

            if zip_key != result_val:
                continue

            # 相同 ⇒ 貼 B~AK，5列原位貼上
            for i in range(5):  # B3~B7 / B8~B12 ...
                zip_r = start + i - 2

                for c in range(1, max_c):  # B~AK
                    val = safe(df, zip_r, c)

                    ws_result.cell(
                        row=start + i,
                        column=2 + (c - 1),
                        value=val
                    )

        

        result_sheets[sheet_name] = result_df

    # 8. 輸出 result.xlsx
    result_path = os.path.join(work, "result.xlsx")

    with pd.ExcelWriter(result_path, engine="openpyxl") as writer:
        for sheet_name, df in result_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    # =========================
    # 7. 回傳結果檔
    # =========================
    return FileResponse(
        result_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="result.xlsx"
    )

