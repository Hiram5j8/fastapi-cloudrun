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

@app.post("/process")
async def process(
    base_xls: UploadFile = File(...),
    data_zip: UploadFile = File(...)
):
    # === 建立工作目錄 ===
    work = tempfile.mkdtemp(prefix="work_")

    # === 儲存 base.xlsx ===
    base_path = os.path.join(work, "base.xlsx")
    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)

    # === 複製成 result.xlsx ===
    result_path = os.path.join(work, "result.xlsx")
    shutil.copyfile(base_path, result_path)

    # === 儲存並解壓 ZIP ===
    zip_path = os.path.join(work, "data.zip")
    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    unzip_dir = os.path.join(work, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, 'r') as z:
        z.extractall(unzip_dir)

    # === 開啟 result.xlsx ===
    dst_wb = load_workbook(result_path)
    dst_ws = dst_wb.active


    engine, master_wb = open_excel(master_file)

    sheets = (
        master_wb.sheetnames
        if engine == "openpyxl"
        else master_wb.sheet_names
    )
    
    paste_row = dst_ws.max_row + 2

    # === 找 zip 內所有 Excel ===
    files = glob.glob(os.path.join(unzip_dir, "**", "*.xlsx"), recursive=True)

    for file in files:
        src_wb = load_workbook(file, data_only=True)
        src_ws = src_wb.active
        
        for sh in sheets:
            # ===== 複製 C1:AK36 =====
                for r in range(1, 37):          # 1~36
                for c in range(3, 38):      # C(3) ~ AK(37)
                    val = src_ws.cell(row=r, column=c).value

                    dst_ws.cell(
                        row=paste_row + (r - 1),
                        column=c,          # 貼回同樣 C~AK
                        value=val
                    )

        #paste_row += 40  # 每個檔案間隔
    # === 儲存 ===
    dst_wb.save(result_path)

    return FileResponse(
        result_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="result.xlsx"
    )