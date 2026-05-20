from fastapi import FastAPI, UploadFile, File
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import zipfile, shutil, os
import pandas as pd
from openpyxl import Workbook

app = FastAPI()

# ✅ static 一定只能掛 /static（安全）
app.mount("/static", StaticFiles(directory="static", html=True), name="static")


# ✅ 首頁
@app.get("/")
def home():
    return FileResponse("static/index.html")


# ✅ API
import os, shutil, zipfile, tempfile, glob
from datetime import datetime

import pandas as pd
import numpy as np
from fastapi import UploadFile, File
from fastapi.responses import FileResponse
from openpyxl import load_workbook

@app.post("/process")
async def process(base_xls: UploadFile = File(...),
                  data_zip: UploadFile = File(...)):
    # 建立工作目錄
    work = tempfile.mkdtemp(prefix="work_")

    # === 儲存 base.xlsx ===
    base_path = os.path.join(work, "base.xlsx")
    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)

    # === 複製成 result.xlsx（最後回傳這個） ===
    #result_path = os.path.join(work, "result.xlsx")
    #shutil.copy(base_path, result_path)
   

    # === 儲存 ZIP ===
    zip_path = os.path.join(work, "data.zip")
    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # === 解壓 ZIP ===
    unzip_dir = os.path.join(work, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, 'r') as z:
        z.extractall(unzip_dir)

    # === 找到 ZIP 內唯一的 Excel ===
    files = glob.glob(os.path.join(unzip_dir, "**", "*.xls*"), recursive=True)

    if len(files) != 1:
        return {"error": "ZIP 內必須且只能有一個 Excel 檔"}

    src_excel = files[0]

    # === 讀取後另存成 result.xlsx ===
    df = pd.read_excel(src_excel)

    result_path = os.path.join(work, "result.xlsx")
    df.to_excel(result_path, index=False)

    # === 回傳下載 ===
    return FileResponse(
        result_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="result.xlsx"
    )