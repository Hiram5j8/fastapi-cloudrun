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

    work = tempfile.mkdtemp(prefix="work_")

    # === 儲存 base.xlsx ===
    base_path = os.path.join(work, "base.xlsx")
    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)

    # === 複製成 result.xlsx（最後回傳這個） ===
    result_path = os.path.join(work, "result.xlsx")
    shutil.copy(base_path, result_path)

    # === 儲存 ZIP ===
    zip_path = os.path.join(work, "data.zip")
    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # === 解壓 ZIP ===
    unzip_dir = os.path.join(work, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, 'r') as z:
        z.extractall(unzip_dir)

    files = glob.glob(os.path.join(unzip_dir, "**", "*.xls*"), recursive=True)

    # === 開啟 result.xlsx（openpyxl）===
    wb = load_workbook(result_path)

    LEAVE = {"休", "請假", "休假", "特休", "請假一天"}

    for file in files:

        try:
            data = pd.read_excel(file, sheet_name=None)
        except:
            continue

        for sh, df in data.items():

            if sh not in wb.sheetnames:
                continue

            ws = wb[sh]

            if df.shape[1] < 2:
                continue

            # 範例：從第3列開始，每5列一組
            for start in range(3, 153, 5):

                r0 = start - 2
                if r0 >= len(df):
                    break

                src = str(df.iloc[r0, 0]).strip()
                if not src:
                    continue

                # 找空列（A欄）
                found = None
                for r in range(3, 153, 5):
                    if not ws.cell(r, 1).value:
                        ws.cell(r, 1).value = src
                        found = r
                        break

                if not found:
                    continue

                # ===== Bulk block（B~AK 共36欄）=====
                block = [[None]*36 for _ in range(5)]

                for i in range(5):

                    r = start + i - 2

                    # B欄
                    val = df.iloc[r, 1] if r < len(df) else None
                    if pd.notna(val):
                        block[i][0] = str(val)

                    # C~AK
                    for c in range(2, 37):
                        val = df.iloc[r, c] if (r < len(df) and c < df.shape[1]) else None

                        if pd.isna(val):
                            continue

                        if isinstance(val, (int, float, np.number)):
                            block[i][c-1] = val
                        elif isinstance(val, str) and val.strip() in LEAVE:
                            block[i][c-1] = "請假"

                # 一次貼上
                for i in range(5):
                    for j in range(36):
                        ws.cell(found + i, 2 + j).value = block[i][j]

    wb.save(result_path)

    return FileResponse(
        result_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="result.xlsx"
    )