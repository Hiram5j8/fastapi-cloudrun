import os
import re
import math
import glob
import hashlib
import zipfile
import tempfile
import unicodedata
import numpy as np
import pandas as pd

from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from openpyxl import load_workbook

app = FastAPI()
# ✅ static 一定只能掛 /static（安全）
app.mount("/static", StaticFiles(directory="static", html=True), name="static")


# ✅ 首頁
@app.get("/")
def home():
    return FileResponse("static/index.html")


# ===============================
# 正規化（防假空白）
# ===============================
def norm(v):
    if v is None:
        return ""

    if isinstance(v, float) and math.isnan(v):
        return ""

    if isinstance(v, float) and v.is_integer():
        v = int(v)

    v = str(v)
    v = unicodedata.normalize("NFKC", v)

    for ch in ["\xa0","\u3000","\ufeff","\u200b","\u200c","\u200d","\u00ad"]:
        v = v.replace(ch, "")

    v = re.sub(r"[\x00-\x1F\x7F]", "", v)
    v = re.sub(r"\s+", "", v)

    return v.lower()


def key(v):
    return hashlib.md5(norm(v).encode()).hexdigest()


# ===============================
# 建立 master index
# ===============================
def build_index(ws):
    idx = {}
    for r in range(3, 153, 5):
        v = ws.cell(r, 1).value
        if norm(v) != "":
            idx[key(v)] = r
    return idx


# ===============================
# API
# ===============================
@app.post("/process")
async def process(master_xls: UploadFile = File(...),
                  data_zip: UploadFile = File(...)):

    work = tempfile.mkdtemp()
    master_path = os.path.join(work, "master.xlsx")
    zip_path = os.path.join(work, "data.zip")

    # 儲存上傳檔
    with open(master_path, "wb") as f:
        f.write(await master_xls.read())

    with open(zip_path, "wb") as f:
        f.write(await data_zip.read())

    # 解壓 ZIP
    unzip_dir = os.path.join(work, "unzipped")
    with zipfile.ZipFile(zip_path, 'r') as z:
        z.extractall(unzip_dir)

    files = glob.glob(os.path.join(unzip_dir, "**", "*.xls*"), recursive=True)

    # 複製 master 為 result
    result_path = os.path.join(work, "result.xlsx")
    wb = load_workbook(master_path)
    wb.save(result_path)
    wb = load_workbook(result_path)

    LEAVE = {"休", "請假", "休假", "特休", "請假一天"}

    for file in files:
        try:
            data = pd.read_excel(file, sheet_name=None)
        except:
            continue

        for sh in wb.sheetnames:

            if sh not in data:
                continue

            ws = wb[sh]
            df = data[sh]

            if df.shape[1] < 2:
                continue

            index = build_index(ws)

            for start in range(3, 153, 5):

                r0 = start - 2
                if r0 >= len(df):
                    break

                src = norm(df.iloc[r0, 0])
                if src == "":
                    continue

                found = index.get(key(src))

                # 不存在 → 新增
                if found is None:
                    for r in range(3, 153, 5):
                        if norm(ws.cell(r, 1).value) == "":
                            found = r
                            ws.cell(r, 1).value = src
                            index[key(src)] = r
                            break

                if found is None:
                    continue

                # B 欄
                for i in range(5):
                    r = start + i - 2
                    if r < len(df):
                        val = df.iloc[r, 1]
                        if pd.notna(val):
                            ws.cell(found + i, 2).value = str(val)

                # C ~ AK (第37欄)
                for i in range(5):
                    r = start + i - 2
                    for c in range(2, 37):
                        if r < len(df) and c < df.shape[1]:
                            val = df.iloc[r, c]

                            if pd.isna(val):
                                continue

                            if isinstance(val, (int, float, np.number)):
                                ws.cell(found + i, c + 1).value = val

                            elif isinstance(val, str) and val.strip() in LEAVE:
                                ws.cell(found + i, c + 1).value = "請假"

    wb.save(result_path)

    return FileResponse(result_path, filename="result.xlsx")