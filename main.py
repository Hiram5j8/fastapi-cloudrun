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


@app.post("/process")
async def process(base_xls: UploadFile = File(...), data_zip: UploadFile = File(...)):
    import logging

    work = "/tmp/work"
    os.makedirs(work, exist_ok=True)

    base_path = os.path.join(work, base_xls.filename)
    zip_path = os.path.join(work, data_zip.filename)

    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)

    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    base_df = pd.read_excel(base_path)

    # 🔥 防呆
    if base_df.shape[0] < 3 or base_df.shape[1] < 1:
        return {"error": "base.xlsx 格式不正確"}

    base_key = str(base_df.iloc[2, 0])

    unzip_dir = os.path.join(work, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, 'r') as z:
        z.extractall(unzip_dir)

    result_wb = Workbook()
    result_ws = result_wb.active

    hit_count = 0

    for root, _, files in os.walk(unzip_dir):
        for file in files:
            if file.endswith((".xls", ".xlsx")):
                fpath = os.path.join(root, file)

                try:
                    df = pd.read_excel(fpath)
                except Exception as e:
                    continue

                if df.shape[0] < 3 or df.shape[1] < 1:
                    continue

                key = str(df.iloc[2, 0])

                if key == base_key:
                    hit_count += 1

                    for c in range(2, min(37, df.shape[1]), 5):
                        row = df.iloc[2, c:c+5].tolist()
                        result_ws.append(row)

    # 🔥 如果完全沒資料 → 寫提示
    if hit_count == 0:
        result_ws.append(["NO MATCH FOUND"])

    result_path = os.path.join(work, "result.xlsx")
    result_wb.save(result_path)

    return FileResponse(
        result_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="result.xlsx"
    )