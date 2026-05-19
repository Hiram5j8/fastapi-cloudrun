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
@app.post("/process")
async def process(base_xls: UploadFile = File(...), data_zip: UploadFile = File(...)):
    work = "/tmp/work"
    unzip_dir = os.path.join(work, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    # === 儲存 base.xlsx ===
    base_path = os.path.join(work, "base.xlsx")
    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)

    # === 複製成 result.xlsx ===
    result_path = os.path.join(work, "result.xlsx")
    shutil.copy(base_path, result_path)
    """
    # === 儲存並解壓 zip ===
    zip_path = os.path.join(work, "data.zip")
    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(unzip_dir)

    # === 開啟 result.xlsx ===
    result_wb = load_workbook(result_path)
    result_ws = result_wb.active

    # 從 result 最後一列開始往下貼
    paste_start_row = result_ws.max_row + 2

    # === 走訪 zip 內 Excel ===
    for root, dirs, files in os.walk(unzip_dir):
        for file in files:
            if file.endswith(".xlsx"):
                file_path = os.path.join(root, file)

                src_wb = load_workbook(file_path, data_only=True)
                src_ws = src_wb.active

                # 複製 C1:AK36
                for r in range(1, 37):           # 1~36
                    for c in range(3, 38):       # C(3) ~ AK(37)
                        value = src_ws.cell(row=r, column=c).value
                        result_ws.cell(
                            row=paste_start_row + (r - 1),
                            column=c - 2,        # 貼到 A 開始
                            value=value
                        )

                paste_start_row += 40  # 每個檔案間隔

    # === 儲存 ===
    result_wb.save(result_path)
    """
    return FileResponse(
        result_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="output.xlsx"
    )