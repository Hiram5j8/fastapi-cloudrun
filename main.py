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
async def process(
    base_xls: UploadFile = File(...),
    data_zip: UploadFile = File(...)
):
    work = "/tmp/work"
    unzip_dir = os.path.join(work, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    # === 儲存上傳檔案 ===
    base_path = os.path.join(work, base_xls.filename)
    zip_path = os.path.join(work, data_zip.filename)

    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)

    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # === 解壓縮 zip ===
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(unzip_dir)

    # === 建立結果 Excel ===
    # 上傳檔案存起來
    upload_path = os.path.join(work, base_xls.filename)
  
    # 直接複製成 result.xlsx
    result_path = os.path.join(work, "result.xlsx")
    shutil.copy(upload_path, result_path)
    
    result_wb = Workbook()
    result_ws = result_wb.active
    result_ws.title = "Merged Data"

    write_row = 1

    # === 走訪 zip 內所有 Excel 檔 ===
    for root, dirs, files in os.walk(unzip_dir):
        for file in files:
            if file.endswith(".xlsx"):
                file_path = os.path.join(root, file)
                src_wb = load_workbook(file_path, data_only=True)
                src_ws = src_wb.active

                # 檔名當分隔標題
                result_ws.cell(row=write_row, column=1, value=f"=== {file} ===")
                write_row += 1

                # 將每個儲存格的值寫入 result
                for row in src_ws.iter_rows(values_only=True):
                    for col_idx, cell_value in enumerate(row, start=1):
                        result_ws.cell(row=write_row, column=col_idx, value=cell_value)
                    write_row += 1

                write_row += 2  # 空兩行分隔

    # === 儲存結果 ===
    result_path = os.path.join(work, "result.xlsx")
    result_wb.save(result_path)

    # === 回傳下載 ===
    return FileResponse(
        result_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="result.xlsx"
    )