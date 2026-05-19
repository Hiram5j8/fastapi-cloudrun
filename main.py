from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
import zipfile
import shutil
import os
from openpyxl import load_workbook, Workbook

app = FastAPI()

# 掛載前端網頁
app.mount("/", StaticFiles(directory="static", html=True), name="static")


@app.post("/process")
async def process(base_xls: UploadFile = File(...), data_zip: UploadFile = File(...)):
    work_dir = "/tmp/work"
    os.makedirs(work_dir, exist_ok=True)

    base_path = os.path.join(work_dir, base_xls.filename)
    zip_path = os.path.join(work_dir, data_zip.filename)

    # 儲存上傳檔
    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)

    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # 解壓縮 ZIP
    unzip_dir = os.path.join(work_dir, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(unzip_dir)

    # 讀 base.xls
    base_wb = load_workbook(base_path)
    base_ws = base_wb.active

    result_wb = Workbook()
    result_ws = result_wb.active
    result_row = 1

    # 規則：C3~AK3 每5格一組，比對 A3
    base_key = base_ws["A3"].value

    for root, dirs, files in os.walk(unzip_dir):
        for file in files:
            if file.endswith(".xls") or file.endswith(".xlsx"):
                file_path = os.path.join(root, file)
                wb = load_workbook(file_path)
                ws = wb.active

                if ws["A3"].value == base_key:
                    cols = range(3, 38, 5)  # C(3) ~ AK(37)

                    for c in cols:
                        values = [ws.cell(row=3, column=c+i).value for i in range(5)]
                        result_ws.append(values)
                        result_row += 1

    result_path = os.path.join(work_dir, "result.xlsx")
    result_wb.save(result_path)

    return FileResponse(result_path, filename="result.xlsx")
