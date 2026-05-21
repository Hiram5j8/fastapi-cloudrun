from fastapi import FastAPI, UploadFile, File
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import os
import shutil
import zipfile
import tempfile
import glob

import pandas as pd
from openpyxl import Workbook

app = FastAPI()

# static
app.mount("/static", StaticFiles(directory="static", html=True), name="static")


# 首頁
@app.get("/")
def home():
    return FileResponse("static/index.html")


# API
@app.post("/process")
async def process(
    base_xls: UploadFile = File(...),
    data_zip: UploadFile = File(...)
):
    # =========================
    # 建立工作目錄
    # =========================
    work = tempfile.mkdtemp(prefix="work_")

    # =========================
    # 儲存 ZIP
    # =========================
    zip_path = os.path.join(work, "data.zip")

    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # =========================
    # 解壓 ZIP
    # =========================
    unzip_dir = os.path.join(work, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(unzip_dir)

    # =========================
    # 找全部 Excel
    # =========================
    files = glob.glob(
        os.path.join(unzip_dir, "**", "*.xls*"),
        recursive=True
    )

    if len(files) == 0:
        return {"error": "ZIP 內沒有 Excel"}

    # =========================
    # 建立輸出 Excel
    # =========================
    result_path = os.path.join(work, "result.xlsx")

    writer = pd.ExcelWriter(
        result_path,
        engine="openpyxl"
    )

    sheet_index = 1

    # =========================
    # 逐一處理 Excel
    # =========================
    for excel_file in files:

        try:
            # 讀取所有 sheet
            xls = pd.ExcelFile(excel_file)

            for sheet_name in xls.sheet_names:

                # 讀 sheet
                df = pd.read_excel(
                    excel_file,
                    sheet_name=sheet_name
                )

                # sheet 名稱避免重複
                new_sheet_name = f"S{sheet_index}_{sheet_name}"

                # Excel sheet 最長 31 字
                new_sheet_name = new_sheet_name[:31]

                # 寫入
                df.to_excel(
                    writer,
                    sheet_name=new_sheet_name,
                    index=False
                )

                sheet_index += 1

        except Exception as e:
            print("錯誤:", excel_file, e)

    # =========================
    # 儲存
    # =========================
    writer.close()

    # =========================
    # 回傳
    # =========================
    return FileResponse(
        result_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="result.xlsx"
    )