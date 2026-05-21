from fastapi import FastAPI, UploadFile, File
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

import os
import shutil
import zipfile
import tempfile
import glob

from openpyxl import Workbook, load_workbook

app = FastAPI()

app.mount("/static", StaticFiles(directory="static", html=True), name="static")


@app.get("/")
def home():
    return FileResponse("static/index.html")


@app.post("/process")
async def process(data_zip: UploadFile = File(...)):

    # =========================
    # 工作目錄
    # =========================
    work = tempfile.mkdtemp(prefix="work_")

    # ZIP
    zip_path = os.path.join(work, "data.zip")

    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # 解壓
    unzip_dir = os.path.join(work, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(unzip_dir)

    # 找所有 Excel
    files = glob.glob(
        os.path.join(unzip_dir, "**", "*.xls*"),
        recursive=True
    )

    if not files:
        return {"error": "ZIP 沒有 Excel"}

    # =========================
    # 建立輸出 Workbook
    # =========================
    out_wb = Workbook()

    # 刪除預設 sheet
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    # 紀錄每個 sheet 已寫到哪
    sheet_row_map = {}

    # =========================
    # 處理每個 Excel
    # =========================
    for excel_file in files:

        try:
            wb = load_workbook(excel_file, data_only=True)

            for sheet_name in wb.sheetnames:

                src_ws = wb[sheet_name]

                # 目的 sheet
                if sheet_name not in out_wb.sheetnames:
                    out_ws = out_wb.create_sheet(sheet_name)
                    sheet_row_map[sheet_name] = {}
                else:
                    out_ws = out_wb[sheet_name]

                # =========================
                # A3, A8, A13...
                # =========================
                for start_row in range(3, 29, 5):

                    key = src_ws[f"A{start_row}"].value

                    if key is None:
                        continue

                    key = str(key).strip()

                    # ==================================
                    # 找此 key 要貼到哪裡
                    # ==================================
                    if key not in sheet_row_map[sheet_name]:

                        # 第一次出現
                        if len(sheet_row_map[sheet_name]) == 0:
                            target_row = 3
                        else:
                            target_row = max(
                                sheet_row_map[sheet_name].values()
                            ) + 5

                        sheet_row_map[sheet_name][key] = target_row

                        # 寫 A 欄 key
                        out_ws[f"A{target_row}"] = key

                    else:
                        target_row = sheet_row_map[sheet_name][key]

                    # ==================================
                    # 複製 B~AJ 共5列
                    # ==================================
                    for r_offset in range(5):

                        src_r = start_row + r_offset
                        dst_r = target_row + r_offset

                        # B=2 ~ AJ=36
                        for col in range(2, 37):

                            value = src_ws.cell(
                                row=src_r,
                                column=col
                            ).value

                            out_ws.cell(
                                row=dst_r,
                                column=col
                            ).value = value

        except Exception as e:
            print("錯誤:", excel_file, e)

    # =========================
    # 儲存
    # =========================
    result_path = os.path.join(work, "result.xlsx")

    out_wb.save(result_path)

    # =========================
    # 回傳
    # =========================
    return FileResponse(
        result_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="result.xlsx"
    )