from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles

from openpyxl import load_workbook, Workbook

from dataclasses import dataclass
from typing import Any, List

from io import BytesIO
import zipfile
import os
import shutil
import tempfile
import re

# =====================================================
# FastAPI
# =====================================================

app = FastAPI()

# ✅ static
app.mount(
    "/static",
    StaticFiles(directory="static", html=True),
    name="static"
)

# ✅ 首頁
@app.get("/")
def home():
    return FileResponse("static/index.html")

# =========================
# 版本 API
# =========================
VERSION = "2026-05-22 1339"

@app.get("/version")
def version():
    return {
        "version": VERSION
    }

# =====================================================
# RowSeries
# =====================================================

class RowSeries:

    def __init__(
        self,
        sheet,
        start_row,
        start_col,
        values
    ):

        self.sheet = sheet
        self.start_row = start_row
        self.start_col = start_col
        self.values = values

    def write_txt(self, fp):

        row_text = []

        for col_offset, value in enumerate(self.values):

            # 空值跳過
            if value is None:
                continue

            value_str = str(value).strip()

            # 空白跳過
            if value_str == "":
                continue

            col_num = self.start_col + col_offset

            cell_ref = f"R{self.start_row}C{col_num}"

            row_text.append(
                f"{cell_ref}={value_str}"
            )

        # 整列都空 -> 不輸出
        if not row_text:
            return

        fp.write(
            ",".join(row_text) + "\n"
        )

# =========================
# 比較差異
# =========================
def compare_sheet(base_ws, src_ws):

    diff_rows = []

    max_row = min(base_ws.max_row, src_ws.max_row)
    max_col = min(base_ws.max_column, src_ws.max_column)

    for r in range(1, max_row + 1):

        row_values = []
        has_diff = False

        for c in range(1, max_col + 1):

            base_val = base_ws.cell(r, c).value
            src_val = src_ws.cell(r, c).value

            # 差異判斷
            if base_val != src_val:
                has_diff = True

            row_values.append(src_val)

        # 有差異才輸出
        if has_diff:
            diff_rows.append((r, row_values))

    return diff_rows


# =========================
# API
# =========================
@app.post("/process")
async def process(
    base_xls: UploadFile = File(...),
    data_zip: UploadFile = File(...)
):

    # 工作目錄
    work = tempfile.mkdtemp(prefix="excel_")

    # base
    base_path = os.path.join(work, base_xls.filename)

    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)
    # 直接複製成 result.xlsx
    result_path = os.path.join(work, "result.xlsx")
    shutil.copy(upload_path, result_path)
    
    # zip
    zip_path = os.path.join(work, data_zip.filename)

    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # 解壓縮
    unzip_dir = os.path.join(work, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(unzip_dir)

    # 開啟 base
    base_wb = load_workbook(base_path)
    
    #change xlsx
    result_wb = Workbook()
    result_ws = result_wb.active
    result_ws.title = "RESULT"
    
    LEAVE = {"休", "請假", "休假", "特休", "請假一天"}
    
    # TXT 結果
    txt_path = os.path.join(work, "Result.txt")

    with open(txt_path, "w", encoding="utf-8") as txt_fp:

        # 掃描 ZIP
        for file in os.listdir(unzip_dir):

            if not file.endswith((".xlsx", ".xlsm")):
                continue

            src_path = os.path.join(unzip_dir, file)

            print("處理:", file)
            """
            txt_fp.write("=" * 60 + "\n")
            txt_fp.write(f"FILE: {file}\n")
            txt_fp.write("=" * 60 + "\n")
            """
            src_wb = load_workbook(src_path, data_only=False)

            # sheet 比較
            for sheet_name in src_wb.sheetnames:

                if sheet_name not in base_wb.sheetnames:
                    continue

                base_ws = base_wb[sheet_name]
                src_ws = src_wb[sheet_name]

                txt_fp.write(f"\n[SHEET] {sheet_name}\n")

                # 差異
                diff_rows = compare_sheet(base_ws, src_ws)
                
                for row_num, values in diff_rows:

                    for col_offset, value in enumerate(values):

                        if value is None:
                            continue

                        value_str = str(value).strip()

                        if value_str == "":
                            continue

                        col_num = 1 + col_offset

                        result_ws.cell(
                            row=row_num,
                            column=col_num,
                            value=value_str
                        )
                
                """
                # 輸出 TXT
                for row_num, values in diff_rows:

                    rs = RowSeries(
                        sheet=src_ws,
                        start_row=row_num,
                        start_col=1,
                        values=values
                    )

                    rs.write_txt(txt_fp)
                """                    
    
    # 儲存
    result_path = os.path.join(work, "Result.xlsx")

    result_wb.save(result_path)

    return FileResponse(
        path=result_path,
        filename="Result.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

"""
    return FileResponse(
        txt_path,
        filename="Result.txt",
        media_type="text/plain"
    )
"""