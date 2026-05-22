from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles

from openpyxl import load_workbook

from dataclasses import dataclass
from typing import Any, List

from io import BytesIO
import zipfile

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
# RowSeries
# =========================
class RowSeries:
    def __init__(self, sheet, start_row, start_col, values):
        self.sheet = sheet
        self.start_row = start_row
        self.start_col = start_col
        self.values = values

    def write(self):
        """
        只寫 value
        不覆蓋 style
        """
        for col_offset, value in enumerate(self.values):
            cell = self.sheet.cell(
                row=self.start_row,
                column=self.start_col + col_offset
            )

            cell.value = value


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

        # 有差異才記錄
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

    # zip
    zip_path = os.path.join(work, data_zip.filename)

    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # 解壓縮
    unzip_dir = os.path.join(work, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(unzip_dir)

    # 開 base workbook
    result_wb = load_workbook(base_path)

    # 掃描 ZIP Excel
    for file in os.listdir(unzip_dir):

        if not file.endswith((".xlsx", ".xlsm")):
            continue

        src_path = os.path.join(unzip_dir, file)

        print("處理:", file)

        src_wb = load_workbook(src_path, data_only=False)

        # sheet 比對
        for sheet_name in src_wb.sheetnames:

            if sheet_name not in result_wb.sheetnames:
                continue

            base_ws = result_wb[sheet_name]
            src_ws = src_wb[sheet_name]

            # 差異分析
            diff_rows = compare_sheet(base_ws, src_ws)

            # 寫回
            for row_num, values in diff_rows:

                rs = RowSeries(
                    sheet=base_ws,
                    start_row=row_num,
                    start_col=1,
                    values=values
                )

                rs.write()

    # 輸出
    result_path = os.path.join(work, "result.xlsx")

    result_wb.save(result_path)

    return FileResponse(
        result_path,
        filename="result.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )