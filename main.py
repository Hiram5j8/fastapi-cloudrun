import os
import shutil
import zipfile
import tempfile
import pandas as pd

from fastapi import FastAPI, UploadFile, File
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from openpyxl import load_workbook

app = FastAPI()

# static 首頁
app.mount("/static", StaticFiles(directory="static", html=True), name="static")

@app.get("/")
def home():
    return FileResponse("static/index.html")


# =========================
# 你的資料處理邏輯（可自行修改）
# =========================
def process_sheet(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    df = df.copy()
    return df


# =========================
# 核心：貼到 A3 ~ AK38（不破壞格式）
# =========================
def paste_df_to_A3_AK38(ws, df):
    start_row = 3   # A3
    start_col = 1   # A
    max_rows = 36   # 3~38
    max_cols = 37   # A~AK

    rows = min(df.shape[0], max_rows)
    cols = min(df.shape[1], max_cols)

    for r in range(rows):
        for c in range(cols):
            ws.cell(
                row=start_row + r,
                column=start_col + c
            ).value = df.iat[r, c]


# =========================
# API
# =========================
@app.post("/process")
async def process(
    base_xls: UploadFile = File(...),
    data_zip: UploadFile = File(...)
):
    work = tempfile.mkdtemp(prefix="work_")

    base_path = os.path.join(work, "base.xlsx")
    zip_path = os.path.join(work, "data.zip")

    # 儲存上傳檔案
    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)

    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # 載入範本（保留所有格式）
    wb = load_workbook(base_path)

    # 解壓 ZIP
    extract_dir = os.path.join(work, "extract")
    os.makedirs(extract_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(extract_dir)

    # 找 ZIP 裡所有 Excel
    excel_files = []
    for root, _, files in os.walk(extract_dir):
        for f in files:
            if f.endswith((".xls", ".xlsx")):
                excel_files.append(os.path.join(root, f))

    if not excel_files:
        return {"error": "ZIP 內沒有 Excel 檔"}

    # === 逐一處理每個 Excel、每個 Sheet ===
    for excel_path in excel_files:
        xls = pd.ExcelFile(excel_path)

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            result_df = process_sheet(df, sheet_name)

            # 範本沒有這個 sheet 就跳過
            if sheet_name not in wb.sheetnames:
                continue

            ws = wb[sheet_name]

            # ⭐貼到 A3
            paste_df_to_A3_AK38(ws, result_df)

    # 輸出結果
    result_path = os.path.join(work, "result.xlsx")
    wb.save(result_path)

    return FileResponse(
        result_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="result.xlsx"
    )