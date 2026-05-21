# =========================
# FastAPI 主程式
# =========================

# FastAPI 主框架
from fastapi import FastAPI, UploadFile, File

# 提供靜態網頁（HTML/CSS/JS）
from fastapi.staticfiles import StaticFiles

# 回傳檔案下載
from fastapi.responses import FileResponse

# ZIP 解壓縮
import zipfile

# 檔案複製
import shutil

# 作業系統檔案操作
import os

# pandas 處理 Excel
import pandas as pd

# openpyxl 建立/讀取 Excel
from openpyxl import Workbook
from openpyxl import load_workbook

# 建立暫存資料夾
import tempfile

# 搜尋檔案
import glob

# 日期時間
from datetime import datetime

# numpy 數值運算
import numpy as np


# =========================
# 建立 FastAPI App
# =========================
app = FastAPI()


# =========================
# 掛載 static 資料夾
# =========================
# 讓瀏覽器可以讀取 static 裡面的檔案
# 例如:
# static/index.html
# static/style.css
#
# 網址:
# http://127.0.0.1:8000/static/xxx
# =========================
app.mount("/static", StaticFiles(directory="static", html=True), name="static")


# =========================
# 首頁 API
# =========================
# 開啟網站首頁時
# 直接回傳 static/index.html
# =========================
@app.get("/")
def home():

    # 回傳 HTML 首頁
    return FileResponse("static/index.html")


# =====================================================
# 找到真正有資料的 sheet
# =====================================================
def find_data_sheet(wb):

    # wb.sheetnames = 所有工作表名稱
    for name in wb.sheetnames:

        # 取得工作表
        ws = wb[name]

        # -------------------------------------------------
        # 判斷 A3 是否有資料
        # row=3
        # column=1 = A欄
        # -------------------------------------------------
        if ws.cell(row=3, column=1).value is not None:

            # 找到有資料的 sheet
            return ws

    # -------------------------------------------------
    # 如果全部都沒資料
    # 回傳目前作用中的 sheet
    # -------------------------------------------------
    return wb.active


# =====================================================
# 每個 sheet 的處理邏輯
# =====================================================
def process_sheet(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    """
    df = 目前 sheet 的資料
    sheet_name = 工作表名稱
    """

    # 複製 DataFrame
    # 避免修改原始資料
    df = df.copy()

    # =================================================
    # 這裡可以加入你的資料處理邏輯
    # =================================================

    # 範例:
    # 新增一欄 sheet 名稱
    # df["sheet_name"] = sheet_name

    # 回傳處理後資料
    return df


# =====================================================
# API : /process
# =====================================================
# 上傳:
# 1. base_xls
# 2. data_zip
#
# 然後:
# 1. 解壓 ZIP
# 2. 找 Excel
# 3. 讀所有 sheet
# 4. 處理資料
# 5. 輸出 result.xlsx
# =====================================================
@app.post("/process")
async def process(

    # 上傳的 base Excel
    base_xls: UploadFile = File(...),

    # 上傳的 ZIP
    data_zip: UploadFile = File(...)
):

    # =================================================
    # 1. 建立工作資料夾
    # =================================================
    # tempfile.mkdtemp()
    # 會建立:
    # /tmp/work_xxxxx
    # =================================================
    work = tempfile.mkdtemp(prefix="work_")


    # =================================================
    # 定義檔案路徑
    # =================================================
    base_path = os.path.join(work, "base.xlsx")

    zip_path = os.path.join(work, "data.zip")


    # =================================================
    # 2. 儲存 base.xlsx
    # =================================================
    with open(base_path, "wb") as f:

        # 把上傳檔案複製到硬碟
        shutil.copyfileobj(base_xls.file, f)


    # =================================================
    # 3. 儲存 ZIP
    # =================================================
    with open(zip_path, "wb") as f:

        # 把 ZIP 複製到硬碟
        shutil.copyfileobj(data_zip.file, f)


    # =================================================
    # 4. 解壓 ZIP
    # =================================================

    # 解壓目錄
    extract_dir = os.path.join(work, "extract")

    # 建立資料夾
    os.makedirs(extract_dir, exist_ok=True)


    # 開啟 ZIP
    with zipfile.ZipFile(zip_path, "r") as z:

        # 解壓全部檔案
        z.extractall(extract_dir)


    # =================================================
    # 5. 尋找 ZIP 內的 Excel
    # =================================================
    excel_file = None

    # os.walk()
    # 遞迴搜尋所有資料夾
    for root, _, files in os.walk(extract_dir):

        for f in files:

            # 找 .xls 或 .xlsx
            if f.endswith((".xls", ".xlsx")):

                # 組合完整路徑
                excel_file = os.path.join(root, f)

                # 找到就停止
                break


    # =================================================
    # 如果 ZIP 內沒有 Excel
    # =================================================
    if not excel_file:

        return {
            "error": "ZIP 內沒有 Excel 檔"
        }


    # =================================================
    # 6. 讀取 Excel
    # =================================================
    # pd.ExcelFile()
    # 可以一次讀所有 sheet
    # =================================================
    xls = pd.ExcelFile(excel_file)


    # =================================================
    # 儲存處理結果
    # =================================================
    result_sheets = {}


    # =================================================
    # 7. 依序處理每個 sheet
    # =================================================
    for i, sheet_name in enumerate(xls.sheet_names):

        # ---------------------------------------------
        # 讀取目前 sheet
        # ---------------------------------------------
        df = pd.read_excel(
            xls,
            sheet_name=sheet_name
        )

        # 顯示目前進度
        print(f"處理第 {i+1} 張 sheet: {sheet_name}")


        # ---------------------------------------------
        # 呼叫你的處理函式
        # ---------------------------------------------
        result_df = process_sheet(
            df,
            sheet_name
        )


        # ---------------------------------------------
        # 儲存結果
        # key = sheet 名稱
        # value = DataFrame
        # ---------------------------------------------
        result_sheets[sheet_name] = result_df


    # =================================================
    # 8. 輸出 result.xlsx
    # =================================================
    result_path = os.path.join(work, "result.xlsx")


    # =================================================
    # 建立 ExcelWriter
    # =================================================
    with pd.ExcelWriter(
        result_path,
        engine="openpyxl"
    ) as writer:

        # 逐一輸出 sheet
        for sheet_name, df in result_sheets.items():

            # 寫入 Excel
             # 若 sheet 不存在就建立
            if sheet_name not in writer.book.sheetnames:
                writer.book.create_sheet(sheet_name)

            # 取得 worksheet
            ws = writer.book[sheet_name]

            # =========================
            # 從 A3 開始貼
            # =========================
            start_row = 3
            start_col = 1

            # 寫入資料（只改值）
            for r_idx, row in enumerate(df.values):

                for c_idx, value in enumerate(row):

                    ws.cell(
                        row=start_row + r_idx,
                        column=start_col + c_idx
                    ).value = value


    # =================================================
    # 9. 回傳結果檔案
    # =================================================
    return FileResponse(

        # 要下載的檔案
        result_path,

        # Excel MIME 類型
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

        # 下載檔名
        filename="result.xlsx"
    )
