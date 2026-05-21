from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from openpyxl import load_workbook
import os
import shutil
import zipfile
import tempfile
import pandas as pd



app = FastAPI()

# ✅ static 一定只能掛 /static（安全）
app.mount("/static", StaticFiles(directory="static", html=True), name="static")

# ✅ 首頁
@app.get("/")
def home():
    return FileResponse("static/index.html")



@app.post("/process")
async def process(
    base_xls: UploadFile = File(...),
    data_zip: UploadFile = File(...)
):

    # ==========================================
    # 建立工作目錄
    # ==========================================
    work = tempfile.mkdtemp(prefix="work_")

    # ==========================================
    # 儲存 base.xlsx
    # ==========================================
    base_path = os.path.join(work, "base.xlsx")

    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)

    # ==========================================
    # 儲存 ZIP
    # ==========================================
    zip_path = os.path.join(work, "data.zip")

    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # ==========================================
    # 解壓 ZIP
    # ==========================================
    unzip_dir = os.path.join(work, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(unzip_dir)

    # ==========================================
    # 全部 Excel 讀進記憶體
    # ==========================================
    all_sheets = []

    for root, dirs, files in os.walk(unzip_dir):

        for file in files:

            if file.lower().endswith((".xls", ".xlsx")):

                excel_path = os.path.join(root, file)

                print(f"讀取檔案: {file}")

                try:

                    xls = pd.ExcelFile(excel_path)

                    for sheet_name in xls.sheet_names:

                        print(f"  Sheet: {sheet_name}")

                        # ==========================
                        # 讀取 sheet
                        # ==========================
                        df = pd.read_excel(
                            xls,
                            sheet_name=sheet_name,
                            header=None
                        )

                        # ==========================
                        # 清除全空白列
                        # ==========================
                        df = df.dropna(how="all")

                        # ==========================
                        # 清除全空白欄
                        # ==========================
                        df = df.dropna(axis=1, how="all")

                        # ==========================
                        # NaN → None
                        # ==========================
                        df = df.where(pd.notna(df), None)

                        # ==========================
                        # 空白字串 → None
                        # ==========================
                        df = df.map(
                            lambda x: None
                            if isinstance(x, str) and x.strip() == ""
                            else x
                        )

                        # ==========================
                        # 加入記憶體
                        # ==========================
                        all_sheets.append({
                            "file": file,
                            "sheet": sheet_name,
                            "data": df
                        })

                except Exception as e:

                    print("讀取失敗:", file)
                    print(e)

    print("================================")
    print("總 sheet 數:", len(all_sheets))
    print("================================")

    # ==========================================
    # 開啟 base.xlsx
    # ==========================================
    wb = load_workbook(base_path)

    # 使用第一張工作表
    ws = wb.active

    # ==========================================
    # 開始寫入
    # ==========================================

    # 從 A3 開始
    start_row = 3

    for idx, item in enumerate(all_sheets):

        file_name = item["file"]
        sheet_name = item["sheet"]
        df = item["data"]

        print(f"寫入 [{idx+1}] {file_name} - {sheet_name}")

        # ==========================
        # 最大範圍限制
        # A~AK = 37欄
        # ==========================
        max_r = min(len(df), 34)
        max_c = min(len(df.columns), 37)

        # ==========================
        # 寫入 Excel
        # ==========================
        for r in range(max_r):

            for c in range(max_c):

                value = df.iat[r, c]

                # ======================
                # 忽略空值
                # ======================
                if value is None:
                    continue

                # ======================
                # 字串去空白
                # ======================
                if isinstance(value, str):
                    value = value.strip()

                    if value == "":
                        continue

                # ======================
                # 有值才寫入
                # ======================
                ws.cell(
                    row=start_row + r,
                    column=1 + c,
                    value=value
                )

        # ==================================
        # 下一個區塊
        # 每份資料間隔40列
        # ==================================
        start_row += 40

    # ==========================================
    # 輸出 result.xlsx
    # ==========================================
    result_path = os.path.join(work, "result.xlsx")

    wb.save(result_path)

    print("完成:", result_path)

    return FileResponse(
        result_path,
        filename="result.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )