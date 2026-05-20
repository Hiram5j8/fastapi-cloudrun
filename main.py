import os
import glob
import zipfile
import shutil
import tempfile

from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from openpyxl import load_workbook

app = FastAPI()

@app.post("/process")
async def process(
    base_xls: UploadFile = File(...),
    data_zip: UploadFile = File(...)
):
    # === 建立工作目錄 ===
    work = tempfile.mkdtemp(prefix="work_")

    # === 儲存 base.xlsx ===
    base_path = os.path.join(work, "base.xlsx")
    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)

    # === 複製成 result.xlsx（保留 base 結構）===
    result_path = os.path.join(work, "result.xlsx")
    shutil.copyfile(base_path, result_path)

    # === 儲存 ZIP ===
    zip_path = os.path.join(work, "data.zip")
    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # === 解壓 ZIP ===
    unzip_dir = os.path.join(work, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, 'r') as z:
        z.extractall(unzip_dir)

    # === 找 ZIP 內 Excel（必須只有一個）===
    files = glob.glob(os.path.join(unzip_dir, "**", "*.xls*"), recursive=True)


    # === 開啟來源 Excel & result Excel ===
    src_wb = load_workbook(src_excel, data_only=False)
    dst_wb = load_workbook(result_path)

    # === 將來源每個 sheet 複製到 result ===
    for sh in src_wb.sheetnames:

        src_ws = src_wb[sh]

        # result 沒這個 sheet 就建立
        if sh not in dst_wb.sheetnames:
            dst_ws = dst_wb.create_sheet(sh)
        else:
            dst_ws = dst_wb[sh]

        # 複製所有儲存格內容
        for row in src_ws.iter_rows():
            for cell in row:
                dst_ws.cell(
                    row=cell.row,
                    column=cell.column
                ).value = cell.value

    # === 儲存 result.xlsx ===
    dst_wb.save(result_path)

    # === 回傳下載 ===
    return FileResponse(
        result_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="result.xlsx"
    )