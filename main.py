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
    work = "/tmp/work"
    os.makedirs(work, exist_ok=True)

    #base_path = os.path.join(work, base_xls.filename)
    zip_path = os.path.join(work, data_zip.filename)

    #with open(base_path, "wb") as f:
        #shutil.copyfileobj(base_xls.file, f)

    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # Excel
    base_df = pd.read_excel(base_path)
    base_key = base_df.iloc[2, 0]

    unzip_dir = os.path.join(work, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, 'r') as z:
        z.extractall(unzip_dir)

    result_wb = Workbook()
    result_ws = result_wb.active

    result_path = os.path.join(work, "result.xlsx")
    result_wb.save(result_path)

    #return FileResponse(result_path, filename="result.xlsx")


    return FileResponse(
        result_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="result.xlsx"
    )
