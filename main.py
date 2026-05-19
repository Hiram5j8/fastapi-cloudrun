from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
import zipfile, shutil, os
import pandas as pd
from openpyxl import Workbook

app = FastAPI()
app.mount("/", StaticFiles(directory="static", html=True), name="static")


def xls_to_xlsx(path):
    if path.endswith(".xls"):
        df = pd.read_excel(path, engine="xlrd")
        new_path = path + "x"
        df.to_excel(new_path, index=False)
        return new_path
    return path


@app.post("/process")
async def process(base_xls: UploadFile = File(...), data_zip: UploadFile = File(...)):
    work = "/tmp/work"
    os.makedirs(work, exist_ok=True)

    base_path = os.path.join(work, base_xls.filename)
    zip_path = os.path.join(work, data_zip.filename)

    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)
    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # 轉 base.xls
    base_path = xls_to_xlsx(base_path)
    base_df = pd.read_excel(base_path)

    base_key = base_df.iloc[2, 0]  # A3

    unzip_dir = os.path.join(work, "unzipped")
    os.makedirs(unzip_dir, exist_ok=True)

    with zipfile.ZipFile(zip_path, 'r') as z:
        z.extractall(unzip_dir)

    result_wb = Workbook()
    result_ws = result_wb.active

    for root, _, files in os.walk(unzip_dir):
        for file in files:
            if file.endswith((".xls", ".xlsx")):
                fpath = os.path.join(root, file)
                fpath = xls_to_xlsx(fpath)

                df = pd.read_excel(fpath)

                if df.iloc[2, 0] == base_key:
                    for c in range(2, 37, 5):  # C~AK
                        row = df.iloc[2, c:c+5].tolist()
                        result_ws.append(row)

    result_path = os.path.join(work, "result.xlsx")
    result_wb.save(result_path)

    return FileResponse(result_path, filename="result.xlsx")