from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import os
import shutil

app = FastAPI()

@app.post("/process")
async def process(base_xls: UploadFile = File(...), data_zip: UploadFile = File(...)):
    work = "/tmp/work"
    os.makedirs(work, exist_ok=True)

    # 上傳檔案存起來
    upload_path = os.path.join(work, base_xls.filename)

    with open(upload_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)

    # 直接複製成 result.xlsx
    result_path = os.path.join(work, "result.xlsx")
    shutil.copy(upload_path, result_path)

    # 回傳檔案
    return FileResponse(
        result_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="result.xlsx"
    )