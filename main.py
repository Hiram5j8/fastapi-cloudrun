from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

import pandas as pd
import zipfile
import shutil
import os

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ===== 掛載靜態網頁（最重要）=====
app.mount("/", StaticFiles(directory="static", html=True), name="static")

WORK_DIR = "/tmp/work"
os.makedirs(WORK_DIR, exist_ok=True)


# ===== 共用 Excel 規則 =====
def extract_rule_from_excel(path):
    df = pd.read_excel(path, header=None)

    key = df.iloc[2, 0]           # A3
    row_data = df.iloc[2, 2:37]  # C3~AK3

    groups = [row_data[i:i+5].tolist() for i in range(0, len(row_data), 5)]

    rows = []
    for idx, g in enumerate(groups, start=1):
        rows.append([key, f"Group{idx}"] + g)

    return rows


# ===== 單一 XLS =====
@app.post("/api/process-xls")
async def process_xls(file: UploadFile = File(...)):
    input_path = f"{WORK_DIR}/{file.filename}"

    with open(input_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    rows = extract_rule_from_excel(input_path)

    df_out = pd.DataFrame(rows,
                          columns=["Key(A3)", "Group", "V1", "V2", "V3", "V4", "V5"])

    output_path = f"{WORK_DIR}/result.xlsx"
    df_out.to_excel(output_path, index=False)

    return FileResponse(output_path, filename="result.xlsx")


# ===== ZIP 彙整 =====
@app.post("/api/process-zip")
async def process_zip(file: UploadFile = File(...)):
    zip_path = f"{WORK_DIR}/upload.zip"

    with open(zip_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    extract_dir = f"{WORK_DIR}/unzipped"
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

    all_rows = []

    for root, dirs, files in os.walk(extract_dir):
        for name in files:
            if name.endswith((".xls", ".xlsx")):
                full_path = os.path.join(root, name)
                rows = extract_rule_from_excel(full_path)
                all_rows.extend(rows)

    df_out = pd.DataFrame(all_rows,
                          columns=["Key(A3)", "Group", "V1", "V2", "V3", "V4", "V5"])

    output_path = f"{WORK_DIR}/summary.xlsx"
    df_out.to_excel(output_path, index=False)

    return FileResponse(output_path, filename="summary.xlsx")