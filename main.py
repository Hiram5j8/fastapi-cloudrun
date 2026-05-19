import os
import sys
import glob
import math
import shutil
import hashlib
import unicodedata
import numpy as np
import pandas as pd
import tkinter as tk

from datetime import datetime
from openpyxl import load_workbook
from tkinter import filedialog, messagebox

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