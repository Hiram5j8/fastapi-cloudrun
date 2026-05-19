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
    import logging

    work = "/tmp/work"
    work2 = "/tmp/work2"
    os.makedirs(work, exist_ok=True)

    base_path = os.path.join(work, base_xls.filename)
    zip_path = os.path.join(work2, data_zip.filename)
    
    #shutil.copy(master_file, out_file)

    with open(base_path, "wb") as f:
        shutil.copyfileobj(base_xls.file, f)
         
    #base_df = pd.read_excel(base_path)
    
    base_path = os.path.join(work, "result.xlsx")
    result_wb.save(base_path)

    return FileResponse(
        base_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="result.xlsx"
    )