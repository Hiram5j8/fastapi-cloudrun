from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import shutil
import os

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def root():
    return {"message": "File API running 🚀"}

@app.post("/process-file")
async def process_file(file: UploadFile = File(...)):
    # 存上傳檔案
    input_path = f"/tmp/{file.filename}"
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # === 這裡放你的 Python 處理程式 ===
    output_path = f"/tmp/processed_{file.filename}"
    with open(input_path, "r") as fin, open(output_path, "w") as fout:
        content = fin.read()
        fout.write("Processed:\n" + content)
    # =================================

    return FileResponse(output_path, filename=f"processed_{file.filename}")