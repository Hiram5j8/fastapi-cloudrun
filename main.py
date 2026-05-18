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
    return {"message": "File API running  xxxxxxxxxxxxxxxx 🚀"}
# 掛載靜態網頁資料夾
app.mount("/", StaticFiles(directory="static", html=True), name="static")

@app.post("/process")
async def process_file(file: UploadFile = File(...)):
    input_path = f"/tmp/{file.filename}"
    output_path = f"/tmp/processed_{file.filename}"

    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # === 這裡放你的 Python 處理程式 ===
    with open(input_path, "r", encoding="utf-8") as f:
        data = f.read()

    processed_data = data.upper()  # ← 換成你的邏輯

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(processed_data)
    # ====================================

    return FileResponse(output_path, filename=f"processed_{file.filename}")
