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

@app.post("/process-file")
async def process_file(file: UploadFile = File(...)):
    input_path = f"/tmp/{file.filename}"
    output_path = f"/tmp/processed_{file.filename}"

    # 存檔
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    print("input exists:", os.path.exists(input_path))

    # 測試：單純複製檔案（保證成功）
    shutil.copy(input_path, output_path)

    print("output exists:", os.path.exists(output_path))

    return FileResponse(output_path, filename=f"processed_{file.filename}")
