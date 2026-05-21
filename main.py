from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles

from openpyxl import load_workbook

from dataclasses import dataclass
from typing import Any, List

from io import BytesIO
import zipfile

# =====================================================
# FastAPI
# =====================================================

app = FastAPI()

# ✅ static
app.mount(
    "/static",
    StaticFiles(directory="static", html=True),
    name="static"
)

# ✅ 首頁
@app.get("/")
def home():
    return FileResponse("static/index.html")

# =====================================================
# RowSeries Engine
# =====================================================

@dataclass
class RowSeries:
    sheet: str
    start_row: int
    start_col: int
    values: List[Any]

# =====================================================
# 建立 base map
# sheet -> row -> col -> value
# =====================================================

def build_base_map(wb):

    base_map = {}

    for sheet in wb.sheetnames:

        ws = wb[sheet]

        sheet_map = {}

        for row in ws.iter_rows():

            for cell in row:

                if cell.value is None:
                    continue

                r = cell.row
                c = cell.column

                if r not in sheet_map:
                    sheet_map[r] = {}

                sheet_map[r][c] = cell.value

        base_map[sheet] = sheet_map

    return base_map

# =====================================================
# ZIP Excel -> RowSeries(diff only)
# =====================================================

def extract_diff_rows(zip_bytes, base_map):

    row_series_list = []

    with zipfile.ZipFile(BytesIO(zip_bytes)) as z:

        for filename in z.namelist():

            # 只處理 xlsx
            if not filename.endswith(".xlsx"):
                continue

            print("處理:", filename)

            with z.open(filename) as f:

                wb = load_workbook(
                    BytesIO(f.read()),
                    data_only=True,
                    read_only=True
                )

                for sheet in wb.sheetnames:

                    if sheet not in base_map:
                        continue

                    ws = wb[sheet]

                    for row in ws.iter_rows():

                        values = []
                        changed = False

                        row_num = row[0].row

                        for cell in row:

                            col_num = cell.column

                            new_val = cell.value

                            base_val = (
                                base_map
                                .get(sheet, {})
                                .get(row_num, {})
                                .get(col_num)
                            )

                            # 相同 -> None
                            if new_val == base_val:
                                values.append(None)

                            else:
                                values.append(new_val)
                                changed = True

                        # 整列沒變 -> skip
                        if not changed:
                            continue

                        rs = RowSeries(
                            sheet=sheet,
                            start_row=row_num,
                            start_col=1,
                            values=values
                        )

                        row_series_list.append(rs)

    return row_series_list

# =====================================================
# RowSeries -> Excel
# =====================================================

def write_series(ws, rs: RowSeries):

    for offset, value in enumerate(rs.values):

        # None 跳過
        if value is None:
            continue

        ws.cell(
            row=rs.start_row,
            column=rs.start_col + offset
        ).value = value

# =====================================================
# API
# =====================================================

@app.post("/merge")
async def merge(
    base_xlsx: UploadFile = File(...),
    data_zip: UploadFile = File(...)
):

    # =====================================
    # 1. 讀 base.xlsx
    # =====================================

    base_bytes = await base_xlsx.read()

    base_wb = load_workbook(BytesIO(base_bytes))

    # =====================================
    # 2. 建立 base map
    # =====================================

    base_map = build_base_map(base_wb)

    # =====================================
    # 3. 讀 ZIP
    # =====================================

    zip_bytes = await data_zip.read()

    # =====================================
    # 4. 建立 RowSeries(diff only)
    # =====================================

    rows = extract_diff_rows(zip_bytes, base_map)

    print("diff rows:", len(rows))

    # =====================================
    # 5. 寫入 result
    # =====================================

    result_wb = load_workbook(BytesIO(base_bytes))

    for rs in rows:

        if rs.sheet not in result_wb.sheetnames:
            continue

        ws = result_wb[rs.sheet]

        write_series(ws, rs)

    # =====================================
    # 6. 輸出記憶體
    # =====================================

    output = BytesIO()

    result_wb.save(output)

    output.seek(0)

    # =====================================
    # 7. 回傳 result.xlsx
    # =====================================

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": "attachment; filename=result.xlsx"
        }
    )