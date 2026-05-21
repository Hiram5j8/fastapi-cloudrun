from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from openpyxl import load_workbook
from dataclasses import dataclass
from typing import Any
from io import BytesIO
import zipfile
from fastapi.staticfiles import StaticFiles

app = FastAPI()

app.mount("/static", StaticFiles(directory="static", html=True), name="static")


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
    values: list[Any]

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




# =========================
# xls 轉 xlsx
# =========================
def convert_xls_to_xlsx(xls_path, out_dir):

    subprocess.run([
        "libreoffice",
        "--headless",
        "--convert-to",
        "xlsx",
        xls_path,
        "--outdir",
        out_dir
    ], check=True)

    filename = os.path.basename(xls_path)
    xlsx_name = os.path.splitext(filename)[0] + ".xlsx"

    return os.path.join(out_dir, xlsx_name)


# =========================
# 複製儲存格（保格式）
# =========================
def copy_cell(src_cell, dst_cell):

    dst_cell.value = src_cell.value

    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.border = copy(src_cell.border)
        dst_cell.alignment = copy(src_cell.alignment)
        dst_cell.number_format = copy(src_cell.number_format)
        dst_cell.protection = copy(src_cell.protection)


# =========================
# 合併工作表
# =========================
def merge_sheet(src_ws, dst_ws):

    max_row = src_ws.max_row
    max_col = src_ws.max_column

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):

            src_cell = src_ws.cell(r, c)
            dst_cell = dst_ws.cell(r, c)

            copy_cell(src_cell, dst_cell)

    # 欄寬
    for col_letter, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col_letter].width = dim.width

    # 列高
    for row_num, dim in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[row_num].height = dim.height

    # 合併儲存格
    for merged_range in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged_range))



# =========================
# API
# =========================
@app.post("/process")
async def process(data_zip: UploadFile = File(...)):

    # 建立暫存目錄
    work_dir = tempfile.mkdtemp(prefix="excel_")

    zip_path = os.path.join(work_dir, "data.zip")

    with open(zip_path, "wb") as f:
        shutil.copyfileobj(data_zip.file, f)

    # 解壓 ZIP
    extract_dir = os.path.join(work_dir, "extract")
    os.makedirs(extract_dir)

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

    # 建立 result workbook
    from openpyxl import Workbook
    result_wb = Workbook()

    # 移除預設 sheet
    default_sheet = result_wb.active
    result_wb.remove(default_sheet)

    # 處理所有 xls/xlsx
    for file_name in os.listdir(extract_dir):

        file_path = os.path.join(extract_dir, file_name)

        # xls -> xlsx
        if file_name.lower().endswith(".xls"):

            xlsx_path = convert_xls_to_xlsx(
                file_path,
                extract_dir
            )

        elif file_name.lower().endswith(".xlsx"):

            xlsx_path = file_path

        else:
            continue

        # 讀取 workbook
        src_wb = load_workbook(xlsx_path)

        # 所有 sheet
        for sheet_name in src_wb.sheetnames:

            src_ws = src_wb[sheet_name]

            # result sheet
            if sheet_name in result_wb.sheetnames:
                dst_ws = result_wb[sheet_name]
            else:
                dst_ws = result_wb.create_sheet(sheet_name)

            # 找最後一列
            start_row = dst_ws.max_row + 1

            if dst_ws.max_row == 1 and dst_ws["A1"].value is None:
                start_row = 1

            # 複製資料
            for r in range(1, src_ws.max_row + 1):
                for c in range(1, src_ws.max_column + 1):

                    src_cell = src_ws.cell(r, c)
                    dst_cell = dst_ws.cell(
                        start_row + r - 1,
                        c
                    )

                    copy_cell(src_cell, dst_cell)

    # 強制重新計算公式
    result_wb.calculation.fullCalcOnLoad = True

    # 輸出
    result_path = os.path.join(work_dir, "result.xlsx")

    result_wb.save(result_path)

    return FileResponse(
        result_path,
        filename="result.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )