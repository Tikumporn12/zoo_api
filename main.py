from fastapi import FastAPI, HTTPException
from openpyxl import load_workbook
from pathlib import Path

app = FastAPI(
    title="Zoo Visitors API",
    description="API for Zoo Visitors Fiscal Year 2569",
    version="1.0.0"
)

# ตำแหน่งไฟล์ Excel
EXCEL_FILE = Path(__file__).resolve().parent / "zoovisitors.xlsx"

# ข้อมูลอ้างอิง
SOURCE = "https://data.go.th/dataset/zoo-visitors2569"
RESOURCE_ID = "0ecc96cf-d914-42a1-b184-dbbd013c3c9c"
FISCAL_YEAR = 2569


# อ่านข้อมูลจาก Excel
def load_zoo_data():
    wb = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True
    )

    try:
        ws = wb["statistics"]

        headers = next(ws.iter_rows(values_only=True))

        data = []

        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue

            item = {
                headers[i]: row[i]
                for i in range(15)
            }

            data.append(item)

        return data

    finally:
        wb.close()


@app.get("/")
def read_root():
    return {
        "message": "Zoo Visitors API is running"
    }


# แสดงข้อมูลผู้เข้าชมทั้งหมด
@app.get("/api/zoo-visitors")
def get_zoo_visitors():
    return {
        "source": SOURCE,
        "resource_id": RESOURCE_ID,
        "fiscal_year": FISCAL_YEAR,
        "data": load_zoo_data()
    }


# ค้นหาข้อมูลตามสวนสัตว์
@app.get("/api/zoo-visitors/search")
def search_zoo_visitors(zoo: str):
    data = load_zoo_data()

    results = [
        item for item in data
        if item["สวนสัตว์"] == zoo
    ]

    if not results:
        raise HTTPException(
            status_code=404,
            detail="Zoo not found"
        )

    return {
        "source": SOURCE,
        "resource_id": RESOURCE_ID,
        "fiscal_year": FISCAL_YEAR,
        "data": results
    }