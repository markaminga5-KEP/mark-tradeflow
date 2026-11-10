import io, uuid
from pathlib import Path
import pandas as pd
from docx import Document
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

app = FastAPI()
OUT = Path("outputs"); OUT.mkdir(exist_ok=True)
Path("static").mkdir(exist_ok=True)
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/", response_class=HTMLResponse)
def home():
    return Path("static/index.html").read_text(encoding="utf-8")

@app.post("/generate")
async def generate(files: list[UploadFile] = File(...), top_n: int = 10):
    job = uuid.uuid4().hex[:8]
    tables = {}
    for f in files:
        raw = await f.read()
        name = f.filename.lower()
        if name.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(raw))
        elif name.endswith(".xlsx"):
            df = pd.read_excel(io.BytesIO(raw))
        else:
            continue
        tables[Path(f.filename).stem[:25]] = df.head(top_n)
    if not tables:
        raise HTTPException(400, "No readable .csv or .xlsx files uploaded")

    xlsx_path = OUT / f"{job}.xlsx"
    with pd.ExcelWriter(xlsx_path) as writer:
        for sheet, df in tables.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

    doc = Document()
    doc.add_heading("Trade Flow Report", 0)
    for title, df in tables.items():
        doc.add_heading(title, 1)
        t = doc.add_table(rows=1, cols=len(df.columns))
        t.style = "Light Grid Accent 1"
        for i, c in enumerate(df.columns):
            t.rows[0].cells[i].text = str(c)
        for _, row in df.iterrows():
            cells = t.add_row().cells
            for i, v in enumerate(row):
                cells[i].text = str(v)
    docx_path = OUT / f"{job}.docx"
    doc.save(docx_path)

    return {"docx": f"/download/{job}.docx", "xlsx": f"/download/{job}.xlsx"}

@app.get("/download/{name}")
def download(name: str):
    path = OUT / Path(name).name
    if not path.exists():
        raise HTTPException(404, "File not found")
    return FileResponse(path, filename=path.name)
