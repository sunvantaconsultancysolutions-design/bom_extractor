import base64
import io
import os
import sys
import traceback
from pathlib import Path

import openpyxl
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, Response, StreamingResponse
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pydantic import BaseModel

print(f"[STARTUP] Python {sys.version}", file=sys.stderr)
print("[STARTUP] BOM Extractor v5.0 — browser-direct mode", file=sys.stderr)

app = FastAPI(title="BOM Extractor", version="5.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------------------------------------------------------------------
# favicon
# ---------------------------------------------------------------------------

_FAVICON_ICO = base64.b64decode(
    "AAABAAEAAQEAAAEAGAAwAAAAFgAAACgAAAABAAAAAgAAAAEAGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8AAAA="
)

@app.get("/favicon.ico", include_in_schema=False)
def favicon():
    return Response(content=_FAVICON_ICO, media_type="image/x-icon")


# ---------------------------------------------------------------------------
# health
# ---------------------------------------------------------------------------

@app.get("/api/health", include_in_schema=False)
def health():
    return {"status": "ok", "version": "5.0.0", "mode": "browser-direct"}


# ---------------------------------------------------------------------------
# frontend
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse, include_in_schema=False)
def root():
    html_path = Path(__file__).parent / "index.html"
    if not html_path.exists():
        return "<h1>index.html not found</h1>"
    return html_path.read_text(encoding="utf-8")


# ---------------------------------------------------------------------------
# Excel styling helpers
# ---------------------------------------------------------------------------

def _thin():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

HDR_FILL = PatternFill("solid", start_color="1F4E79")
SUB_FILL  = PatternFill("solid", start_color="2E75B6")
ALT_FILL  = PatternFill("solid", start_color="EBF3FB")

CENTER = Alignment(horizontal="center", vertical="center")
LEFT   = Alignment(horizontal="left",   vertical="center")
RIGHT  = Alignment(horizontal="right",  vertical="center")

def _cell(cell, value, fill=None, align=None, bold=False):
    cell.value = value
    if fill:
        cell.fill = fill
    cell.border = _thin()
    cell.font = Font(
        name="Arial", size=9, bold=bold,
        color="FFFFFF" if fill in (HDR_FILL, SUB_FILL) else "000000",
    )
    cell.alignment = align or LEFT


# ---------------------------------------------------------------------------
# Excel builder (server-side — kept as fallback, browser does it by default)
# ---------------------------------------------------------------------------

def build_excel(sections: list, filename: str) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BOM"
    row = 1

    for sec in sections:
        ws.merge_cells(f"A{row}:E{row}")
        _cell(ws.cell(row=row, column=1), sec.get("title", "BILL OF MATERIALS"), HDR_FILL, CENTER, True)
        row += 1

        unit = sec.get("unit", "")
        if unit:
            ws.merge_cells(f"A{row}:E{row}")
            _cell(ws.cell(row=row, column=1), f"Unit: {unit}", SUB_FILL, CENTER)
            row += 1

        profiles = sec.get("profiles", [])
        if profiles:
            ws.merge_cells(f"A{row}:E{row}")
            _cell(ws.cell(row=row, column=1), "STRUCTURAL PROFILES", SUB_FILL, CENTER, True)
            row += 1
            for col, h in enumerate(["MARK", "QTY", "PROFILE", "LENGTH (mm)", "WEIGHT (kg)"], 1):
                _cell(ws.cell(row=row, column=col), h, SUB_FILL, CENTER, True)
            row += 1
            for i, p in enumerate(profiles):
                fill = ALT_FILL if i % 2 == 0 else None
                for col, val in enumerate([p.get("mark"), p.get("qty"), p.get("profile"), p.get("length"), p.get("black_weight")], 1):
                    _cell(ws.cell(row=row, column=col), val, fill)
                row += 1
            tp = sec.get("total_profiles_weight")
            if tp is not None:
                _cell(ws.cell(row=row, column=4), "Total Profiles Weight:", None, RIGHT, True)
                _cell(ws.cell(row=row, column=5), tp, None, RIGHT, True)
                row += 1

        fasteners = sec.get("fasteners", [])
        if fasteners:
            row += 1
            ws.merge_cells(f"A{row}:E{row}")
            _cell(ws.cell(row=row, column=1), "FASTENERS", SUB_FILL, CENTER, True)
            row += 1
            for col, h in enumerate(["DESCRIPTION", "QTY", "WEIGHT (kg)"], 1):
                _cell(ws.cell(row=row, column=col), h, SUB_FILL, CENTER, True)
            row += 1
            for i, f in enumerate(fasteners):
                fill = ALT_FILL if i % 2 == 0 else None
                for col, val in enumerate([f.get("description"), f.get("qty"), f.get("black_weight")], 1):
                    _cell(ws.cell(row=row, column=col), val, fill)
                row += 1
            tf = sec.get("total_fasteners_weight")
            if tf is not None:
                _cell(ws.cell(row=row, column=2), "Total Fasteners Weight:", None, RIGHT, True)
                _cell(ws.cell(row=row, column=3), tf, None, RIGHT, True)
                row += 1

        dt = sec.get("drawing_total_weight")
        if dt is not None:
            row += 1
            ws.merge_cells(f"A{row}:D{row}")
            _cell(ws.cell(row=row, column=1), "DRAWING TOTAL WEIGHT (kg):", HDR_FILL, RIGHT, True)
            _cell(ws.cell(row=row, column=5), dt, HDR_FILL, CENTER, True)
            row += 1

        row += 2

    for col, w in zip("ABCDE", [22, 10, 28, 18, 18]):
        ws.column_dimensions[col].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# /api/build-excel  (optional server-side fallback)
# ---------------------------------------------------------------------------

class BuildExcelRequest(BaseModel):
    filename: str
    sections: list

@app.post("/api/build-excel")
async def build_excel_endpoint(payload: BuildExcelRequest):
    if not payload.sections:
        raise HTTPException(status_code=422, detail="No BOM sections provided")
    try:
        excel_bytes = build_excel(payload.sections, payload.filename)
        stem = Path(payload.filename).stem
        return StreamingResponse(
            io.BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{stem}_BOM.xlsx"'},
        )
    except Exception as e:
        print(traceback.format_exc(), file=sys.stderr)
        raise HTTPException(status_code=500, detail=f"Excel build failed: {str(e)}")
