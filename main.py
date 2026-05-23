"""
BOM Extractor – Vercel-ready FastAPI app
Entrypoint: main.py  (Vercel auto-detects FastAPI here)
"""

import base64
import io
import json
import os
import traceback
from pathlib import Path
from typing import Optional

import fitz  # PyMuPDF
import openpyxl
import anthropic
from fastapi import FastAPI, File, Header, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, Response, StreamingResponse
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pydantic import BaseModel

# ─── App ──────────────────────────────────────────────────────────────────────
app = FastAPI(title="BOM Extractor", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─── Favicon (avoid 404) ──────────────────────────────────────────────────────
# Minimal 1x1 transparent ICO to prevent browser 404 errors
_FAVICON_ICO = base64.b64decode(
    "AAABAAEAAQEAAAEAGAAwAAAAFgAAACgAAAABAAAAAgAAAAEAGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8AAAA="
)

@app.get("/favicon.ico", include_in_schema=False)
def favicon():
    return Response(content=_FAVICON_ICO, media_type="image/x-icon")

# ─── Serve frontend ───────────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse, include_in_schema=False)
def root():
    """Serve the single-page frontend."""
    html_path = Path(__file__).parent / "index.html"
    return html_path.read_text()

# ─── Claude system prompt ─────────────────────────────────────────────────────
SYSTEM_PROMPT = """You are an expert engineering drawing parser specialising in steel tower Bills of Materials (BOM).

Given a page image from an engineering drawing, extract ALL BOM tables visible on the page.

Return ONLY a valid JSON object (no markdown, no explanation) with this structure:
{
  "sections": [
    {
      "title": "BILL OF MATERIALS FOR (CAGE-1)",
      "unit": "QTY. PER (TOWER)",
      "profiles": [
        {"mark": "100H", "qty": 2, "profile": "L 90x6 H", "length": 4981, "black_weight": 82.66}
      ],
      "fasteners": [
        {"description": "M 16x40", "qty": 184, "black_weight": 28.46}
      ],
      "total_profiles_weight": 5100.80,
      "total_fasteners_weight": 210.83,
      "drawing_total_weight": 5311.62
    }
  ],
  "has_bom": true
}

If no BOM table is found on the page return: {"has_bom": false, "sections": []}

Rules:
- Extract every row — do not skip any.
- length is in mm (integer), black_weight is in kg (float).
- If a cell is empty/unclear use null.
- For fastener rows that have no separate length column, set length to null.
- Preserve the exact mark/description text (e.g. "207H", "150AH", "Filler Washer M 16 #8").
"""


# ─── Helpers ──────────────────────────────────────────────────────────────────
def pdf_to_images(pdf_bytes: bytes, dpi: int = 150) -> list[tuple[int, bytes]]:
    """Convert PDF pages to JPEG images. DPI=150 balances quality vs payload size."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    for i, page in enumerate(doc):
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
        # Use quality=80 to reduce JPEG size while keeping text readable
        images.append((i + 1, pix.tobytes(output="jpeg")))
    doc.close()
    return images


def extract_bom_from_page(client: anthropic.Anthropic, page_num: int, jpeg_bytes: bytes) -> dict:
    b64 = base64.standard_b64encode(jpeg_bytes).decode()
    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
                {"type": "text", "text": f"Extract all BOM tables from this engineering drawing page (page {page_num})."},
            ],
        }],
    )
    raw = msg.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw)


# ─── Excel builder ────────────────────────────────────────────────────────────
def _thin():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def _thick():
    s = Side(style="medium", color="2E75B6")
    return Border(left=s, right=s, top=s, bottom=s)

HDR_FILL  = PatternFill("solid", start_color="1F4E79")
SUB_FILL  = PatternFill("solid", start_color="2E75B6")
SEC_FILL  = PatternFill("solid", start_color="D6E4F0")
ALT_FILL  = PatternFill("solid", start_color="EBF3FB")
TOT_FILL  = PatternFill("solid", start_color="FFF2CC")
GTOT_FILL = PatternFill("solid", start_color="F4B942")
WHT_FILL  = PatternFill("solid", start_color="FFFFFF")
C = Alignment(horizontal="center", vertical="center", wrap_text=True)
L = Alignment(horizontal="left",   vertical="center", wrap_text=True)
R = Alignment(horizontal="right",  vertical="center")


def sc(cell, val, fill=None, font=None, align=None, fmt=None):
    cell.value = val
    cell.fill  = fill or WHT_FILL
    cell.font  = font or Font(name="Arial", size=9)
    cell.alignment = align or L
    cell.border = _thin()
    if fmt:
        cell.number_format = fmt


def build_excel(all_sections: list[dict], filename: str) -> bytes:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    summary_rows = []

    for sec in all_sections:
        title_raw = sec.get("title", "BOM")
        short = title_raw.replace("BILL OF MATERIALS FOR", "").strip("() ").replace("/", "-")[:31] or "BOM"
        ws = wb.create_sheet(title=short)
        ws.sheet_view.showGridLines = False

        ws.merge_cells("A1:F1")
        c = ws["A1"]; c.value = title_raw
        c.fill = HDR_FILL; c.font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        c.alignment = C; c.border = _thin(); ws.row_dimensions[1].height = 22

        ws.merge_cells("A2:F2")
        c = ws["A2"]; c.value = sec.get("unit", "")
        c.fill = SUB_FILL; c.font = Font(name="Arial", bold=True, color="FFFFFF", size=9)
        c.alignment = C; c.border = _thin(); ws.row_dimensions[2].height = 14

        row = 3
        for col, h in enumerate(["MARK","QTY","PROFILE","LENGTH (mm)","BLACK WEIGHT (kg)","NOTES"], 1):
            c = ws.cell(row=row, column=col); c.value = h; c.fill = SUB_FILL
            c.font = Font(name="Arial", bold=True, color="FFFFFF", size=9)
            c.alignment = C; c.border = _thin()
        ws.row_dimensions[row].height = 26; row += 1

        profiles = sec.get("profiles", [])
        if profiles:
            ws.merge_cells(f"A{row}:F{row}")
            c = ws.cell(row=row, column=1); c.value = "▶  STRUCTURAL PROFILES"
            c.fill = SEC_FILL; c.font = Font(name="Arial", bold=True, size=9, color="1F4E79")
            c.alignment = L; c.border = _thin(); ws.row_dimensions[row].height = 14; row += 1
            for i, p in enumerate(profiles):
                fill = ALT_FILL if i % 2 == 0 else WHT_FILL
                ws.row_dimensions[row].height = 13
                sc(ws.cell(row=row,column=1), p.get("mark"),         fill, fmt="@")
                sc(ws.cell(row=row,column=2), p.get("qty"),          fill, align=C)
                sc(ws.cell(row=row,column=3), p.get("profile"),      fill)
                sc(ws.cell(row=row,column=4), p.get("length"),       fill, align=R, fmt="#,##0")
                sc(ws.cell(row=row,column=5), p.get("black_weight"), fill, align=R, fmt="#,##0.00")
                sc(ws.cell(row=row,column=6), "",                    fill); row += 1

        tp = sec.get("total_profiles_weight")
        if tp is not None:
            ws.merge_cells(f"A{row}:D{row}")
            c = ws.cell(row=row, column=1); c.value = "TOTAL BLACK WEIGHT (profiles) kg"
            c.fill = TOT_FILL; c.font = Font(name="Arial", bold=True, size=9, color="7F3F00")
            c.alignment = L; c.border = _thin()
            sc(ws.cell(row=row,column=5), tp, TOT_FILL, Font(name="Arial",bold=True,size=9,color="7F3F00"), R, "#,##0.00")
            sc(ws.cell(row=row,column=6), "", TOT_FILL); ws.row_dimensions[row].height = 15; row += 1

        fasteners = sec.get("fasteners", [])
        if fasteners:
            ws.merge_cells(f"A{row}:F{row}")
            c = ws.cell(row=row, column=1); c.value = "▶  FASTENERS"
            c.fill = SEC_FILL; c.font = Font(name="Arial", bold=True, size=9, color="1F4E79")
            c.alignment = L; c.border = _thin(); ws.row_dimensions[row].height = 14; row += 1
            for col, h in enumerate(["DESCRIPTION","QTY","BLACK WEIGHT (kg)","","",""], 1):
                c = ws.cell(row=row, column=col); c.value = h; c.fill = SUB_FILL
                c.font = Font(name="Arial", bold=True, color="FFFFFF", size=9)
                c.alignment = C; c.border = _thin()
            ws.row_dimensions[row].height = 14; row += 1
            for i, f in enumerate(fasteners):
                fill = ALT_FILL if i % 2 == 0 else WHT_FILL
                ws.merge_cells(f"A{row}:B{row}")
                sc(ws.cell(row=row,column=1), f.get("description"), fill)
                sc(ws.cell(row=row,column=2), "",                   fill)
                sc(ws.cell(row=row,column=3), f.get("qty"),         fill, align=C)
                sc(ws.cell(row=row,column=4), f.get("black_weight"),fill, align=R, fmt="#,##0.00")
                sc(ws.cell(row=row,column=5), "", fill); sc(ws.cell(row=row,column=6), "", fill)
                ws.row_dimensions[row].height = 13; row += 1

            tf = sec.get("total_fasteners_weight")
            if tf is not None:
                ws.merge_cells(f"A{row}:C{row}")
                c = ws.cell(row=row, column=1); c.value = "TOTAL BLACK WEIGHT (fasteners) kg"
                c.fill = TOT_FILL; c.font = Font(name="Arial", bold=True, size=9, color="7F3F00")
                c.alignment = L; c.border = _thin()
                sc(ws.cell(row=row,column=4), tf, TOT_FILL, Font(name="Arial",bold=True,size=9,color="7F3F00"), R, "#,##0.00")
                for col in [2,5,6]: sc(ws.cell(row=row,column=col), "", TOT_FILL)
                ws.row_dimensions[row].height = 15; row += 1

        dt = sec.get("drawing_total_weight")
        if dt is not None:
            ws.merge_cells(f"A{row}:D{row}")
            c = ws.cell(row=row, column=1); c.value = "DRAWING TOTAL BLACK WEIGHT (kg)"
            c.fill = GTOT_FILL; c.font = Font(name="Arial", bold=True, size=10)
            c.alignment = L; c.border = _thick()
            c2 = ws.cell(row=row,column=5); c2.value = dt; c2.fill = GTOT_FILL
            c2.font = Font(name="Arial",bold=True,size=10); c2.alignment = R
            c2.border = _thick(); c2.number_format = "#,##0.00"
            sc(ws.cell(row=row,column=6), "", GTOT_FILL); ws.row_dimensions[row].height = 18

        ws.column_dimensions["A"].width = 14; ws.column_dimensions["B"].width = 6
        ws.column_dimensions["C"].width = 18; ws.column_dimensions["D"].width = 14
        ws.column_dimensions["E"].width = 20; ws.column_dimensions["F"].width = 10
        ws.freeze_panes = "A4"

        summary_rows.append({"section": short, "title": title_raw,
                              "profiles": tp, "fasteners": sec.get("total_fasteners_weight"), "total": dt})

    # Summary sheet
    ws = wb.create_sheet("SUMMARY", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 20; ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 22; ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 22

    ws.merge_cells("A1:E1"); c = ws["A1"]; c.value = "BILL OF MATERIALS – SUMMARY"
    c.fill = HDR_FILL; c.font = Font(name="Arial", bold=True, color="FFFFFF", size=14)
    c.alignment = C; c.border = _thin(); ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:E2"); c = ws["A2"]; c.value = f"Source: {filename}"
    c.fill = SUB_FILL; c.font = Font(name="Arial", color="FFFFFF", size=9)
    c.alignment = C; c.border = _thin(); ws.row_dimensions[2].height = 14

    row = 4
    for col, h in enumerate(["SHEET","COMPONENT / SECTION","PROFILES WT. (kg)","FASTENERS WT. (kg)","DRAWING TOTAL (kg)"], 1):
        c = ws.cell(row=row, column=col); c.value = h; c.fill = SUB_FILL
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=9)
        c.alignment = C; c.border = _thin()
    ws.row_dimensions[row].height = 26; row += 1

    gp = gf = gt = 0.0
    for i, sr in enumerate(summary_rows):
        fill = ALT_FILL if i % 2 == 0 else WHT_FILL
        ws.row_dimensions[row].height = 15
        sc(ws.cell(row=row,column=1), sr["section"],   fill, Font(name="Arial",bold=True,size=9), C)
        sc(ws.cell(row=row,column=2), sr["title"],     fill)
        sc(ws.cell(row=row,column=3), sr["profiles"],  fill, align=R, fmt="#,##0.00")
        sc(ws.cell(row=row,column=4), sr["fasteners"], fill, align=R, fmt="#,##0.00")
        sc(ws.cell(row=row,column=5), sr["total"],     fill, Font(name="Arial",bold=True,size=9), R, "#,##0.00")
        gp += sr["profiles"] or 0; gf += sr["fasteners"] or 0; gt += sr["total"] or 0; row += 1

    ws.merge_cells(f"A{row}:B{row}"); c = ws.cell(row=row, column=1); c.value = "GRAND TOTAL"
    c.fill = GTOT_FILL; c.font = Font(name="Arial", bold=True, size=11)
    c.alignment = C; c.border = _thick()
    for col, val in [(3, round(gp,2)),(4, round(gf,2)),(5, round(gt,2))]:
        c = ws.cell(row=row, column=col); c.value = val; c.fill = GTOT_FILL
        c.font = Font(name="Arial",bold=True,size=11); c.alignment = R
        c.border = _thick(); c.number_format = "#,##0.00"
    ws.row_dimensions[row].height = 22; ws.freeze_panes = "A5"

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


# ─── API routes ───────────────────────────────────────────────────────────────
@app.get("/api/health")
def health():
    return {"status": "ok"}


def _process_pdf(pdf_bytes: bytes, filename: str, api_key: str) -> StreamingResponse:
    """Shared logic for both upload endpoints."""
    try:
        pages = pdf_to_images(pdf_bytes, dpi=150)
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Could not read PDF: {e}")

    client = anthropic.Anthropic(api_key=api_key)
    all_sections: list[dict] = []
    errors: list[str] = []

    for page_num, jpeg_bytes in pages:
        try:
            result = extract_bom_from_page(client, page_num, jpeg_bytes)
            if result.get("has_bom"):
                all_sections.extend(result.get("sections", []))
        except Exception as e:
            errors.append(f"Page {page_num}: {e}")

    if not all_sections:
        raise HTTPException(status_code=422,
            detail=f"No BOM tables found. Errors: {'; '.join(errors) or 'None'}")

    try:
        xlsx_bytes = build_excel(all_sections, filename)
    except Exception as e:
        raise HTTPException(status_code=500,
            detail=f"Excel build failed: {e}\n{traceback.format_exc()}")

    stem = Path(filename).stem
    return StreamingResponse(
        io.BytesIO(xlsx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{stem}_BOM.xlsx"'},
    )


@app.post("/api/extract-bom")
async def extract_bom(
    file: UploadFile = File(...),
    x_api_key: Optional[str] = Header(default=None, alias="X-API-Key"),
):
    api_key = x_api_key or os.getenv("ANTHROPIC_API_KEY", "")
    if not api_key:
        raise HTTPException(status_code=401, detail="Anthropic API key required.")

    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF files are accepted.")

    pdf_bytes = await file.read()
    if len(pdf_bytes) > 50 * 1024 * 1024:
        raise HTTPException(status_code=413, detail="File too large (max 50 MB).")

    return _process_pdf(pdf_bytes, file.filename, api_key)


class Base64Upload(BaseModel):
    """Accepts PDF as base64 string to bypass Vercel's multipart body limit."""
    filename: str
    data: str  # base64-encoded PDF bytes
    api_key: Optional[str] = None


@app.post("/api/extract-bom-b64")
async def extract_bom_b64(
    payload: Base64Upload,
    x_api_key: Optional[str] = Header(default=None, alias="X-API-Key"),
):
    """Fallback endpoint: accepts base64-encoded PDF in JSON body.
    This avoids Vercel's multipart form-data body size limit.
    """
    api_key = payload.api_key or x_api_key or os.getenv("ANTHROPIC_API_KEY", "")
    if not api_key:
        raise HTTPException(status_code=401, detail="Anthropic API key required.")

    if not payload.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF files are accepted.")

    try:
        pdf_bytes = base64.b64decode(payload.data)
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid base64 data.")

    if len(pdf_bytes) > 50 * 1024 * 1024:
        raise HTTPException(status_code=413, detail="File too large (max 50 MB).")

    return _process_pdf(pdf_bytes, payload.filename, api_key)
