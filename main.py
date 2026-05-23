import asyncio
import base64
import io
import json
import os
import sys
import traceback
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import openpyxl
from anthropic import Anthropic
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, Response, StreamingResponse
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pydantic import BaseModel

print(f"[STARTUP] Python {sys.version}", file=sys.stderr)

API_KEY = os.getenv("ANTHROPIC_API_KEY", "").strip()

print(f"[STARTUP] API key configured: {bool(API_KEY)}", file=sys.stderr)

if not API_KEY:
    print("[ERROR] Missing ANTHROPIC_API_KEY", file=sys.stderr)

# Use sync client — more reliable in Vercel Lambda / serverless environments
client = Anthropic(
    api_key=API_KEY,
    timeout=55.0,
    max_retries=1,
)

# Thread pool to run sync Anthropic calls without blocking the event loop
_executor = ThreadPoolExecutor(max_workers=4)

app = FastAPI(title="BOM Extractor", version="4.0.0")

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
# health — includes connectivity check
# ---------------------------------------------------------------------------

@app.get("/api/health", include_in_schema=False)
def health():
    import httpx
    reachable = False
    reach_detail = ""
    try:
        r = httpx.get("https://api.anthropic.com", timeout=5)
        reachable = True
        reach_detail = f"HTTP {r.status_code}"
    except Exception as e:
        reach_detail = str(e)

    return {
        "status": "ok",
        "version": "4.0.0",
        "api_key_configured": bool(API_KEY),
        "anthropic_reachable": reachable,
        "anthropic_reach_detail": reach_detail,
    }


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
# Claude prompt
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """
You are an expert engineering drawing parser specialising in steel tower Bills of Materials (BOM).

Given a page image from an engineering drawing, extract ALL BOM tables visible on the page.

Return ONLY valid JSON — no markdown fences, no explanation, no preamble.

Format:
{
  "sections": [
    {
      "title": "BILL OF MATERIALS",
      "unit": "QTY PER TOWER",
      "profiles": [
        {
          "mark": "100H",
          "qty": 2,
          "profile": "L 90x6",
          "length": 4981,
          "black_weight": 82.66
        }
      ],
      "fasteners": [
        {
          "description": "M16x40",
          "qty": 184,
          "black_weight": 28.46
        }
      ],
      "total_profiles_weight": 5100.8,
      "total_fasteners_weight": 210.83,
      "drawing_total_weight": 5311.62
    }
  ],
  "has_bom": true
}

If no BOM exists on the page:
{
  "has_bom": false,
  "sections": []
}

Rules:
- Extract ALL rows without omission
- Preserve exact text values
- length must be an integer (mm)
- black_weight must be a float (kg)
- Use null for any value that is unclear or missing
- Return ONLY the JSON object
"""


# ---------------------------------------------------------------------------
# extraction — sync call in thread pool
# ---------------------------------------------------------------------------

async def extract_bom_from_image(
    page_num: int,
    image_b64: str,
    media_type: str = "image/jpeg",
):
    def _sync_call():
        print(f"[PAGE {page_num}] sending to Claude …", file=sys.stderr)

        msg = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": image_b64,
                            },
                        },
                        {
                            "type": "text",
                            "text": f"Extract all BOM tables from page {page_num}. Return only valid JSON.",
                        },
                    ],
                }
            ],
        )

        print(f"[PAGE {page_num}] Claude responded OK", file=sys.stderr)
        return msg.content[0].text.strip()

    loop = asyncio.get_event_loop()
    raw = await loop.run_in_executor(_executor, _sync_call)

    # Strip accidental markdown fences just in case
    if raw.startswith("```"):
        parts = raw.split("```")
        raw = parts[1] if len(parts) > 1 else raw
        if raw.startswith("json"):
            raw = raw[4:].lstrip()

    return json.loads(raw)


# ---------------------------------------------------------------------------
# Excel styling helpers
# ---------------------------------------------------------------------------

def _thin_border():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)


HDR_FILL = PatternFill("solid", start_color="1F4E79")
SUB_FILL  = PatternFill("solid", start_color="2E75B6")
ALT_FILL  = PatternFill("solid", start_color="EBF3FB")

CENTER = Alignment(horizontal="center", vertical="center")
LEFT   = Alignment(horizontal="left",   vertical="center")
RIGHT  = Alignment(horizontal="right",  vertical="center")


def _set_cell(cell, value, fill=None, align=None, bold=False):
    cell.value = value
    if fill:
        cell.fill = fill
    cell.border = _thin_border()
    cell.font = Font(
        name="Arial",
        size=9,
        bold=bold,
        color="FFFFFF" if fill in (HDR_FILL, SUB_FILL) else "000000",
    )
    cell.alignment = align or LEFT


# ---------------------------------------------------------------------------
# Excel builder
# ---------------------------------------------------------------------------

def build_excel(sections: list, filename: str) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BOM"

    row = 1

    for section in sections:
        # Section title
        ws.merge_cells(f"A{row}:G{row}")
        _set_cell(ws.cell(row=row, column=1), section.get("title", "BILL OF MATERIALS"), HDR_FILL, CENTER, True)
        row += 1

        # Unit row
        unit = section.get("unit", "")
        if unit:
            ws.merge_cells(f"A{row}:G{row}")
            _set_cell(ws.cell(row=row, column=1), f"Unit: {unit}", SUB_FILL, CENTER, False)
            row += 1

        # ── Profiles ──────────────────────────────────────────────
        profiles = section.get("profiles", [])
        if profiles:
            ws.merge_cells(f"A{row}:G{row}")
            _set_cell(ws.cell(row=row, column=1), "STRUCTURAL PROFILES", SUB_FILL, CENTER, True)
            row += 1

            for col, h in enumerate(["MARK", "QTY", "PROFILE", "LENGTH (mm)", "WEIGHT (kg)"], 1):
                _set_cell(ws.cell(row=row, column=col), h, SUB_FILL, CENTER, True)
            row += 1

            for i, p in enumerate(profiles):
                fill = ALT_FILL if i % 2 == 0 else None
                for col, val in enumerate([
                    p.get("mark"),
                    p.get("qty"),
                    p.get("profile"),
                    p.get("length"),
                    p.get("black_weight"),
                ], 1):
                    _set_cell(ws.cell(row=row, column=col), val, fill)
                row += 1

            # Profiles total
            tp = section.get("total_profiles_weight")
            if tp is not None:
                _set_cell(ws.cell(row=row, column=4), "Total Profiles Weight:", None, RIGHT, True)
                _set_cell(ws.cell(row=row, column=5), tp, None, RIGHT, True)
                row += 1

        # ── Fasteners ─────────────────────────────────────────────
        fasteners = section.get("fasteners", [])
        if fasteners:
            row += 1  # blank spacer
            ws.merge_cells(f"A{row}:G{row}")
            _set_cell(ws.cell(row=row, column=1), "FASTENERS", SUB_FILL, CENTER, True)
            row += 1

            for col, h in enumerate(["DESCRIPTION", "QTY", "WEIGHT (kg)"], 1):
                _set_cell(ws.cell(row=row, column=col), h, SUB_FILL, CENTER, True)
            row += 1

            for i, f in enumerate(fasteners):
                fill = ALT_FILL if i % 2 == 0 else None
                for col, val in enumerate([
                    f.get("description"),
                    f.get("qty"),
                    f.get("black_weight"),
                ], 1):
                    _set_cell(ws.cell(row=row, column=col), val, fill)
                row += 1

            tf = section.get("total_fasteners_weight")
            if tf is not None:
                _set_cell(ws.cell(row=row, column=2), "Total Fasteners Weight:", None, RIGHT, True)
                _set_cell(ws.cell(row=row, column=3), tf, None, RIGHT, True)
                row += 1

        # ── Drawing total ──────────────────────────────────────────
        dt = section.get("drawing_total_weight")
        if dt is not None:
            row += 1
            ws.merge_cells(f"A{row}:D{row}")
            _set_cell(ws.cell(row=row, column=1), "DRAWING TOTAL WEIGHT (kg):", HDR_FILL, RIGHT, True)
            _set_cell(ws.cell(row=row, column=5), dt, HDR_FILL, CENTER, True)
            row += 1

        row += 2  # blank gap between sections

    # Column widths
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 5
    ws.column_dimensions["G"].width = 5

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Request models
# ---------------------------------------------------------------------------

class PageExtractRequest(BaseModel):
    page_num: int
    image_data: str
    media_type: str = "image/jpeg"


class BuildExcelRequest(BaseModel):
    filename: str
    sections: list


# ---------------------------------------------------------------------------
# /api/extract-page
# ---------------------------------------------------------------------------

@app.post("/api/extract-page")
async def extract_page(payload: PageExtractRequest):
    if not API_KEY:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY not configured")

    try:
        result = await extract_bom_from_image(
            payload.page_num,
            payload.image_data,
            payload.media_type,
        )
        return result

    except json.JSONDecodeError as e:
        print(f"[JSON ERROR] {e}", file=sys.stderr)
        raise HTTPException(status_code=422, detail="Claude returned invalid JSON")

    except Exception as e:
        print(traceback.format_exc(), file=sys.stderr)
        raise HTTPException(status_code=500, detail=f"Extraction failed: {str(e)}")


# ---------------------------------------------------------------------------
# /api/build-excel
# ---------------------------------------------------------------------------

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
