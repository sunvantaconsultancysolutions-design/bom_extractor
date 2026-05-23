

import base64
import io
import json
import os
import sys
import traceback
from pathlib import Path

import openpyxl
from anthropic import AsyncAnthropic
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

client = AsyncAnthropic(
    api_key=API_KEY,
    timeout=60.0,
    max_retries=2,
)

app = FastAPI(title="BOM Extractor", version="3.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# -------------------------------------------------------------------
# favicon
# -------------------------------------------------------------------

_FAVICON_ICO = base64.b64decode(
    "AAABAAEAAQEAAAEAGAAwAAAAFgAAACgAAAABAAAAAgAAAAEAGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP8AAAA="
)


@app.get("/favicon.ico", include_in_schema=False)
def favicon():
    return Response(content=_FAVICON_ICO, media_type="image/x-icon")


# -------------------------------------------------------------------
# health
# -------------------------------------------------------------------

@app.get("/api/health", include_in_schema=False)
def health():
    return {
        "status": "ok",
        "api_key_configured": bool(API_KEY),
        "version": "3.0.0"
    }


# -------------------------------------------------------------------
# frontend
# -------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse, include_in_schema=False)
def root():
    html_path = Path(__file__).parent / "index.html"

    if not html_path.exists():
        return "<h1>index.html not found</h1>"

    return html_path.read_text(encoding="utf-8")


# -------------------------------------------------------------------
# Claude prompt
# -------------------------------------------------------------------

SYSTEM_PROMPT = """
You are an expert engineering drawing parser specialising in steel tower Bills of Materials (BOM).

Given a page image from an engineering drawing, extract ALL BOM tables visible on the page.

Return ONLY valid JSON.

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

If no BOM exists:
{
  "has_bom": false,
  "sections": []
}

Rules:
- Extract ALL rows
- Preserve exact text
- length = integer mm
- black_weight = float kg
- unclear values => null
"""

# -------------------------------------------------------------------
# extraction
# -------------------------------------------------------------------


async def extract_bom_from_image(
    page_num: int,
    image_b64: str,
    media_type: str = "image/jpeg"
):
    print(f"[PAGE {page_num}] sending to Claude", file=sys.stderr)

    msg = await client.messages.create(
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
                            "data": image_b64
                        }
                    },
                    {
                        "type": "text",
                        "text": f"Extract all BOM tables from page {page_num}"
                    }
                ]
            }
        ]
    )

    raw = msg.content[0].text.strip()

    if raw.startswith("```"):
        raw = raw.split("```")[1]

        if raw.startswith("json"):
            raw = raw[4:].lstrip()

    return json.loads(raw)


# -------------------------------------------------------------------
# Excel styling
# -------------------------------------------------------------------

def thin():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)


HDR_FILL = PatternFill("solid", start_color="1F4E79")
SUB_FILL = PatternFill("solid", start_color="2E75B6")
ALT_FILL = PatternFill("solid", start_color="EBF3FB")

CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")


def set_cell(cell, value, fill=None, align=None, bold=False):
    cell.value = value

    if fill:
        cell.fill = fill

    cell.border = thin()

    cell.font = Font(
        name="Arial",
        size=9,
        bold=bold,
        color="FFFFFF" if fill in [HDR_FILL, SUB_FILL] else "000000"
    )

    cell.alignment = align or LEFT


# -------------------------------------------------------------------
# Excel build
# -------------------------------------------------------------------

def build_excel(sections, filename):
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "BOM"

    row = 1

    for section in sections:
        ws.merge_cells(f"A{row}:E{row}")

        set_cell(
            ws.cell(row=row, column=1),
            section.get("title", "BOM"),
            HDR_FILL,
            CENTER,
            True
        )

        row += 1

        headers = [
            "MARK",
            "QTY",
            "PROFILE",
            "LENGTH",
            "WEIGHT"
        ]

        for col, h in enumerate(headers, 1):
            set_cell(
                ws.cell(row=row, column=col),
                h,
                SUB_FILL,
                CENTER,
                True
            )

        row += 1

        for i, p in enumerate(section.get("profiles", [])):
            fill = ALT_FILL if i % 2 == 0 else None

            values = [
                p.get("mark"),
                p.get("qty"),
                p.get("profile"),
                p.get("length"),
                p.get("black_weight")
            ]

            for col, val in enumerate(values, 1):
                set_cell(
                    ws.cell(row=row, column=col),
                    val,
                    fill
                )

            row += 1

        row += 2

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 15

    buf = io.BytesIO()

    wb.save(buf)

    return buf.getvalue()


# -------------------------------------------------------------------
# API models
# -------------------------------------------------------------------

class PageExtractRequest(BaseModel):
    page_num: int
    image_data: str
    media_type: str = "image/jpeg"


class BuildExcelRequest(BaseModel):
    filename: str
    sections: list


# -------------------------------------------------------------------
# extract route
# -------------------------------------------------------------------

@app.post("/api/extract-page")
async def extract_page(payload: PageExtractRequest):

    if not API_KEY:
        raise HTTPException(
            status_code=500,
            detail="ANTHROPIC_API_KEY not configured"
        )

    try:
        result = await extract_bom_from_image(
            payload.page_num,
            payload.image_data,
            payload.media_type
        )

        return result

    except json.JSONDecodeError:
        raise HTTPException(
            status_code=422,
            detail="Claude returned invalid JSON"
        )

    except Exception as e:
        print(traceback.format_exc(), file=sys.stderr)

        raise HTTPException(
            status_code=500,
            detail=f"Extraction failed: {str(e)}"
        )


# -------------------------------------------------------------------
# excel route
# -------------------------------------------------------------------

@app.post("/api/build-excel")
async def build_excel_endpoint(payload: BuildExcelRequest):

    if not payload.sections:
        raise HTTPException(
            status_code=422,
            detail="No BOM sections found"
        )

    try:
        excel_bytes = build_excel(
            payload.sections,
            payload.filename
        )

        stem = Path(payload.filename).stem

        return StreamingResponse(
            io.BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition":
                f'attachment; filename="{stem}_BOM.xlsx"'
            }
        )

    except Exception as e:
        print(traceback.format_exc(), file=sys.stderr)

        raise HTTPException(
            status_code=500,
            detail=f"Excel build failed: {str(e)}"
        )
