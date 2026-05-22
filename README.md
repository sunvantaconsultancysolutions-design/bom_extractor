# BOM Extractor – Vercel Deployment

Upload an engineering drawing PDF → Claude reads every BOM table → download formatted Excel.

---

## Deploy to Vercel (3 steps)

### 1. Push to GitHub

Your repo should have this structure:

```
/
├── main.py           ← FastAPI backend + frontend route
├── index.html        ← Single-page UI
├── requirements.txt  ← Python dependencies
└── vercel.json       ← Vercel config (routing + maxDuration)
```

### 2. Import on Vercel

1. Go to [vercel.com/new](https://vercel.com/new)
2. Import your GitHub repo
3. Framework preset: **Other** (Vercel auto-detects FastAPI)
4. Add environment variable:
   - **Key:** `ANTHROPIC_API_KEY`
   - **Value:** `sk-ant-…`
5. Click **Deploy**

### 3. Done

Your app is live at `https://your-project.vercel.app`

---

## Environment variables

| Variable | Required | Description |
|---|---|---|
| `ANTHROPIC_API_KEY` | Yes (or enter in UI) | Your Anthropic API key |

Set it in **Vercel → Project → Settings → Environment Variables**.

If not set server-side, users can enter their own key in the UI.

---

## Notes

- **Timeout:** `vercel.json` sets `maxDuration: 300` (5 min). Requires **Vercel Pro**. Hobby plan caps at 60s — fine for PDFs up to ~5 pages.
- **File size:** Vercel request body limit is 4.5 MB by default. The app limits uploads to 50 MB; for very large PDFs deploy locally instead.
- **Bundle size:** PyMuPDF is ~80 MB unzipped, well within Vercel's 500 MB limit.

---

## Local development

```bash
pip install -r requirements.txt uvicorn
uvicorn main:app --reload --port 8000
# open http://localhost:8000
```
