# BOM Extractor – Vercel Deployment (Fixed)

Upload an engineering drawing PDF → Claude reads every BOM table → download formatted Excel.

## What's Fixed in v2.0.1

✅ **Better error messages** — Server logs now show what's going wrong  
✅ **API key input in UI** — Users can enter their own key if not configured server-side  
✅ **Health check endpoint** — `/api/health` to debug configuration issues  
✅ **Improved timeout handling** — Sets explicit `maxDuration: 300` for 5-minute PDF processing  
✅ **CORS properly configured** — No more "cross-origin" errors  
✅ **Better error details** — Clear messages for missing API keys, timeouts, invalid JSON  

---

## Deploy to Vercel (4 steps)

### 1. Prepare your repo

Your GitHub repo should have this exact structure:

```
/
├── main.py           ← FastAPI backend + frontend route
├── index.html        ← Single-page UI
├── requirements.txt  ← Python dependencies
└── vercel.json       ← Vercel config (routing + maxDuration)
```

### 2. Push to GitHub

Commit all files and push to your GitHub repository.

### 3. Import on Vercel

1. Go to [vercel.com/new](https://vercel.com/new)
2. **Click "Import Git Repository"** and select your repo
3. **Framework preset:** Select **Other** (Vercel auto-detects FastAPI)
4. **Environment variables:** Click to expand and add:

| Key | Value |
|---|---|
| `ANTHROPIC_API_KEY` | `sk-ant-v0-...` (your API key from https://console.anthropic.com) |

**⚠️ Important:** If you don't have a Pro plan, the free Hobby plan caps functions at 60 seconds. Set `maxDuration: 300` in `vercel.json` requires **Vercel Pro** ($20/month). For free tier, reduce `maxDuration` to 60 and stick to short PDFs (1-2 pages).

5. **Click Deploy**

### 4. Test your deployment

Your app is live at `https://your-project.vercel.app`

**First thing to test:** 
- Visit `https://your-project.vercel.app/api/health` in your browser
- You should see JSON like:
  ```json
  {
    "status": "ok",
    "api_key_configured": true,
    "version": "2.0.1"
  }
  ```

If `api_key_configured` is `false`, your environment variable wasn't set. Go to **Vercel → Project → Settings → Environment Variables** and add it.

---

## Troubleshooting

### "404 Not Found" error
- **Check:** Is your `vercel.json` file in the root of your repo?
- **Check:** Are `main.py` and `index.html` in the root (not in a subfolder)?

### "401: Anthropic API key required"
- **Problem:** API key not set as environment variable
- **Solution:** 
  1. Go to [Vercel project settings](https://vercel.com/dashboard)
  2. Find your project → **Settings** → **Environment Variables**
  3. Add: `ANTHROPIC_API_KEY` = `sk-ant-v0-...`
  4. Redeploy: Click **Deployments** → **... menu** → **Redeploy**
  5. Or users can paste their API key in the UI (it won't be stored)

### "Request timeout" or "Function invoked 120s"
- **Problem:** PDF processing took too long
- **Hobby plan (free):** Max 60 seconds. Try smaller PDFs (1-3 pages max)
- **Pro plan:** Supports up to 5 minutes. If still timing out, your PDF may be too large/complex
- **Note:** `vercel.json` sets `maxDuration: 300` which requires **Vercel Pro** — change to `60` if on free tier

### "BODY_LIMIT exceeded" or "Request too large"
- **Problem:** Page image exceeded Vercel's 4.5 MB request body limit
- **This is very rare** — the app auto-reduces JPEG quality if too large
- **If it happens:** Your PDF page might be a high-res scan. Try reducing DPI or splitting the PDF

### Server returns HTML error page instead of JSON
- **Problem:** FastAPI crashed or isn't running
- **Debug:** Check Vercel logs:
  1. Go to **Vercel → Deployments**
  2. Click the failed deployment
  3. **Scroll down to logs** and look for Python errors
  4. Share the error message in your issue tracker

### PDF.js not loading (blank page or "PDF.js is not defined")
- **Problem:** CDN blocked or slow to load
- **Solution:** Hard refresh your browser (Ctrl+Shift+R or Cmd+Shift+R)
- **Alternative:** Check browser console (F12 → Console) for CDN errors

### Browser console shows "CORS error" or "Access-Control-Allow-Origin"
- **This should be fixed in v2.0.1** — if you still see it:
  1. Clear browser cache (Ctrl+Shift+Delete)
  2. Hard refresh (Ctrl+Shift+R)
  3. Try a different browser (Chrome, Firefox, Safari)

---

## Environment variables

| Variable | Required | Description |
|---|---|---|
| `ANTHROPIC_API_KEY` | No (optional) | Your Anthropic API key. If not set, users must enter it in the UI. |

Set in **Vercel → Project → Settings → Environment Variables** or the **Environment Variable** section during deployment.

---

## Performance notes

- **Timeout:** `vercel.json` sets `maxDuration: 300` (5 min). **Requires Vercel Pro.**
  - On Hobby plan (free): Change to `maxDuration: 60` for 60 seconds max
  - PDFs under 5 pages process in 20-40 seconds
  - Large scanned PDFs may timeout — tell users to split them

- **File size:** Vercel request body limit is 4.5 MB by default
  - App limits uploads to 50 MB client-side
  - Page images auto-resize if too large (quality reduction)
  - For PDFs over 20 pages, consider deploying locally instead

- **Bundle size:** PyMuPDF (~80 MB unzipped) fits easily within Vercel's 500 MB limit

---

## Local development

```bash
# Install dependencies
pip install -r requirements.txt

# Run FastAPI locally
uvicorn main:app --reload --port 8000

# Open in browser
open http://localhost:8000
```

---

## Debugging steps

If deployment still fails:

1. **Check Vercel logs:**
   - Deployments tab → Click failed deployment → View logs
   - Look for Python import errors or missing packages

2. **Test `/api/health` endpoint:**
   ```bash
   curl https://your-project.vercel.app/api/health
   ```
   Should return JSON with status and version

3. **Check file structure:**
   ```
   ls -la
   # Should show: main.py, index.html, requirements.txt, vercel.json
   ```

4. **Verify requirements.txt has all packages:**
   ```
   fastapi>=0.111.0
   python-multipart>=0.0.9
   anthropic>=0.28.0
   openpyxl>=3.1.2
   PyMuPDF>=1.24.0
   pydantic>=2.0.0
   ```

5. **Check for Python version issues:**
   - vercel.json specifies `"runtime": "python3.11"`
   - If it fails, try removing the runtime line (auto-detected)

---

## API Reference

### `/` (GET)
Serves the frontend HTML.

### `/api/health` (GET)
Returns server status and configuration:
```json
{
  "status": "ok",
  "api_key_configured": true,
  "version": "2.0.1"
}
```

### `/api/extract-page` (POST)
Extracts BOM from a single page image. Called once per page from the browser.

**Request:**
```json
{
  "page_num": 1,
  "image_data": "base64-encoded-jpeg",
  "media_type": "image/jpeg"
}
```

**Response:**
```json
{
  "has_bom": true,
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
  ]
}
```

### `/api/build-excel` (POST)
Builds the final Excel workbook from all BOM sections.

**Request:**
```json
{
  "filename": "drawing.pdf",
  "sections": [
    { "title": "...", "profiles": [...], "fasteners": [...] }
  ]
}
```

**Response:** Binary Excel file (application/vnd.openxmlformats-officedocument.spreadsheetml.sheet)

---

## Support

- **Issues?** Check the troubleshooting section above
- **Anthropic API questions?** See [console.anthropic.com](https://console.anthropic.com)
- **Vercel questions?** See [vercel.com/docs](https://vercel.com/docs)
