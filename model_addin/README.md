# MODEL() Excel Online Add-in

## Repo Structure (important — keep this exact layout)

```
your-repo/                  ← GitHub repo root = Render root
├── server.py               ← Flask backend
├── requirements.txt        ← Python dependencies
├── render.yaml             ← Render config
├── Procfile                ← Render start command (backup)
├── addin/                  ← GitHub Pages serves this folder
│   ├── functions.js
│   ├── functions.json
│   ├── functions.html
│   └── manifest.xml
└── README.md
```

---

## Deploy in 5 Steps

### Step 1 — Push to GitHub
Create a new repo on github.com, then push this whole folder:
```bash
git init
git add .
git commit -m "initial"
git remote add origin https://github.com/YOUR-USERNAME/YOUR-REPO.git
git push -u origin main
```

### Step 2 — Deploy to Render

1. Go to **render.com** → New → **Web Service**
2. Connect GitHub → select your repo
3. Render will auto-detect `render.yaml` — just click **Deploy**
4. Leave **Root Directory blank** (empty) — the `server.py` is at root
5. Your API URL: `https://model-addin-api.onrender.com`
6. Test it: visit `https://your-app.onrender.com/health` → should show `{"status":"ok"}`

### Step 3 — Update your Render URL

Edit `addin/functions.js` line 4:
```js
const API = "https://your-actual-app.onrender.com";
```
Commit and push.

### Step 4 — Enable GitHub Pages

1. Repo → **Settings → Pages**
2. Source: **Deploy from a branch**
3. Branch: `main`, Folder: `/addin`
4. Save → your add-in URL: `https://YOUR-USERNAME.github.io/YOUR-REPO/`

### Step 5 — Update manifest.xml then upload to Excel Online

Edit `addin/manifest.xml` — replace every placeholder:
- `YOUR-GITHUB-USERNAME` → your GitHub username
- `YOUR-REPO-NAME` → your repo name  
- `YOUR-APP-NAME` → your Render app name

Then in **Excel Online**:
1. Insert → Office Add-ins → **Upload My Add-in**
2. Upload `addin/manifest.xml`

---

## Use it

```excel
=MODEL("AAPL", "Revenue", "TTM")
=MODEL("AAPL", "EPS", "LQ")
=MODEL("TSLA", "Free Cash Flow", "LY-1",, 1000000)
=MODEL("AAPL", "PE Ratio", "TTM")
=MODEL("AAPL", "Sector", "LY")

=MODELPRICE("AAPL", "Price")
=MODELPRICE("AAPL", "Close", 30)
=MODELPRICE("BTCUSD", "Price")
=MODELPRICE("EURUSD", "Price")
=MODELPRICE("^GSPC", "Close", 10)

=MODELFUNDS("SPY", "Expense Ratio")
=MODELFUNDS("QQQ", "AUM")

=MODELOPTIONS("AAPL", "Call", "Strike")
=MODELOPTIONS("AAPL", "Put", "IV", "2025-06-20")
```

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Render "root directory doesn't exist" | Leave Root Directory **blank** in Render dashboard |
| `#ERROR: Cannot reach API` | Check `functions.js` line 4 has your real Render URL |
| Render slow first call | Free tier sleeps — use UptimeRobot to ping `/health` every 14 min |
| GitHub Pages 404 | Confirm Pages is set to `/addin` folder, not `/` root |
| Manifest rejected | Replace ALL `YOUR-*` placeholders in `manifest.xml` |
