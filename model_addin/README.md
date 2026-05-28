# MODEL() Excel Online Add-in

## Repo Structure

```
your-repo/                  ← GitHub repo root
├── server.py               ← Flask backend
├── requirements.txt        ← Python dependencies  
├── Procfile                ← Render start command
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
```bash
git init
git add .
git commit -m "initial"
git remote add origin https://github.com/YOUR-USERNAME/YOUR-REPO.git
git push -u origin main
```

### Step 2 — Deploy to Render (manual config — important)

1. Go to **render.com** → New → **Web Service**
2. Connect GitHub → select your repo
3. Fill in these fields **exactly**:

| Field | Value |
|---|---|
| **Root Directory** | *(leave completely blank)* |
| **Runtime** | Python 3 |
| **Build Command** | `pip install -r requirements.txt` |
| **Start Command** | `gunicorn server:app --workers 2 --timeout 60` |

4. Click **Deploy**
5. Test: visit `https://your-app.onrender.com/health` → `{"status":"ok"}`

### Step 3 — Update your Render URL in functions.js

Edit `addin/functions.js` line 4:
```js
const API = "https://your-actual-app.onrender.com";
```
Commit and push.

### Step 4 — Enable GitHub Pages

1. Repo → **Settings → Pages**
2. Branch: `main`, Folder: `/addin` → Save
3. Your URL: `https://YOUR-USERNAME.github.io/YOUR-REPO/`

### Step 5 — Edit manifest.xml and upload to Excel Online

Replace in `addin/manifest.xml`:
- `YOUR-GITHUB-USERNAME` → your GitHub username
- `YOUR-REPO-NAME` → your repo name
- `YOUR-APP-NAME` → your Render app name

In Excel Online: **Insert → Office Add-ins → Upload My Add-in** → pick `manifest.xml`

---

## Usage

```excel
=MODEL("AAPL", "Revenue", "TTM")
=MODEL("AAPL", "EPS", "LQ")
=MODEL("TSLA", "Free Cash Flow", "LY-1",, 1000000)
=MODEL("AAPL", "PE Ratio", "TTM")

=MODELPRICE("AAPL", "Price")
=MODELPRICE("AAPL", "Close", 30)
=MODELPRICE("BTCUSD", "Price")
=MODELPRICE("EURUSD", "Price")
=MODELPRICE("^GSPC", "Close", 10)

=MODELFUNDS("SPY", "Expense Ratio")
=MODELFUNDS("QQQ", "AUM")

=MODELOPTIONS("AAPL", "Call", "Strike")
```

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `could not open requirements.txt` | Root Directory in Render must be **blank** (not "backend" or ".") |
| `#ERROR: Cannot reach API` | Update `functions.js` line 4 with your real Render URL |
| Render slow first call | Free tier sleeps — ping `/health` every 14 min with UptimeRobot |
| GitHub Pages 404 | Pages folder must be `/addin`, not `/` |
| Manifest rejected | Replace ALL `YOUR-*` placeholders in `manifest.xml` |
