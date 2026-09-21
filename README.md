# ANALYST.OS — standalone build

This is your AnalystOS component turned into a normal, runnable web app,
so it works outside the Claude.ai artifact preview — on your own laptop,
a college lab PC, or any server, for your semester project / demo / viva.

## Why the AI feature broke elsewhere

Inside a Claude.ai artifact, `fetch("https://api.anthropic.com/...")` is
silently authenticated for you by the browser sandbox. Anywhere else, that
same call has no API key and fails. This build fixes that by adding a tiny
backend (`server.py`) that holds your own Anthropic API key and forwards
requests to it — the frontend now calls `/api/chat` on the same server
instead of calling Anthropic directly.

Data storage (`window.storage`, used for your saved analysis history) is
also artifact-only, so this build shims it with the browser's `localStorage`
when it detects it's not running inside Claude.ai — that's already handled
in `App.jsx`, no action needed.

## One-time setup

You'll need [Node.js](https://nodejs.org) (v18+) and Python 3.9+ installed.

```bash
# 1. Install frontend dependencies
npm install

# 2. Build the frontend into static files
npm run build

# 3. Install backend dependencies
pip install -r requirements.txt

# 4. Get an API key from https://console.anthropic.com/settings/keys
#    and set it as an environment variable:
export ANTHROPIC_API_KEY="sk-ant-your-key-here"      # macOS/Linux
set ANTHROPIC_API_KEY=sk-ant-your-key-here            # Windows (cmd)
$env:ANTHROPIC_API_KEY="sk-ant-your-key-here"         # Windows (PowerShell)
```

## Run it

```bash
python server.py
```

Then open **http://localhost:5000** — the whole app, including AI Insights
and Ask Your Data, runs from this one server and one command going forward.

## Running it permanently / for a demo day

- **Keep it running locally:** just re-run `python server.py` any time; no
  rebuild needed unless you edit `src/App.jsx` (then re-run `npm run build`).
- **Put it on a small cloud VM** (e.g. a free-tier instance, Render, Railway,
  PythonAnywhere): copy this whole folder up, run the same setup steps, set
  `ANTHROPIC_API_KEY` as an environment variable on that host, and run
  `python server.py` (or point a process manager like `pm2`/`systemd`/
  `gunicorn` at it) so it stays up between reboots.
- **Never commit your API key** to Git or hand in the key itself with your
  project — hand in the code, and mention the key is supplied via
  environment variable in your report/README.

## Project layout

```
analystos/
├── src/
│   ├── App.jsx        # your original component (AI calls now point to /api/chat)
│   └── main.jsx        # React entry point
├── index.html
├── package.json
├── vite.config.js
├── server.py           # Flask backend: serves the app + proxies AI calls
├── requirements.txt
└── README.md
```

## Local development (optional)

If you want hot-reload while editing `App.jsx`, run two terminals:

```bash
# terminal 1
python server.py

# terminal 2
npm run dev
```

Then open the Vite dev URL it prints (usually http://localhost:5173) —
`vite.config.js` is already set up to forward `/api` calls to the Flask
server on port 5000.
