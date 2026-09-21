"""
ANALYST.OS backend.

Two jobs:
1. Serve the built frontend (the `dist/` folder produced by `npm run build`).
2. Provide POST /api/chat, which forwards a prompt to the real Anthropic API
   using YOUR OWN api key (kept server-side, never shipped to the browser).

This is what lets the AI Insights / Ask Your Data features work anywhere you
run this app — your laptop, a college lab machine, a cloud VM — not just
inside a Claude.ai artifact preview.

Setup:
    pip install -r requirements.txt
    export ANTHROPIC_API_KEY="sk-ant-..."      (Windows: set ANTHROPIC_API_KEY=...)
    npm install
    npm run build
    python server.py

Then open http://localhost:5000
"""

import os
from pathlib import Path

from flask import Flask, request, jsonify, send_from_directory
import anthropic

DIST_DIR = Path(__file__).parent / "dist"
MODEL = "claude-sonnet-4-6"

app = Flask(__name__, static_folder=str(DIST_DIR), static_url_path="")

api_key = os.environ.get("ANTHROPIC_API_KEY")
client = anthropic.Anthropic(api_key=api_key) if api_key else None


@app.post("/api/chat")
def chat():
    if client is None:
        return jsonify(
            error="Server has no ANTHROPIC_API_KEY set. Set it as an "
            "environment variable and restart server.py."
        ), 500

    data = request.get_json(force=True, silent=True) or {}
    prompt = (data.get("prompt") or "").strip()
    if not prompt:
        return jsonify(error="Missing 'prompt' in request body."), 400

    try:
        resp = client.messages.create(
            model=MODEL,
            max_tokens=1000,
            messages=[{"role": "user", "content": prompt}],
        )
        text = "\n".join(
            block.text for block in resp.content if block.type == "text"
        ).strip()
        return jsonify(text=text)
    except anthropic.APIError as e:
        return jsonify(error=f"Anthropic API error: {e}"), 502
    except Exception as e:  # noqa: BLE001
        return jsonify(error=f"Unexpected server error: {e}"), 500


@app.get("/")
@app.get("/<path:path>")
def serve_frontend(path=""):
    if not DIST_DIR.exists():
        return (
            "Frontend not built yet. Run:\n"
            "  npm install\n"
            "  npm run build\n"
            "then restart this server.",
            200,
        )
    full_path = DIST_DIR / path
    if path and full_path.exists():
        return send_from_directory(DIST_DIR, path)
    return send_from_directory(DIST_DIR, "index.html")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"ANALYST.OS running at http://localhost:{port}")
    if client is None:
        print("WARNING: ANTHROPIC_API_KEY is not set — AI features will fail.")
    app.run(host="0.0.0.0", port=port, debug=False)
