#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""WBS-Outlook 同期ツール Web UI サーバー"""

import json
import re
import sys
import subprocess
import threading
import uuid
from pathlib import Path
from queue import Queue, Empty

from flask import Flask, render_template, request, jsonify, send_file, Response

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

BASE_DIR   = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

PYTHON = sys.executable

# ジョブ管理: {job_id: {"queue": Queue, "done": bool}}
_jobs: dict[str, dict] = {}


# ─── ファイル名バリデーション ────────────────────────────────────────────

def _safe_name(name: str) -> str | None:
    """パストラバーサル防止 + .xlsx のみ許可"""
    name = Path(name).name
    if not name.lower().endswith(".xlsx"):
        return None
    if re.search(r'[<>:"/\\|?*\x00-\x1f]', name):
        return None
    return name


# ─── ルーティング ────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/files", methods=["GET"])
def list_files():
    files = sorted(f.name for f in UPLOAD_DIR.glob("*.xlsx"))
    return jsonify(files)


@app.route("/upload", methods=["POST"])
def upload():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "ファイルが選択されていません"}), 400
    name = _safe_name(f.filename or "")
    if not name:
        return jsonify({"error": ".xlsx ファイルのみアップロードできます"}), 400
    f.save(UPLOAD_DIR / name)
    return jsonify({"message": f"{name} をアップロードしました", "filename": name})


@app.route("/files/<filename>", methods=["GET"])
def download_file(filename):
    name = _safe_name(filename)
    if not name:
        return jsonify({"error": "無効なファイル名です"}), 400
    path = UPLOAD_DIR / name
    if not path.exists():
        return jsonify({"error": "ファイルが見つかりません"}), 404
    return send_file(path, as_attachment=True, download_name=name)


@app.route("/files/<filename>", methods=["DELETE"])
def delete_file(filename):
    name = _safe_name(filename)
    if not name:
        return jsonify({"error": "無効なファイル名です"}), 400
    path = UPLOAD_DIR / name
    if path.exists():
        path.unlink()
    return jsonify({"message": f"{name} を削除しました"})


# ─── コマンド実行 ────────────────────────────────────────────────────────

@app.route("/run", methods=["POST"])
def run_command():
    data     = request.get_json(silent=True) or {}
    command  = data.get("command", "")
    filename = data.get("filename", "")
    ap_mail  = data.get("ap_mail", "").strip()
    g_start  = data.get("gantt_start", "")
    g_end    = data.get("gantt_end", "")

    script = str(BASE_DIR / "wbs_outlook_sync.py")
    cmd    = [PYTHON, "-X", "utf8", script]

    if command in ("wbs2outlook", "outlook2wbs"):
        if not filename or not ap_mail:
            return jsonify({"error": "ファイルと AP_MailAddress を指定してください"}), 400
        name = _safe_name(filename)
        if not name:
            return jsonify({"error": "無効なファイル名です"}), 400
        cmd += [command, "--file", str(UPLOAD_DIR / name), "--id", ap_mail]
        if command == "outlook2wbs":
            cmd += ["--output", str(UPLOAD_DIR / name)]

    elif command in ("add-time-cols", "add-status-col"):
        if not filename:
            return jsonify({"error": "ファイルを指定してください"}), 400
        name = _safe_name(filename)
        if not name:
            return jsonify({"error": "無効なファイル名です"}), 400
        cmd += [command, "--file", str(UPLOAD_DIR / name)]

    elif command == "new-wbs":
        if not filename or not g_start or not g_end:
            return jsonify({"error": "ファイル名・ガント開始日・終了日を指定してください"}), 400
        name = _safe_name(filename if filename.endswith(".xlsx") else filename + ".xlsx")
        if not name:
            return jsonify({"error": "無効なファイル名です"}), 400
        cmd += ["new-wbs", "--output", str(UPLOAD_DIR / name),
                "--gantt-start", g_start, "--gantt-end", g_end]

    else:
        return jsonify({"error": f"不明なコマンド: {command}"}), 400

    # ジョブ起動
    job_id = str(uuid.uuid4())
    q = Queue()
    _jobs[job_id] = {"queue": q, "done": False}

    def _run():
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                cwd=str(BASE_DIR),
            )
            for line in proc.stdout:
                q.put(line.rstrip("\n"))
            proc.wait()
            q.put(f"\n--- 終了コード: {proc.returncode} ---")
        except Exception as e:
            q.put(f"エラー: {e}")
        finally:
            _jobs[job_id]["done"] = True
            q.put(None)  # sentinel

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"job_id": job_id})


# ─── SSE ストリーム ──────────────────────────────────────────────────────

@app.route("/run/stream/<job_id>")
def stream(job_id):
    job = _jobs.get(job_id)
    if not job:
        return jsonify({"error": "ジョブが見つかりません"}), 404

    def _generate():
        q = job["queue"]
        while True:
            try:
                line = q.get(timeout=120)
            except Empty:
                yield "data: [TIMEOUT]\n\n"
                break
            if line is None:
                yield "data: [DONE]\n\n"
                break
            yield f"data: {json.dumps(line, ensure_ascii=False)}\n\n"

    return Response(
        _generate(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


# ─── エントリポイント ────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 50)
    print("  WBS-Outlook 同期ツール")
    print("  http://localhost:5000")
    print("  終了: Ctrl+C")
    print("=" * 50)
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)
