from __future__ import annotations

import cgi
import json
import mimetypes
import re
import shutil
import sys
import traceback
import uuid
from dataclasses import asdict
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

from .agent_run import process_document
from .config import AgentConfig
from .pipeline import _jsonable
from .tools import Toolchain


PROJECT_ROOT = Path(__file__).resolve().parents[2]
RUN_ROOT = PROJECT_ROOT / "runs" / "web-ui"
SUPPORTED_DOC_SUFFIXES = {".doc", ".docx"}


def serve(host: str = "127.0.0.1", port: int = 8765, open_port_search: bool = True) -> int:
    RUN_ROOT.mkdir(parents=True, exist_ok=True)
    selected_port = port
    server: ThreadingHTTPServer | None = None
    while server is None:
        try:
            server = ThreadingHTTPServer((host, selected_port), ThesisWebHandler)
        except OSError:
            if not open_port_search:
                raise
            selected_port += 1
            if selected_port > port + 20:
                raise
    print(f"Thesis Agent web UI: http://{host}:{selected_port}", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping Thesis Agent web UI.", flush=True)
    finally:
        server.server_close()
    return 0


class ThesisWebHandler(BaseHTTPRequestHandler):
    server_version = "ThesisAgentWeb/0.1"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self._send_html(INDEX_HTML)
            return
        if parsed.path.startswith("/download/"):
            self._send_download(parsed.path.removeprefix("/download/"))
            return
        if parsed.path == "/api/health":
            self._send_json({"ok": True})
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path != "/api/chat":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        try:
            payload = self._handle_chat()
        except Exception as exc:  # pragma: no cover - keeps the local web UI debuggable.
            payload = {
                "reply": f"处理失败：{exc}",
                "artifacts": [],
                "error": traceback.format_exc(),
            }
            self._send_json(payload, status=500)
            return
        self._send_json(payload)

    def log_message(self, fmt: str, *args) -> None:
        sys.stderr.write("[%s] %s\n" % (self.log_date_time_string(), fmt % args))

    def _handle_chat(self) -> dict[str, object]:
        content_type = self.headers.get("Content-Type", "")
        if content_type.startswith("multipart/form-data"):
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": content_type,
                },
            )
            message = str(form.getvalue("message", "") or "").strip()
            uploaded_files = _save_uploads(form)
        else:
            length = int(self.headers.get("Content-Length", "0") or "0")
            body = self.rfile.read(length).decode("utf-8") if length else "{}"
            data = json.loads(body or "{}")
            message = str(data.get("message", "") or "").strip()
            uploaded_files = []

        if not uploaded_files:
            return {
                "reply": _ordinary_chat_reply(message),
                "artifacts": [],
            }
        return _process_uploaded_materials(message, uploaded_files)

    def _send_html(self, html: str) -> None:
        body = html.encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, payload: dict[str, object], status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_download(self, encoded_path: str) -> None:
        relative = unquote(encoded_path)
        path = (PROJECT_ROOT / relative).resolve()
        runs_root = (PROJECT_ROOT / "runs").resolve()
        if not _is_relative_to(path, runs_root) or not path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        body = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{_url_quote(path.name)}")
        self.end_headers()
        self.wfile.write(body)


def _save_uploads(form: cgi.FieldStorage) -> list[Path]:
    fields = form["files"] if "files" in form else []
    if not isinstance(fields, list):
        fields = [fields]
    if not fields:
        return []

    upload_dir = RUN_ROOT / datetime.now().strftime("%Y%m%d-%H%M%S") / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)
    saved: list[Path] = []
    for field in fields:
        if not getattr(field, "filename", ""):
            continue
        filename = _safe_filename(Path(field.filename).name)
        if not filename:
            continue
        destination = _dedup_path(upload_dir / filename)
        with destination.open("wb") as fh:
            shutil.copyfileobj(field.file, fh)
        saved.append(destination)
    return saved


def _process_uploaded_materials(message: str, uploaded_files: list[Path]) -> dict[str, object]:
    documents = [path for path in uploaded_files if path.suffix.lower() in SUPPORTED_DOC_SUFFIXES and not path.name.startswith("~$")]
    template = _select_template(documents)
    targets = [path for path in documents if path != template and not _looks_like_template(path)]
    if template is None:
        return {
            "reply": "没有找到可用模板。请保留项目内 samples/templates，或同时上传学校模板。",
            "artifacts": [],
        }
    if not targets:
        return {
            "reply": "已收到文件，但没有识别到需要处理的学生材料。请上传 .doc 或 .docx 毕设文档。",
            "artifacts": [_artifact(template, "模板")],
        }

    config = AgentConfig.load()
    toolchain = Toolchain.discover()
    request_dir = uploaded_files[0].parents[1]
    output_root = request_dir / "outputs"
    artifacts: list[dict[str, str]] = []
    summaries: list[str] = []
    errors: list[str] = []
    for target in targets:
        out_dir = output_root / _slug(target.stem)
        try:
            result = process_document(template, target, out_dir, config, toolchain, build_vision=True)
        except Exception as exc:
            errors.append(f"{target.name}: {exc}")
            continue
        payload = _jsonable(asdict(result))
        (out_dir / "web_result.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        status = "通过" if result.gate_passed else "需要复核"
        summaries.append(
            f"{target.name}：{status}，格式分 {result.format_score}/100，内容分 {result.content_score}/100，"
            f"硬阻塞 {result.hard_blocker_count} 项。"
        )
        artifacts.extend(
            [
                _artifact(result.final_docx, "修改后文档"),
                _artifact(result.revision_checklist, "修改意见"),
                _artifact(result.audit_report, "审计报告"),
                _artifact(out_dir / "process_report.md", "处理摘要"),
                _artifact(out_dir / "web_result.json", "结果数据"),
            ]
        )
    reply_parts = []
    if message:
        reply_parts.append(f"已按你的说明处理：{message}")
    reply_parts.extend(summaries)
    if errors:
        reply_parts.append("失败项：" + "；".join(errors))
    return {
        "reply": "\n".join(reply_parts) if reply_parts else "处理完成。",
        "artifacts": artifacts,
    }


def _ordinary_chat_reply(message: str) -> str:
    text = message.strip()
    if not text:
        return "收到。"
    compact = re.sub(r"\s+", "", text.lower())
    if any(token in compact for token in ("你好", "hello", "hi")):
        return "你好，我在。"
    if any(token in compact for token in ("帮助", "怎么用", "能做什么")):
        return "可以直接聊天，也可以上传 .doc/.docx 毕设材料；上传后会返回修改稿、审计报告和返修清单。"
    if any(token in compact for token in ("格式", "论文", "毕设", "开题", "任务书")):
        return "把相关 Word 材料发上来，我会按项目里的模板路线处理并返回结果。"
    return "收到。我可以继续对话；如果要处理毕设材料，请把 Word 文档一起发来。"


def _select_template(documents: list[Path]) -> Path | None:
    uploaded_templates = [path for path in documents if _looks_like_template(path)]
    if uploaded_templates:
        return uploaded_templates[0]
    candidates = [
        PROJECT_ROOT / "samples" / "templates" / "论文格式.doc",
        PROJECT_ROOT / "samples" / "templates" / "论文格式.docx",
    ]
    template_dir = PROJECT_ROOT / "samples" / "templates"
    if template_dir.exists():
        candidates.extend(
            path
            for path in sorted(template_dir.rglob("*"))
            if path.is_file() and path.suffix.lower() in SUPPORTED_DOC_SUFFIXES and not path.name.startswith("~$")
        )
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def _looks_like_template(path: Path) -> bool:
    compact = re.sub(r"\s+", "", path.stem.lower())
    return any(token in compact for token in ("模板", "格式规范", "论文格式", "附件13", "附件16", "附件17"))


def _artifact(path: Path, kind: str) -> dict[str, str]:
    return {
        "kind": kind,
        "name": path.name,
        "url": "/download/" + _url_quote(str(path.resolve().relative_to(PROJECT_ROOT))),
    }


def _safe_filename(value: str) -> str:
    value = value.replace("\\", "/").split("/")[-1]
    value = re.sub(r"[<>:\"|?*\x00-\x1f]", "", value).strip()
    return value[:160]


def _dedup_path(path: Path) -> Path:
    if not path.exists():
        return path
    for idx in range(1, 1000):
        candidate = path.with_name(f"{path.stem}-{idx}{path.suffix}")
        if not candidate.exists():
            return candidate
    return path.with_name(f"{path.stem}-{uuid.uuid4().hex[:8]}{path.suffix}")


def _slug(value: str) -> str:
    slug = re.sub(r"[^\w\u4e00-\u9fff.-]+", "-", value).strip("-")
    return slug[:80] or uuid.uuid4().hex[:8]


def _url_quote(value: str) -> str:
    from urllib.parse import quote

    return quote(value, safe="/.-_")


def _is_relative_to(path: Path, root: Path) -> bool:
    try:
        path.relative_to(root)
        return True
    except ValueError:
        return False


INDEX_HTML = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Thesis Agent</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #f6f7f4;
      --panel: #ffffff;
      --panel-2: #eef3ed;
      --text: #1e241f;
      --muted: #647066;
      --line: #d9dfd8;
      --accent: #1f7a57;
      --accent-2: #0f5b40;
      --warn: #996b12;
      --danger: #aa3333;
      --shadow: 0 18px 48px rgba(28, 41, 34, .12);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      background: var(--bg);
      color: var(--text);
      font-family: ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      letter-spacing: 0;
    }
    .app {
      height: 100vh;
      display: grid;
      grid-template-columns: 280px minmax(0, 1fr);
    }
    .rail {
      border-right: 1px solid var(--line);
      background: #fbfcfa;
      padding: 16px;
      display: flex;
      flex-direction: column;
      gap: 14px;
      min-width: 0;
    }
    .brand {
      display: flex;
      align-items: center;
      gap: 10px;
      font-weight: 700;
      min-height: 40px;
    }
    .mark {
      width: 32px;
      height: 32px;
      border-radius: 8px;
      display: grid;
      place-items: center;
      color: white;
      background: var(--accent);
      flex: 0 0 auto;
    }
    .side-section {
      border-top: 1px solid var(--line);
      padding-top: 12px;
      min-width: 0;
    }
    .side-title {
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
      text-transform: uppercase;
      margin-bottom: 8px;
    }
    .artifact-list, .file-list {
      display: grid;
      gap: 8px;
    }
    .artifact-link, .file-chip {
      min-height: 36px;
      display: flex;
      align-items: center;
      gap: 8px;
      padding: 8px 10px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--panel);
      color: var(--text);
      text-decoration: none;
      font-size: 13px;
      min-width: 0;
    }
    .artifact-link span, .file-chip span {
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .main {
      min-width: 0;
      display: grid;
      grid-template-rows: auto minmax(0, 1fr) auto;
    }
    .topbar {
      min-height: 58px;
      padding: 12px 20px;
      border-bottom: 1px solid var(--line);
      background: rgba(255,255,255,.82);
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
    }
    .status {
      color: var(--muted);
      font-size: 13px;
      white-space: nowrap;
    }
    .messages {
      overflow: auto;
      padding: 22px 20px;
    }
    .thread {
      max-width: 900px;
      margin: 0 auto;
      display: grid;
      gap: 18px;
    }
    .msg {
      display: grid;
      grid-template-columns: 34px minmax(0, 1fr);
      gap: 12px;
      align-items: start;
    }
    .avatar {
      width: 34px;
      height: 34px;
      border-radius: 8px;
      display: grid;
      place-items: center;
      background: var(--panel-2);
      color: var(--accent-2);
      font-weight: 700;
      flex: 0 0 auto;
    }
    .msg.user .avatar {
      background: #f1e8d2;
      color: var(--warn);
    }
    .bubble {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 13px 14px;
      box-shadow: 0 1px 0 rgba(0,0,0,.02);
      white-space: pre-wrap;
      line-height: 1.55;
      overflow-wrap: anywhere;
    }
    .msg.user .bubble {
      background: #fffaf0;
    }
    .downloads {
      margin-top: 12px;
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
    }
    .download {
      display: inline-flex;
      align-items: center;
      gap: 7px;
      min-height: 34px;
      border-radius: 8px;
      padding: 7px 10px;
      border: 1px solid #bdd7cc;
      background: #f3fbf7;
      color: var(--accent-2);
      text-decoration: none;
      font-size: 13px;
      max-width: 280px;
    }
    .download span {
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .composer-wrap {
      border-top: 1px solid var(--line);
      background: rgba(246,247,244,.92);
      padding: 14px 20px 18px;
    }
    .composer {
      max-width: 900px;
      margin: 0 auto;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      padding: 10px;
    }
    .attachments {
      display: none;
      gap: 8px;
      flex-wrap: wrap;
      padding: 0 0 8px;
    }
    .attachments.active { display: flex; }
    .attach-chip {
      border-radius: 8px;
      border: 1px solid var(--line);
      background: var(--panel-2);
      padding: 6px 9px;
      font-size: 13px;
      max-width: 260px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .input-row {
      display: grid;
      grid-template-columns: 42px minmax(0, 1fr) 42px;
      gap: 8px;
      align-items: end;
    }
    textarea {
      width: 100%;
      min-height: 42px;
      max-height: 180px;
      resize: vertical;
      border: 0;
      outline: 0;
      padding: 10px 8px;
      font: inherit;
      line-height: 1.45;
      color: var(--text);
      background: transparent;
    }
    .icon-btn {
      width: 42px;
      height: 42px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #f9faf8;
      color: var(--text);
      display: grid;
      place-items: center;
      cursor: pointer;
    }
    .icon-btn:hover { border-color: #b8c8bd; background: #f4f7f3; }
    .send {
      background: var(--accent);
      color: white;
      border-color: var(--accent);
    }
    .send:hover { background: var(--accent-2); }
    .send:disabled, .icon-btn:disabled {
      opacity: .55;
      cursor: not-allowed;
    }
    .spinner {
      width: 18px;
      height: 18px;
      border: 2px solid rgba(31,122,87,.22);
      border-top-color: var(--accent);
      border-radius: 50%;
      animation: spin .8s linear infinite;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    input[type=file] { display: none; }
    @media (max-width: 760px) {
      .app { grid-template-columns: 1fr; }
      .rail { display: none; }
      .topbar { padding: 10px 14px; }
      .messages { padding: 16px 12px; }
      .composer-wrap { padding: 10px 12px 14px; }
      .msg { grid-template-columns: 30px minmax(0, 1fr); }
      .avatar { width: 30px; height: 30px; }
    }
  </style>
</head>
<body>
  <div class="app">
    <aside class="rail">
      <div class="brand">
        <div class="mark">T</div>
        <div>Thesis Agent</div>
      </div>
      <div class="side-section">
        <div class="side-title">文件</div>
        <div id="sideFiles" class="file-list"></div>
      </div>
      <div class="side-section">
        <div class="side-title">产物</div>
        <div id="sideArtifacts" class="artifact-list"></div>
      </div>
    </aside>
    <main class="main">
      <header class="topbar">
        <strong>对话</strong>
        <div id="status" class="status">就绪</div>
      </header>
      <section class="messages" id="messages">
        <div class="thread" id="thread">
          <div class="msg assistant">
            <div class="avatar">A</div>
            <div class="bubble">可以开始。</div>
          </div>
        </div>
      </section>
      <footer class="composer-wrap">
        <form class="composer" id="chatForm">
          <div id="attachments" class="attachments"></div>
          <div class="input-row">
            <button class="icon-btn" type="button" id="pickFiles" title="上传" aria-label="上传">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                <path d="M12 5v10m0-10 4 4m-4-4-4 4" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                <path d="M5 15v2a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-2" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
              </svg>
            </button>
            <textarea id="message" rows="1"></textarea>
            <button class="icon-btn send" type="submit" id="sendBtn" title="发送" aria-label="发送">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                <path d="m5 12 14-7-7 14-2-5-5-2Z" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/>
              </svg>
            </button>
          </div>
          <input id="fileInput" type="file" multiple accept=".doc,.docx,.pdf,.zip,.rar" />
        </form>
      </footer>
    </main>
  </div>
  <script>
    const form = document.getElementById('chatForm');
    const fileInput = document.getElementById('fileInput');
    const pickFiles = document.getElementById('pickFiles');
    const message = document.getElementById('message');
    const thread = document.getElementById('thread');
    const messages = document.getElementById('messages');
    const attachments = document.getElementById('attachments');
    const statusEl = document.getElementById('status');
    const sendBtn = document.getElementById('sendBtn');
    const sideFiles = document.getElementById('sideFiles');
    const sideArtifacts = document.getElementById('sideArtifacts');
    let selectedFiles = [];

    pickFiles.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', () => {
      selectedFiles = Array.from(fileInput.files || []);
      renderAttachments();
    });
    message.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        form.requestSubmit();
      }
    });
    form.addEventListener('submit', async (event) => {
      event.preventDefault();
      const text = message.value.trim();
      if (!text && selectedFiles.length === 0) return;
      appendMessage('user', text || selectedFiles.map(file => file.name).join('\\n'));
      const formData = new FormData();
      formData.append('message', text);
      for (const file of selectedFiles) formData.append('files', file);
      setBusy(true);
      try {
        const response = await fetch('/api/chat', { method: 'POST', body: formData });
        const data = await response.json();
        appendMessage('assistant', data.reply || '完成。', data.artifacts || []);
        if (data.artifacts) renderArtifacts(data.artifacts);
      } catch (error) {
        appendMessage('assistant', '请求失败：' + error);
      } finally {
        message.value = '';
        selectedFiles = [];
        fileInput.value = '';
        renderAttachments();
        setBusy(false);
      }
    });

    function setBusy(busy) {
      sendBtn.disabled = busy;
      pickFiles.disabled = busy;
      statusEl.innerHTML = busy ? '<span class="spinner"></span>' : '就绪';
    }
    function renderAttachments() {
      attachments.classList.toggle('active', selectedFiles.length > 0);
      attachments.innerHTML = selectedFiles.map(file => `<div class="attach-chip">${escapeHtml(file.name)}</div>`).join('');
      sideFiles.innerHTML = selectedFiles.map(file => `<div class="file-chip"><span>${escapeHtml(file.name)}</span></div>`).join('');
    }
    function renderArtifacts(artifacts) {
      sideArtifacts.innerHTML = artifacts.map(a => `<a class="artifact-link" href="${a.url}"><span>${escapeHtml(a.kind)} · ${escapeHtml(a.name)}</span></a>`).join('');
    }
    function appendMessage(role, text, artifacts = []) {
      const node = document.createElement('div');
      node.className = `msg ${role}`;
      const initial = role === 'user' ? 'U' : 'A';
      const links = artifacts.length ? `<div class="downloads">${artifacts.map(a => `
        <a class="download" href="${a.url}">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true">
            <path d="M12 4v10m0 0 4-4m-4 4-4-4" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
            <path d="M5 20h14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
          </svg>
          <span>${escapeHtml(a.kind)} · ${escapeHtml(a.name)}</span>
        </a>`).join('')}</div>` : '';
      node.innerHTML = `<div class="avatar">${initial}</div><div class="bubble">${escapeHtml(text)}${links}</div>`;
      thread.appendChild(node);
      messages.scrollTop = messages.scrollHeight;
    }
    function escapeHtml(value) {
      return String(value || '').replace(/[&<>"']/g, ch => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
      }[ch]));
    }
  </script>
</body>
</html>
"""


def main(argv: list[str] | None = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(prog="thesis-agent-web")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args(argv)
    return serve(args.host, args.port)


if __name__ == "__main__":
    raise SystemExit(main())
