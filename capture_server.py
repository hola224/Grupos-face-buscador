#!/usr/bin/env python3
"""
Servidor local para capturar grupos visibles desde el navegador.

Corre en 127.0.0.1, entrega un bookmarklet y recibe la captura para actualizar
facebook_groups.json + facebook_groups.xlsx.
"""

from __future__ import annotations

import argparse
import html
import json
import secrets
from dataclasses import dataclass
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import quote, urlparse, parse_qs

from facebook_group_crm import (
    DEFAULT_DATA_FILE,
    DEFAULT_EXCEL_FILE,
    DEFAULT_MIN_MEMBERS,
    add_or_update,
    build_record,
    empty_counters,
    export_excel,
    format_counters,
    load_database,
    save_database,
)


DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8765
DEFAULT_SCROLLS = 12
DEFAULT_SCROLL_DELAY_MS = 1200


@dataclass
class ServerConfig:
    host: str
    port: int
    data_path: Path
    excel_path: Path
    min_members: int
    query: str
    location: str
    category: str
    token: str
    scrolls: int
    delay_ms: int

    @property
    def endpoint(self) -> str:
        return f"http://{self.host}:{self.port}/capture"

    @property
    def receiver_url(self) -> str:
        return f"http://{self.host}:{self.port}/receiver?token={self.token}"

    @property
    def home_url(self) -> str:
        return f"http://{self.host}:{self.port}/"


def load_bookmarklet_source(config: ServerConfig) -> str:
    source_path = Path(__file__).with_name("browser_bookmarklet_capture.js")
    source = source_path.read_text(encoding="utf-8")
    return (
        source.replace("__CAPTURE_ENDPOINT__", config.endpoint)
        .replace("__RECEIVER_URL__", config.receiver_url)
        .replace("__CAPTURE_TOKEN__", config.token)
        .replace("__MAX_SCROLLS__", str(config.scrolls))
        .replace("__SCROLL_DELAY_MS__", str(config.delay_ms))
    )


def bookmarklet_href(config: ServerConfig) -> str:
    source = load_bookmarklet_source(config)
    return "javascript:" + quote(source, safe="()[]{}!~*.-_=+;:,/?&|")


def normalize_group_payload(group: dict[str, Any]) -> dict[str, str]:
    return {
        "name": str(group.get("name") or group.get("n") or "").strip(),
        "url": str(group.get("url") or group.get("u") or "").strip(),
        "members": str(group.get("members") or group.get("m") or "").strip(),
        "raw_text": str(group.get("rawText") or group.get("raw_text") or group.get("r") or "").strip(),
    }


def process_payload(payload: dict[str, Any], config: ServerConfig) -> dict[str, Any]:
    data = load_database(config.data_path)
    query = config.query or str(payload.get("query") or payload.get("q") or "").strip()
    location = config.location
    category = config.category
    groups = payload.get("groups") or payload.get("g") or []

    counters = empty_counters()
    processed = 0
    for raw_group in groups:
        if not isinstance(raw_group, dict):
            counters["ignored"] += 1
            continue

        group = normalize_group_payload(raw_group)
        record = build_record(
            name=group["name"],
            url=group["url"],
            members=group["members"] or group["raw_text"],
            query=query,
            location=location,
            category=category,
            notes="captura navegador",
            raw_text=group["raw_text"],
        )
        result = add_or_update(data, record, config.min_members)
        counters[result] += 1
        processed += 1

    save_database(config.data_path, data)
    export_excel(data, config.excel_path)

    return {
        "ok": True,
        "processed": processed,
        "query": query,
        "location": location,
        "category": category,
        "counters": counters,
        "excel": str(config.excel_path.resolve()),
        "data": str(config.data_path.resolve()),
    }


def result_html(result: dict[str, Any], config: ServerConfig) -> str:
    counters = result["counters"]
    summary = format_counters("Captura lista", counters, config.min_members)
    details = [
        ("Procesados", result["processed"]),
        ("Consulta", result.get("query") or "-"),
        ("Ubicación", result.get("location") or "-"),
        ("Categoría", result.get("category") or "-"),
        ("Excel", result["excel"]),
    ]
    rows = "\n".join(
        f"<tr><th>{html.escape(str(label))}</th><td>{html.escape(str(value))}</td></tr>" for label, value in details
    )
    return f"""<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Captura guardada</title>
  <style>
    body {{ font-family: system-ui, sans-serif; margin: 28px; color: #17202a; }}
    h1 {{ font-size: 22px; margin: 0 0 10px; }}
    p {{ line-height: 1.45; }}
    table {{ border-collapse: collapse; margin-top: 16px; width: 100%; }}
    th, td {{ border-bottom: 1px solid #dde3ea; padding: 8px; text-align: left; vertical-align: top; }}
    th {{ width: 130px; color: #546371; }}
    code {{ background: #eef2f6; padding: 2px 5px; border-radius: 4px; }}
  </style>
</head>
<body>
  <h1>Captura guardada</h1>
  <p>{html.escape(summary)}</p>
  <table>{rows}</table>
  <p>Ya puedes volver a Facebook, desplazarte a más resultados y hacer clic otra vez en el bookmarklet.</p>
</body>
</html>"""


def receiver_html(config: ServerConfig) -> str:
    return f"""<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Recibiendo captura</title>
  <style>
    body {{ font-family: system-ui, sans-serif; margin: 28px; color: #17202a; }}
    h1 {{ font-size: 22px; margin: 0 0 10px; }}
    p {{ line-height: 1.45; }}
    table {{ border-collapse: collapse; margin-top: 16px; width: 100%; }}
    th, td {{ border-bottom: 1px solid #dde3ea; padding: 8px; text-align: left; vertical-align: top; }}
    th {{ width: 150px; color: #546371; }}
    code {{ background: #eef2f6; padding: 2px 5px; border-radius: 4px; }}
    .ok {{ color: #05603a; font-weight: 650; }}
    .error {{ color: #b42318; font-weight: 650; }}
  </style>
</head>
<body>
  <h1>Recibiendo captura</h1>
  <p id="status">Esperando datos desde Facebook...</p>
  <div id="details"></div>
  <script>
    const TOKEN = {json.dumps(config.token)};
    const statusEl = document.getElementById("status");
    const detailsEl = document.getElementById("details");
    let alreadyProcessed = false;

    function escapeHtml(value) {{
      return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;");
    }}

    function countersText(counters) {{
      return [
        `${{counters?.added || 0}} nuevos`,
        `${{counters?.updated || 0}} actualizados`,
        `${{counters?.duplicate || 0}} duplicados`,
        `${{counters?.ignored || 0}} ignorados`,
        `${{counters?.too_small || 0}} omitidos por menos de {config.min_members}`,
      ].join(", ");
    }}

    function renderResult(result) {{
      statusEl.className = "ok";
      statusEl.textContent = `Captura lista: ${{countersText(result.counters)}}.`;
      detailsEl.innerHTML = `
        <table>
          <tr><th>Procesados</th><td>${{escapeHtml(result.processed)}}</td></tr>
          <tr><th>Consulta</th><td>${{escapeHtml(result.query || "-")}}</td></tr>
          <tr><th>Ubicación</th><td>${{escapeHtml(result.location || "-")}}</td></tr>
          <tr><th>Categoría</th><td>${{escapeHtml(result.category || "-")}}</td></tr>
          <tr><th>Excel</th><td><code>${{escapeHtml(result.excel)}}</code></td></tr>
        </table>
        <p>Ya puedes volver a Facebook, bajar a más resultados y capturar otra vez.</p>
      `;
    }}

    function renderError(message) {{
      statusEl.className = "error";
      statusEl.textContent = message;
    }}

    async function savePayload(payload) {{
      const response = await fetch(`/capture?token=${{encodeURIComponent(TOKEN)}}`, {{
        method: "POST",
        headers: {{ "Content-Type": "application/json" }},
        body: JSON.stringify(payload),
      }});
      const result = await response.json();
      if (!response.ok || !result.ok) {{
        throw new Error(result.error || `HTTP ${{response.status}}`);
      }}
      renderResult(result);
    }}

    window.addEventListener("message", (event) => {{
      const message = event.data || {{}};
      if (message.token !== TOKEN) {{
        return;
      }}

      if (message.type === "progress") {{
        statusEl.textContent = `Capturando en Facebook... ${{message.count || 0}} grupos encontrados (${{message.step || 0}}/${{message.maxScrolls || "?"}}).`;
        return;
      }}

      if (message.type === "error") {{
        renderError(message.message || "No se pudo capturar.");
        return;
      }}

      if (message.type === "capture") {{
        if (alreadyProcessed) {{
          return;
        }}
        alreadyProcessed = true;
        const count = message.payload?.groups?.length || 0;
        statusEl.textContent = `Guardando ${{count}} grupos en Excel...`;
        savePayload(message.payload).catch((error) => renderError(`Error guardando captura: ${{error.message || error}}`));
      }}
    }});

    window.opener?.postMessage({{ token: TOKEN, type: "ready" }}, "*");
  </script>
</body>
</html>"""


def home_html(config: ServerConfig) -> str:
    href = bookmarklet_href(config)
    return f"""<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Capturador de grupos Facebook</title>
  <style>
    body {{ font-family: system-ui, sans-serif; margin: 34px; max-width: 820px; color: #17202a; }}
    h1 {{ font-size: 26px; margin-bottom: 8px; }}
    p, li {{ line-height: 1.5; }}
    .button {{
      display: inline-block; background: #155eef; color: white; padding: 10px 14px;
      border-radius: 6px; text-decoration: none; font-weight: 650;
    }}
    code {{ background: #eef2f6; padding: 2px 5px; border-radius: 4px; }}
    .panel {{ border: 1px solid #dde3ea; padding: 16px; border-radius: 8px; margin: 16px 0; }}
  </style>
</head>
<body>
  <h1>Capturador de grupos Facebook</h1>
  <p>Arrastra este botón a tu barra de marcadores:</p>
  <p><a class="button" href="{html.escape(href)}">Capturar grupos FB</a></p>
  <div class="panel">
    <p><strong>Configuración activa</strong></p>
    <p>Consulta fija: <code>{html.escape(config.query or "tomar desde URL de Facebook")}</code></p>
    <p>Ubicación: <code>{html.escape(config.location or "-")}</code></p>
    <p>Categoría: <code>{html.escape(config.category or "-")}</code></p>
    <p>Mínimo: <code>{config.min_members}</code> miembros/seguidores</p>
    <p>Scroll automático: <code>{config.scrolls}</code> bajadas, <code>{config.delay_ms}</code> ms entre cada una</p>
    <p>Excel: <code>{html.escape(str(config.excel_path.resolve()))}</code></p>
  </div>
  <ol>
    <li>Abre Facebook y busca grupos.</li>
    <li>Haz clic en <code>Capturar grupos FB</code>.</li>
    <li>El marcador baja lentamente, acumula grupos y actualiza el Excel.</li>
  </ol>
  <p>Este flujo solo lee resultados visibles mientras baja por la búsqueda. No hace publicaciones, no manda mensajes y no entra a grupos.</p>
</body>
</html>"""


class CaptureHandler(BaseHTTPRequestHandler):
    config: ServerConfig

    def log_message(self, format: str, *args: Any) -> None:
        print(f"{self.address_string()} - {format % args}")

    def send_text(self, body: str, status: HTTPStatus = HTTPStatus.OK, content_type: str = "text/html") -> None:
        encoded = body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", f"{content_type}; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Private-Network", "true")
        self.end_headers()
        self.wfile.write(encoded)

    def send_json(self, payload: dict[str, Any], status: HTTPStatus = HTTPStatus.OK) -> None:
        self.send_text(json.dumps(payload, ensure_ascii=False), status, "application/json")

    def do_OPTIONS(self) -> None:
        self.send_response(HTTPStatus.NO_CONTENT)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Private-Network", "true")
        self.send_header("Access-Control-Max-Age", "600")
        self.end_headers()

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.send_text(home_html(self.config))
            return

        if parsed.path == "/receiver":
            params = parse_qs(parsed.query)
            if params.get("token", [""])[0] != self.config.token:
                self.send_text("Token inválido.", HTTPStatus.FORBIDDEN, "text/plain")
                return
            self.send_text(receiver_html(self.config))
            return

        if parsed.path == "/health":
            self.send_json({"ok": True})
            return

        if parsed.path == "/capture":
            params = parse_qs(parsed.query)
            if params.get("token", [""])[0] != self.config.token:
                self.send_text("Token inválido.", HTTPStatus.FORBIDDEN, "text/plain")
                return

            raw_payload = params.get("payload", [""])[0]
            if not raw_payload:
                self.send_text("Falta payload.", HTTPStatus.BAD_REQUEST, "text/plain")
                return

            try:
                payload = json.loads(raw_payload)
            except json.JSONDecodeError as exc:
                self.send_text(f"Payload inválido: {exc}", HTTPStatus.BAD_REQUEST, "text/plain")
                return

            result = process_payload(payload, self.config)
            print(format_counters("Captura lista", result["counters"], self.config.min_members))
            self.send_text(result_html(result, self.config))
            return

        self.send_text("No encontrado.", HTTPStatus.NOT_FOUND, "text/plain")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path != "/capture":
            self.send_text("No encontrado.", HTTPStatus.NOT_FOUND, "text/plain")
            return

        params = parse_qs(parsed.query)
        length = int(self.headers.get("Content-Length") or "0")
        body = self.rfile.read(length).decode("utf-8")

        try:
            payload = json.loads(body)
        except json.JSONDecodeError as exc:
            self.send_json({"ok": False, "error": f"Payload inválido: {exc}"}, HTTPStatus.BAD_REQUEST)
            return

        token = params.get("token", [""])[0] or str(payload.get("token") or "")
        if token != self.config.token:
            self.send_json({"ok": False, "error": "Token inválido."}, HTTPStatus.FORBIDDEN)
            return

        result = process_payload(payload, self.config)
        print(format_counters("Captura lista", result["counters"], self.config.min_members))
        self.send_json(result)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Servidor local para capturar grupos visibles desde Facebook.")
    parser.add_argument("--host", default=DEFAULT_HOST)
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument("--data", default=DEFAULT_DATA_FILE)
    parser.add_argument("--excel", default=DEFAULT_EXCEL_FILE)
    parser.add_argument("--min-members", type=int, default=DEFAULT_MIN_MEMBERS)
    parser.add_argument("--query", default="", help="Consulta fija. Si se omite, se intenta leer desde la URL de Facebook.")
    parser.add_argument("--location", default="", help="Comuna/zona para las filas capturadas.")
    parser.add_argument("--category", default="", help="Categoría comercial para las filas capturadas.")
    parser.add_argument("--token", default="", help="Token opcional para el bookmarklet.")
    parser.add_argument("--scrolls", type=int, default=DEFAULT_SCROLLS, help="Cantidad máxima de bajadas automáticas.")
    parser.add_argument(
        "--delay-ms",
        type=int,
        default=DEFAULT_SCROLL_DELAY_MS,
        help="Pausa entre bajadas automáticas, en milisegundos.",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    config = ServerConfig(
        host=args.host,
        port=args.port,
        data_path=Path(args.data),
        excel_path=Path(args.excel),
        min_members=args.min_members,
        query=args.query,
        location=args.location,
        category=args.category,
        token=args.token or secrets.token_urlsafe(12),
        scrolls=max(0, args.scrolls),
        delay_ms=max(250, args.delay_ms),
    )
    CaptureHandler.config = config
    server = ThreadingHTTPServer((config.host, config.port), CaptureHandler)
    print("Servidor de captura iniciado.")
    print(f"Abre esta URL para instalar el bookmarklet: {config.home_url}")
    print("Presiona Ctrl+C para detener.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor detenido.")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
