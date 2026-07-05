#!/usr/bin/env python3
"""
Planificador diario de publicaciones para tus servicios en grupos de Facebook.

Idea central: publicar de forma sostenible y sin caer en spam, sin exponer tus
credenciales de Facebook y sin violar sus Terminos.

Este script NO inicia sesion en Facebook, NO guarda tu usuario/contrasena ni
cookies, y NO publica por si mismo. Lo que hace es:

  1. Elegir que grupos tocan hoy respetando un enfriamiento por grupo
     (no repetir el mismo grupo antes de X dias) y un tope diario. Asi la
     periodicidad es poco agresiva.
  2. Rotar variantes de tu mensaje y personalizarlo por grupo, para no publicar
     siempre el texto identico (eso es lo que Facebook marca como spam).
  3. Generar una pagina local `plan_hoy.html` con el enlace de cada grupo y el
     mensaje listo para copiar. Tu abres el grupo en tu navegador (donde ya
     estas logueado), pegas y publicas manualmente.
  4. Registrar lo hecho con `done`, que avanza la rotacion usando el mismo
     campo `last_published_at` del CRM.

Reutiliza `facebook_group_crm.py` para leer/guardar la base y regenerar el Excel.
"""

from __future__ import annotations

import argparse
import html
import json
import random
from datetime import date, datetime
from pathlib import Path
from typing import Any

import facebook_group_crm as crm


DEFAULT_CONFIG_FILE = "posting_config.json"
DEFAULT_MESSAGES_FILE = "mensajes_servicios.txt"
DEFAULT_PLAN_HTML = "plan_hoy.html"
DEFAULT_PLAN_JSON = "plan_hoy.json"

DEFAULT_DAILY_LIMIT = 6
DEFAULT_COOLDOWN_DAYS = 12
MESSAGE_SEPARATOR = "---"

STATUS_POSTED = "Publicado esta semana"


DEFAULT_CONFIG: dict[str, Any] = {
    "mensajes_file": DEFAULT_MESSAGES_FILE,
    "mensajes": [],
    "daily_limit": DEFAULT_DAILY_LIMIT,
    "cooldown_days": DEFAULT_COOLDOWN_DAYS,
    "min_priority": 0,
    "min_members": 0,
    "categorias": [],
    "spread_categorias": True,
}

SAMPLE_MESSAGES = """\
¡Hola a todos! 👋 Ofrezco mis servicios profesionales de [TU SERVICIO].
Trabajo con responsabilidad y precios justos. Si necesitas ayuda o una
cotización sin compromiso, escríbeme por interno o al 📱 +56 9 XXXX XXXX.
¡Gracias y buen día!
---
Buenas 🙌 Dejo por aquí mis servicios de [TU SERVICIO] para quien lo necesite.
Atención seria y a tiempo. Puedes contactarme al 📱 +56 9 XXXX XXXX o por
mensaje directo. ¡Saludos al grupo {grupo}!
---
Hola vecinos de {ubicacion} 🌟 Realizo [TU SERVICIO] con años de experiencia.
Si buscas a alguien de confianza, con gusto te atiendo. Escríbeme al
📱 +56 9 XXXX XXXX. ¡Que tengan una excelente semana!
"""


# --------------------------------------------------------------------------- #
# Configuracion y mensajes
# --------------------------------------------------------------------------- #
def load_config(path: Path) -> dict[str, Any]:
    config = dict(DEFAULT_CONFIG)
    if path.exists():
        with path.open("r", encoding="utf-8") as fh:
            user_config = json.load(fh)
        if isinstance(user_config, dict):
            config.update(user_config)
    return config


def load_messages(config: dict[str, Any], base_dir: Path) -> list[str]:
    """Toma las variantes desde el archivo de mensajes o desde el config."""
    messages: list[str] = []

    messages_file = config.get("mensajes_file")
    if messages_file:
        file_path = Path(messages_file)
        if not file_path.is_absolute():
            file_path = base_dir / file_path
        if file_path.exists():
            raw = file_path.read_text(encoding="utf-8")
            messages = split_messages(raw)

    if not messages:
        inline = config.get("mensajes") or []
        messages = [str(item).strip() for item in inline if str(item).strip()]

    return messages


def split_messages(raw: str) -> list[str]:
    blocks: list[str] = []
    current: list[str] = []
    for line in raw.splitlines():
        if line.strip() == MESSAGE_SEPARATOR:
            block = "\n".join(current).strip()
            if block:
                blocks.append(block)
            current = []
        else:
            current.append(line)
    tail = "\n".join(current).strip()
    if tail:
        blocks.append(tail)
    return blocks


def personalize(message: str, group: dict[str, Any]) -> str:
    replacements = {
        "grupo": group.get("name", "") or "",
        "ubicacion": group.get("location", "") or "",
        "categoria": group.get("category", "") or "",
    }
    text = message
    for key, value in replacements.items():
        text = text.replace("{" + key + "}", value)
    return text


# --------------------------------------------------------------------------- #
# Seleccion de grupos del dia
# --------------------------------------------------------------------------- #
def parse_iso_date(value: str) -> date | None:
    if not value:
        return None
    try:
        return datetime.strptime(str(value)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def days_since_last_post(group: dict[str, Any], today: date) -> int | None:
    last = parse_iso_date(group.get("last_published_at", ""))
    if last is None:
        return None
    return (today - last).days


def is_eligible(
    group: dict[str, Any],
    *,
    today: date,
    cooldown_days: int,
    min_priority: int,
    min_members: int,
    categorias: list[str],
) -> bool:
    if not group.get("url"):
        return False

    if min_members:
        members = group.get("members") or 0
        if members < min_members:
            return False

    if min_priority:
        if (group.get("priority") or 0) < min_priority:
            return False

    if categorias:
        wanted = {c.strip().lower() for c in categorias if c.strip()}
        if (group.get("category", "") or "").strip().lower() not in wanted:
            return False

    days = days_since_last_post(group, today)
    if days is not None and days < cooldown_days:
        return False

    return True


def ranking_key(group: dict[str, Any], today: date):
    """Menos reciente primero; nunca publicado va primero de todo."""
    days = days_since_last_post(group, today)
    never_posted = 0 if days is None else 1  # 0 ordena antes que 1
    recency = -(days if days is not None else 0)  # mas dias sin publicar = antes
    priority = -(group.get("priority") or 0)
    members = -(group.get("members") or 0)
    return (never_posted, recency, priority, members)


def spread_by_category(groups: list[dict[str, Any]], limit: int) -> list[dict[str, Any]]:
    """Round-robin por categoria para no publicar todo del mismo rubro un dia."""
    buckets: dict[str, list[dict[str, Any]]] = {}
    order: list[str] = []
    for group in groups:
        key = (group.get("category", "") or "sin categoria").strip().lower()
        if key not in buckets:
            buckets[key] = []
            order.append(key)
        buckets[key].append(group)

    selected: list[dict[str, Any]] = []
    while len(selected) < limit and any(buckets[key] for key in order):
        for key in order:
            if buckets[key]:
                selected.append(buckets[key].pop(0))
                if len(selected) >= limit:
                    break
    return selected


def select_groups(
    data: dict[str, Any],
    *,
    today: date,
    limit: int,
    cooldown_days: int,
    min_priority: int,
    min_members: int,
    categorias: list[str],
    spread: bool,
) -> list[dict[str, Any]]:
    eligible = [
        group
        for group in data.get("groups", [])
        if is_eligible(
            group,
            today=today,
            cooldown_days=cooldown_days,
            min_priority=min_priority,
            min_members=min_members,
            categorias=categorias,
        )
    ]
    eligible.sort(key=lambda g: ranking_key(g, today))

    if spread:
        return spread_by_category(eligible, limit)
    return eligible[:limit]


def assign_messages(
    groups: list[dict[str, Any]], messages: list[str], today: date
) -> list[dict[str, Any]]:
    """Asigna una variante por grupo, rotando y variando dia a dia."""
    plan_items: list[dict[str, Any]] = []
    if not messages:
        for group in groups:
            plan_items.append({"group": group, "message": ""})
        return plan_items

    # Desfase por dia para que un grupo no reciba siempre la misma variante.
    offset = today.toordinal() % len(messages)
    for index, group in enumerate(groups):
        variant = messages[(index + offset) % len(messages)]
        plan_items.append({"group": group, "message": personalize(variant, group)})
    return plan_items


# --------------------------------------------------------------------------- #
# Salidas: consola, JSON y HTML
# --------------------------------------------------------------------------- #
def print_plan(plan_items: list[dict[str, Any]], today: date) -> None:
    if not plan_items:
        print("No hay grupos elegibles para hoy con los filtros actuales.")
        print("Sube el tope diario, baja el enfriamiento o captura mas grupos.")
        return

    print(f"Plan de publicacion para {today.isoformat()} ({len(plan_items)} grupos):")
    print("-" * 78)
    for index, item in enumerate(plan_items, start=1):
        group = item["group"]
        members = group.get("members")
        members_text = f"{members:,}".replace(",", ".") if members else "-"
        days = days_since_last_post(group, today)
        last_text = "nunca" if days is None else f"hace {days} d"
        print(
            f"{index:>2}. {group.get('name', '')[:44]:<44} "
            f"{members_text:>9} miembros  ult: {last_text}"
        )
        print(f"    {group.get('url', '')}")
    print("-" * 78)
    print("Abre plan_hoy.html para copiar cada mensaje y publicar manualmente.")


def write_plan_json(path: Path, plan_items: list[dict[str, Any]], today: date) -> None:
    payload = {
        "date": today.isoformat(),
        "items": [
            {
                "id": item["group"].get("id", ""),
                "name": item["group"].get("name", ""),
                "url": item["group"].get("url", ""),
                "message": item["message"],
            }
            for item in plan_items
        ],
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def render_plan_html(plan_items: list[dict[str, Any]], today: date) -> str:
    cards: list[str] = []
    for index, item in enumerate(plan_items, start=1):
        group = item["group"]
        name = html.escape(group.get("name", ""))
        url = html.escape(group.get("url", ""), quote=True)
        category = html.escape(group.get("category", "") or "—")
        members = group.get("members")
        members_text = f"{members:,}".replace(",", ".") if members else "—"
        message = html.escape(item["message"])
        days = days_since_last_post(group, today)
        last_text = "nunca publicado" if days is None else f"hace {days} días"

        cards.append(
            f"""
    <article class="card" data-index="{index}">
      <header>
        <label class="check"><input type="checkbox" class="done"> <span class="num">{index}</span></label>
        <div class="meta">
          <h2>{name}</h2>
          <p class="sub">{category} · {members_text} miembros · {last_text}</p>
        </div>
      </header>
      <div class="actions">
        <a class="btn open" href="{url}" target="_blank" rel="noopener">1) Abrir grupo</a>
        <button class="btn copy" type="button">2) Copiar mensaje</button>
        <span class="hint">3) Pega, publica y marca la casilla ✓</span>
      </div>
      <textarea class="message" readonly rows="6">{message}</textarea>
    </article>"""
        )

    body = "\n".join(cards) if cards else (
        '<p class="empty">No hay grupos elegibles para hoy. '
        "Sube el tope diario, baja el enfriamiento o captura más grupos.</p>"
    )

    ids = ", ".join(
        html.escape(item["group"].get("id", "")) for item in plan_items if item["group"].get("id")
    )
    done_hint = (
        f"<code>python3 posting_planner.py done</code> (marca los {len(plan_items)} de hoy)"
        if plan_items
        else ""
    )

    return f"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Plan de publicación · {today.isoformat()}</title>
<style>
  :root {{ color-scheme: light dark; }}
  * {{ box-sizing: border-box; }}
  body {{ font-family: system-ui, -apple-system, Segoe UI, Roboto, sans-serif;
    margin: 0; padding: 24px; background: #f4f5f7; color: #1c1e21; }}
  @media (prefers-color-scheme: dark) {{ body {{ background: #18191a; color: #e4e6eb; }} }}
  .wrap {{ max-width: 780px; margin: 0 auto; }}
  h1 {{ font-size: 1.5rem; margin: 0 0 4px; }}
  .lead {{ color: #65676b; margin: 0 0 20px; }}
  @media (prefers-color-scheme: dark) {{ .lead {{ color: #b0b3b8; }} }}
  .card {{ background: #fff; border: 1px solid #dcdfe3; border-radius: 12px;
    padding: 16px; margin-bottom: 16px; }}
  @media (prefers-color-scheme: dark) {{ .card {{ background: #242526; border-color: #3a3b3c; }} }}
  .card.done {{ opacity: .5; }}
  header {{ display: flex; gap: 12px; align-items: flex-start; }}
  .meta h2 {{ font-size: 1.05rem; margin: 0; }}
  .sub {{ margin: 2px 0 0; font-size: .85rem; color: #65676b; }}
  @media (prefers-color-scheme: dark) {{ .sub {{ color: #b0b3b8; }} }}
  .num {{ font-weight: 700; }}
  .check {{ display: flex; align-items: center; gap: 6px; }}
  .actions {{ display: flex; flex-wrap: wrap; gap: 8px; align-items: center; margin: 12px 0 8px; }}
  .btn {{ display: inline-block; border: 0; border-radius: 8px; padding: 8px 14px;
    font-size: .9rem; font-weight: 600; cursor: pointer; text-decoration: none; }}
  .open {{ background: #1877f2; color: #fff; }}
  .copy {{ background: #e4e6eb; color: #1c1e21; }}
  .copy.copied {{ background: #31a24c; color: #fff; }}
  .hint {{ font-size: .82rem; color: #65676b; }}
  @media (prefers-color-scheme: dark) {{ .hint {{ color: #b0b3b8; }} }}
  textarea.message {{ width: 100%; border: 1px solid #dcdfe3; border-radius: 8px;
    padding: 10px; font: inherit; font-size: .92rem; resize: vertical;
    background: #f7f8fa; color: inherit; }}
  @media (prefers-color-scheme: dark) {{ textarea.message {{ background: #18191a; border-color: #3a3b3c; }} }}
  .note {{ background: #fff7e6; border: 1px solid #ffe1a8; border-radius: 10px;
    padding: 12px 14px; font-size: .88rem; margin-bottom: 20px; }}
  @media (prefers-color-scheme: dark) {{ .note {{ background: #3a2f14; border-color: #6b5518; }} }}
  code {{ background: rgba(128,128,128,.18); padding: 1px 6px; border-radius: 5px; }}
  .empty {{ text-align: center; color: #65676b; padding: 40px 0; }}
</style>
</head>
<body>
<div class="wrap">
  <h1>Plan de publicación · {today.isoformat()}</h1>
  <p class="lead">{len(plan_items)} grupos para hoy. Publica tú mismo/a en tu navegador ya conectado. No se usan tus credenciales.</p>
  <div class="note">
    <strong>Para no caer en spam:</strong> respeta el ritmo, no publiques el mismo texto en todos,
    responde comentarios y no repitas grupo antes del enfriamiento. Al terminar, en la terminal:
    {done_hint}
  </div>
  {body}
</div>
<script>
  document.querySelectorAll('.card').forEach(function (card) {{
    var copyBtn = card.querySelector('.copy');
    var textarea = card.querySelector('.message');
    var check = card.querySelector('.done');
    if (copyBtn && textarea) {{
      copyBtn.addEventListener('click', function () {{
        var text = textarea.value;
        var mark = function () {{
          copyBtn.classList.add('copied');
          copyBtn.textContent = '✓ Copiado';
          setTimeout(function () {{
            copyBtn.classList.remove('copied');
            copyBtn.textContent = '2) Copiar mensaje';
          }}, 1800);
        }};
        if (navigator.clipboard && navigator.clipboard.writeText) {{
          navigator.clipboard.writeText(text).then(mark, function () {{
            textarea.select(); document.execCommand('copy'); mark();
          }});
        }} else {{
          textarea.select(); document.execCommand('copy'); mark();
        }}
      }});
    }}
    if (check) {{
      check.addEventListener('change', function () {{
        card.classList.toggle('done', check.checked);
      }});
    }}
  }});
</script>
</body>
</html>
"""


def write_plan_html(path: Path, plan_items: list[dict[str, Any]], today: date) -> None:
    path.write_text(render_plan_html(plan_items, today), encoding="utf-8")


# --------------------------------------------------------------------------- #
# Comandos
# --------------------------------------------------------------------------- #
def command_init(args: argparse.Namespace) -> int:
    config_path = Path(args.config)
    messages_path = Path(DEFAULT_MESSAGES_FILE)

    created = []
    if config_path.exists() and not args.force:
        print(f"Ya existe {config_path} (usa --force para sobrescribir).")
    else:
        config_path.write_text(
            json.dumps(DEFAULT_CONFIG, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        created.append(str(config_path))

    if messages_path.exists() and not args.force:
        print(f"Ya existe {messages_path} (usa --force para sobrescribir).")
    else:
        messages_path.write_text(SAMPLE_MESSAGES, encoding="utf-8")
        created.append(str(messages_path))

    if created:
        print("Creado: " + ", ".join(created))
    print()
    print("Siguientes pasos:")
    print(f"  1. Edita {messages_path} con tus servicios reales (varias variantes separadas por '---').")
    print(f"  2. Ajusta {config_path} (tope diario, enfriamiento, filtros).")
    print("  3. Genera el plan del dia: python3 posting_planner.py plan")
    return 0


def resolve_plan_params(args: argparse.Namespace, config: dict[str, Any]) -> dict[str, Any]:
    def pick(cli_value, config_key, default):
        if cli_value is not None:
            return cli_value
        return config.get(config_key, default)

    categorias = args.categoria if args.categoria else list(config.get("categorias", []))
    return {
        "limit": int(pick(args.limit, "daily_limit", DEFAULT_DAILY_LIMIT)),
        "cooldown_days": int(pick(args.cooldown, "cooldown_days", DEFAULT_COOLDOWN_DAYS)),
        "min_priority": int(pick(args.min_priority, "min_priority", 0)),
        "min_members": int(pick(args.min_members, "min_members", 0)),
        "categorias": categorias,
        "spread": config.get("spread_categorias", True),
    }


def command_plan(args: argparse.Namespace) -> int:
    config_path = Path(args.config)
    config = load_config(config_path)
    base_dir = config_path.resolve().parent

    data = crm.load_database(Path(args.data))
    messages = load_messages(config, base_dir)
    if not messages:
        print("No hay mensajes configurados.")
        print("Corre: python3 posting_planner.py init  y edita mensajes_servicios.txt")
        return 1

    today = parse_iso_date(args.date) or date.today()
    params = resolve_plan_params(args, config)

    groups = select_groups(data, today=today, **params)
    if args.shuffle and groups:
        random.shuffle(groups)

    plan_items = assign_messages(groups, messages, today)

    write_plan_json(Path(args.plan_json), plan_items, today)
    write_plan_html(Path(args.plan_html), plan_items, today)

    print_plan(plan_items, today)
    if plan_items:
        print()
        print(f"Página lista: {Path(args.plan_html).resolve()}")
        print("Ábrela en tu navegador, copia cada mensaje y publica manualmente.")
        print("Cuando termines: python3 posting_planner.py done")
    return 0


def _mark_group_posted(group: dict[str, Any], today: date) -> None:
    group["last_published_at"] = today.isoformat()
    group["status"] = STATUS_POSTED
    group["messages_generated"] = int(group.get("messages_generated", 0) or 0) + 1
    group["updated_at"] = crm.now_iso()


def command_done(args: argparse.Namespace) -> int:
    data_path = Path(args.data)
    data = crm.load_database(data_path)
    today = parse_iso_date(args.date) or date.today()

    identifiers: list[str]
    if args.groups:
        identifiers = list(args.groups)
    else:
        plan_path = Path(args.plan_json)
        if not plan_path.exists():
            print(f"No encontre {plan_path}. Genera el plan con: python3 posting_planner.py plan")
            return 1
        plan = json.loads(plan_path.read_text(encoding="utf-8"))
        identifiers = [item.get("id", "") for item in plan.get("items", []) if item.get("id")]
        if not identifiers:
            print("El plan no tiene grupos con ID para marcar.")
            return 1

    marked = 0
    for identifier in identifiers:
        group = crm.find_by_identifier(data["groups"], identifier)
        if not group:
            print(f"  No encontre: {identifier}")
            continue
        _mark_group_posted(group, today)
        marked += 1
        print(f"  Publicado: {group.get('name')} ({today.isoformat()})")

    if marked:
        crm.save_database(data_path, data)
        crm.export_excel(data, Path(args.excel))
        print(f"Marcados {marked} grupos como publicados. Excel actualizado.")
    else:
        print("No se marco ningun grupo.")
    return 0


def command_status(args: argparse.Namespace) -> int:
    config = load_config(Path(args.config))
    data = crm.load_database(Path(args.data))
    today = date.today()
    cooldown = int(config.get("cooldown_days", DEFAULT_COOLDOWN_DAYS))

    groups = data.get("groups", [])
    total = len(groups)
    never = sum(1 for g in groups if not g.get("last_published_at"))
    eligible = [
        g
        for g in groups
        if is_eligible(
            g,
            today=today,
            cooldown_days=cooldown,
            min_priority=int(config.get("min_priority", 0)),
            min_members=int(config.get("min_members", 0)),
            categorias=list(config.get("categorias", [])),
        )
    ]

    print(f"Grupos en la base: {total}")
    print(f"Nunca publicados: {never}")
    print(f"Elegibles hoy (enfriamiento {cooldown} d): {len(eligible)}")
    print(f"Tope diario configurado: {config.get('daily_limit', DEFAULT_DAILY_LIMIT)}")

    upcoming = []
    for group in groups:
        days = days_since_last_post(group, today)
        if days is not None and days < cooldown:
            upcoming.append((cooldown - days, group))
    if upcoming:
        upcoming.sort(key=lambda item: item[0])
        print()
        print("En enfriamiento (dias para volver a estar disponible):")
        for remaining, group in upcoming[: args.limit]:
            print(f"  {remaining:>2} d  {group.get('name', '')[:52]}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Planifica publicaciones diarias de tus servicios en grupos de Facebook "
            "sin exponer credenciales y sin caer en spam. Tu publicas manualmente."
        )
    )
    parser.add_argument("--data", default=crm.DEFAULT_DATA_FILE, help="Base JSON del CRM.")
    parser.add_argument("--excel", default=crm.DEFAULT_EXCEL_FILE, help="Excel de salida del CRM.")
    parser.add_argument("--config", default=DEFAULT_CONFIG_FILE, help="Config del planificador.")

    subparsers = parser.add_subparsers(dest="command", required=True)

    init_parser = subparsers.add_parser("init", help="Crea config y archivo de mensajes de ejemplo.")
    init_parser.add_argument("--force", action="store_true", help="Sobrescribe archivos existentes.")
    init_parser.set_defaults(func=command_init)

    plan_parser = subparsers.add_parser("plan", help="Genera el plan del dia (consola + plan_hoy.html).")
    plan_parser.add_argument("--date", default="", help="Fecha YYYY-MM-DD. Si se omite usa hoy.")
    plan_parser.add_argument("--limit", type=int, default=None, help="Tope de grupos para hoy.")
    plan_parser.add_argument("--cooldown", type=int, default=None, help="Dias de enfriamiento por grupo.")
    plan_parser.add_argument("--min-priority", type=int, default=None, help="Prioridad minima.")
    plan_parser.add_argument("--min-members", type=int, default=None, help="Miembros minimos.")
    plan_parser.add_argument("--categoria", action="append", help="Filtra por categoria (repetible).")
    plan_parser.add_argument("--shuffle", action="store_true", help="Baraja el orden final.")
    plan_parser.add_argument("--plan-html", default=DEFAULT_PLAN_HTML, help="Ruta del HTML del plan.")
    plan_parser.add_argument("--plan-json", default=DEFAULT_PLAN_JSON, help="Ruta del JSON del plan.")
    plan_parser.set_defaults(func=command_plan)

    done_parser = subparsers.add_parser("done", help="Marca como publicados los grupos del plan (o los que indiques).")
    done_parser.add_argument("groups", nargs="*", help="IDs/URLs/nombres. Si se omite, usa plan_hoy.json.")
    done_parser.add_argument("--date", default="", help="Fecha YYYY-MM-DD. Si se omite usa hoy.")
    done_parser.add_argument("--plan-json", default=DEFAULT_PLAN_JSON, help="Ruta del JSON del plan.")
    done_parser.set_defaults(func=command_done)

    status_parser = subparsers.add_parser("status", help="Muestra elegibles hoy y grupos en enfriamiento.")
    status_parser.add_argument("--limit", type=int, default=20, help="Cuantos grupos en enfriamiento listar.")
    status_parser.set_defaults(func=command_status)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
