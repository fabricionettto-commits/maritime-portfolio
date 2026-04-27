# Meeting sample note: this script generates the operational Knowledge Hub inventory.
from __future__ import annotations

import csv
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path


ROOTS = [
    Path(r"C:\capacity_planner\booking_reports"),
    Path(r"C:\Users\Portfolio_User\Portfolio Workspace\demo service SERVICE - Database\APP"),
    Path(r"C:\capacity_planner\booking_reports\Vessel\Project_Baplie"),
    Path(r"C:\Users\Portfolio_User\Portfolio Workspace\demo service SERVICE - Database\Arquivos_HTML"),
]

VAULT = Path(
    "C:\\Users\\Portfolio_User\\OneDrive - Portfolio Workspace\\\u00c1rea de Trabalho\\Projetos\\Cerebro"
)

IGNORED_PATH_PARTS = {"S&OP_Capacity", "_Arquivado_SOP_Capacity_Migrado_Cerebro_2026-04-26"}
OLD_VESSEL_PATH_PARTS = {"MERCURY / VOYAGE 1", "NEPTUNE / VOYAGE 3", "VENUS / VOYAGE 2", "MydTianjinVOYAGE 4", "JUPITER", "JUPITER"}

def should_ignore_path(path: Path) -> bool:
    parts = set(path.parts)
    if parts & IGNORED_PATH_PARTS:
        return True
    if "Navios" not in parts and parts & OLD_VESSEL_PATH_PARTS:
        return True
    return False

IMPORTANT_EXTENSIONS = {
    ".html",
    ".htm",
    ".py",
    ".xlsx",
    ".xls",
    ".xlsm",
    ".csv",
    ".json",
    ".md",
    ".jpg",
    ".jpeg",
    ".png",
    ".svg",
    ".pdf",
}


def file_uri(path: Path) -> str:
    return path.resolve().as_uri()


def safe_rel(path: Path, root: Path) -> str:
    try:
        return str(path.relative_to(root))
    except ValueError:
        return str(path)


def scan_files() -> list[dict[str, str]]:
    seen: set[str] = set()
    records: list[dict[str, str]] = []
    for root in ROOTS:
        if not root.exists():
            records.append(
                {
                    "root": str(root),
                    "relative_path": "",
                    "name": "",
                    "extension": "",
                    "size_kb": "",
                    "modified": "",
                    "path": "",
                    "uri": "",
                    "status": "missing_root",
                }
            )
            continue
        for path in root.rglob("*"):
            if should_ignore_path(path):
                continue
            if not path.is_file():
                continue
            resolved = str(path.resolve()).lower()
            if resolved in seen:
                continue
            seen.add(resolved)
            ext = path.suffix.lower()
            if ext not in IMPORTANT_EXTENSIONS:
                continue
            stat = path.stat()
            records.append(
                {
                    "root": str(root),
                    "relative_path": safe_rel(path, root),
                    "name": path.name,
                    "extension": ext or "[none]",
                    "size_kb": f"{stat.st_size / 1024:.1f}",
                    "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M"),
                    "path": str(path),
                    "uri": file_uri(path),
                    "status": "ok",
                }
            )
    return sorted(records, key=lambda item: (item["extension"], item["root"], item["relative_path"]))


def write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def write_csv(path: Path, records: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fields = ["root", "relative_path", "name", "extension", "size_kb", "modified", "path", "uri", "status"]
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=fields)
        writer.writeheader()
        writer.writerows(records)


def md_link(record: dict[str, str]) -> str:
    return f"[{record['name']}]({record['uri']})"


def table(records: list[dict[str, str]], limit: int = 120) -> str:
    if not records:
        return "_Nenhum arquivo encontrado._\n"
    lines = ["| Arquivo | Pasta | Modificado | KB |", "|---|---|---:|---:|"]
    for record in records[:limit]:
        lines.append(
            f"| {md_link(record)} | `{record['relative_path']}` | {record['modified']} | {record['size_kb']} |"
        )
    if len(records) > limit:
        lines.append(f"| ... | mais {len(records) - limit} arquivos no inventario CSV |  |  |")
    return "\n".join(lines) + "\n"


def build_index(records: list[dict[str, str]]) -> str:
    counter = Counter(record["extension"] for record in records if record["status"] == "ok")
    roots = sorted({record["root"] for record in records})
    total = sum(counter.values())
    lines = [
        "# Operational Knowledge Hub - Portfolio / Anonymous Logistics",
        "",
        "Este cofre centraliza consulta, rastreabimercuryde e links dos relatorios operacionais.",
        "",
        "## Acesso rapido",
        "",
        "- [[01_Pastas_Monitoradas]]",
        "- [[02_Inventario_Geral]]",
        "- [[03_Relatorios_HTML]]",
        "- [[04_Scripts_Python]]",
        "- [[05_Arquivos_Excel]]",
        "- [[06_Projetos_Chave]]",
        "- [[07_Rotina_de_Atualizacao]]",
        "- [[08_Manuais/08_Manuais|08_Manuais]]",
        "- [[08_Portfolio_Recruiter_Cases]]",
        "- [[09_Portifolio_Respostas_Procedimentos]]",
        "- [[10_Manual_Capacity_Planners]]",
        "- [[11_Manual_Operacional]]",
        "- [[12_SOP_Capacity/00 - Segundo Cerebro|12_SOP_Capacity]]",
        "- [[13_Navios_Operacional]]",
        "",
        "## Resumo",
        "",
        f"- Total indexado: **{total}** arquivos relevantes",
        f"- HTML: **{counter.get('.html', 0) + counter.get('.htm', 0)}**",
        f"- Python: **{counter.get('.py', 0)}**",
        f"- Excel: **{counter.get('.xlsx', 0) + counter.get('.xls', 0) + counter.get('.xlsm', 0)}**",
        f"- Imagens: **{counter.get('.png', 0) + counter.get('.jpg', 0) + counter.get('.jpeg', 0) + counter.get('.svg', 0)}**",
        "",
        "## Raizes monitoradas",
        "",
    ]
    lines.extend(f"- `{root}`" for root in roots)
    lines.extend(
        [
            "",
            "## Observacao",
            "",
            "Este cofre nao move os arquivos originais. Ele cria uma camada de consulta com links para as fontes reais.",
            "",
            f"Ultima geracao: `{datetime.now().strftime('%Y-%m-%d %H:%M')}`",
        ]
    )
    return "\n".join(lines) + "\n"


def build_roots_note(records: list[dict[str, str]]) -> str:
    by_root = defaultdict(list)
    for record in records:
        by_root[record["root"]].append(record)
    lines = ["# Pastas Monitoradas", ""]
    for root in ROOTS:
        root_records = by_root.get(str(root), [])
        ext_counter = Counter(record["extension"] for record in root_records if record["status"] == "ok")
        lines.extend(
            [
                f"## {root}",
                "",
                f"- Status: {'OK' if root.exists() else 'Nao encontrada'}",
                f"- Arquivos relevantes: {sum(ext_counter.values())}",
                f"- HTML: {ext_counter.get('.html', 0) + ext_counter.get('.htm', 0)}",
                f"- Python: {ext_counter.get('.py', 0)}",
                f"- Excel: {ext_counter.get('.xlsx', 0) + ext_counter.get('.xls', 0) + ext_counter.get('.xlsm', 0)}",
                "",
            ]
        )
    return "\n".join(lines)


def build_inventory_note(records: list[dict[str, str]]) -> str:
    ext_counter = Counter(record["extension"] for record in records if record["status"] == "ok")
    lines = ["# Inventario Geral", "", "## Por extensao", "", "| Extensao | Quantidade |", "|---|---:|"]
    for ext, count in ext_counter.most_common():
        lines.append(f"| `{ext}` | {count} |")
    lines.extend(["", "## Arquivos recentes", ""])
    recent = sorted(records, key=lambda item: item["modified"], reverse=True)
    lines.append(table(recent, limit=80))
    lines.extend(["", "Arquivo completo: [[Inventario/all_files.csv]]"])
    return "\n".join(lines)


def build_filtered_note(title: str, records: list[dict[str, str]], extensions: set[str], limit: int = 180) -> str:
    filtered = [record for record in records if record["extension"] in extensions]
    filtered = sorted(filtered, key=lambda item: item["modified"], reverse=True)
    return f"# {title}\n\n{table(filtered, limit=limit)}"


def build_projects_note(records: list[dict[str, str]]) -> str:
    keywords = {
        "Fat Anual / Moves": ["fat_anual", "moves", "dubai"],
        "Booking Board": ["booking_board", "board"],
        "Vessels": ["vessel", "navio", "vessel_one", "mercury", "voyage", "vessel_two"],
        "Project Baplie": ["project_baplie", "baplie"],
    }
    lines = ["# Projetos Chave", ""]
    for title, keys in keywords.items():
        matched = [
            record
            for record in records
            if any(key in (record["path"] + " " + record["name"]).lower().replace(" ", "_") for key in keys)
        ]
        lines.extend([f"## {title}", "", table(matched, limit=60), ""])
    return "\n".join(lines)


def build_routine_note() -> str:
    return """# Rotina de Atualizacao

## Atualizar o cerebro

Rode o gerador a partir da pasta `Fat_Anual`:

```powershell
python generate_cerebro_obsidian.py
```

## Fluxo recomendado

1. Rodar os processos/dashboards normalmente nas pastas originais.
2. Gerar ou atualizar os HTMLs e Excels.
3. Rodar `python generate_cerebro_obsidian.py`.
4. Abrir o Obsidian na pasta `Cerebro`.
5. Consultar pelo `00_Index` ou pelo inventario.

## Filosofia

O cerebro deve ser uma camada de consulta, nao uma copia baguncada dos arquivos.

Os arquivos de trabalho continuam nas pastas originais. O Obsidian guarda:

- mapa das pastas;
- links diretos para relatorios;
- regras de calculo;
- historico de decisoes;
- explicacoes de divergencias;
- notas por projeto.
"""


def main() -> None:
    records = scan_files()
    VAULT.mkdir(parents=True, exist_ok=True)
    (VAULT / "Inventario").mkdir(exist_ok=True)

    write_csv(VAULT / "Inventario" / "all_files.csv", records)
    write_text(VAULT / "00_Index.md", build_index(records))
    write_text(VAULT / "01_Pastas_Monitoradas.md", build_roots_note(records))
    write_text(VAULT / "02_Inventario_Geral.md", build_inventory_note(records))
    write_text(VAULT / "03_Relatorios_HTML.md", build_filtered_note("Relatorios HTML", records, {".html", ".htm"}))
    write_text(VAULT / "04_Scripts_Python.md", build_filtered_note("Scripts Python", records, {".py"}))
    write_text(
        VAULT / "05_Arquivos_Excel.md",
        build_filtered_note("Arquivos Excel", records, {".xlsx", ".xls", ".xlsm", ".csv"}),
    )
    write_text(VAULT / "06_Projetos_Chave.md", build_projects_note(records))
    write_text(VAULT / "07_Rotina_de_Atualizacao.md", build_routine_note())

    print(VAULT)
    print(f"indexed={len([r for r in records if r['status'] == 'ok'])}")


if __name__ == "__main__":
    main()











