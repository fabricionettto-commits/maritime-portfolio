"""
=============================================================================
BOT TELEGRAM - OPERACAO MARITIMA (USSEA / ARTHA SHIPPING)
=============================================================================

FUNCIONAmercuryDES:
  - Consultas sob demanda: /navio, /status, /capacidade, /bookings, /feeder
  - Alertas automaticos: ocupacao critica, conexoes TIGER->demo service, novos bookings
  - Triagem de demanda: responde automaticamente perguntas simples
  - Escalada: avisa o operador so o que precisa de decisao

SETUP:
  1. Instale dependencias:
       pip install python-telegram-bot==20.7 pandas openpyxl xlrd schedule

  2. Crie seu bot:
       - Abra Telegram e fale com @BotFather
       - /newbot -> siga as instrucoes
       - Copie o TOKEN gerado

  3. Descubra seu CHAT_ID (para alertas proativos):
       - Mande qualquer mensagem pro bot
       - Acesse: https://api.telegram.org/bot<TOKEN>/getUpdates
       - Pegue o "id" dentro de "chat"

  4. Configure as variaveis abaixo e rode:
       python -m src.bot

ARQUIVOS NECESSARIOS:
  - BookingReport.xlsx (ou .xls) - exportado do Demo Export
  - ProgramacaoNavios_2026-03-24.csv - schedule de navios (para conexoes TIGER->demo service)
=============================================================================
"""

from __future__ import annotations

import asyncio
import io
import json
import logging
import os
import re
import threading
import time
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import schedule
from telegram import Bot, InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.constants import ParseMode
from telegram.error import BadRequest as TgBadRequest
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

from .config import ENV_FILE, load_settings


# =============================================================================
# CONFIGURACOES - EDITE AQUI
# =============================================================================

SETTINGS = load_settings()
BOT_TOKEN = SETTINGS.bot_token

# Chat IDs autorizados a receber alertas proativos
ADMIN_CHAT_IDS = [SETTINGS.allowed_chat_id] if SETTINGS.allowed_chat_id is not None else []

BASE_DIR = SETTINGS.base_dir
LOG_FILE = SETTINGS.logs_dir / "bot_maritimo.log"
VESSEL_OPERATION_BASE_PATH = BASE_DIR / "config" / "vessel_operation_base.json"

# Caminhos dos arquivos
BOOKING_REPORT_PATH = r"C:\capacity_planner\booking_reports\BookingReport.xls"
SCHEDULE_CSV_PATH = r"C:\capacity_planner\booking_reports\ProgramacaoNavios_2026-03-24.csv"
REPORT_BASE_URL = SETTINGS.report_base_url  # ex: http://192.168.1.100:5000

# Capacidades nominais dos navios demo service (TEU / Peso ton / Reefer plugs)
VESSEL_CAPS = {
    "MERCURY": {"teu": 1216, "weight": 16500, "reefer": 245},
    "VENUS": {"teu": 1217, "weight": 16800, "reefer": 194},
    "NEPTUNE": {"teu": 1060, "weight": 14200, "reefer": 295},
    "CAVAN": {"teu": 1217, "weight": 16800, "reefer": 194},
}

# Alertas automaticos - horarios (formato HH:MM)
ALERT_HORARIOS = ["08:00", "13:00", "18:00"]

# Limiares de risco
LIMIAR_CRITICO_PCT = 92
LIMIAR_FULL_PCT = 85
GAP_MAXIMO_DIAS = 5
GAP_MINIMO_DIAS = 1


# =============================================================================
# LOGGING
# =============================================================================

LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(message)s",
    level=getattr(logging, SETTINGS.log_level.upper(), logging.INFO),
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)


def _registrar_chat_id(chat_id: int | None) -> bool:
    if chat_id is None or chat_id in ADMIN_CHAT_IDS:
        return False

    env_lines = []
    if ENV_FILE.exists():
        env_lines = ENV_FILE.read_text(encoding="utf-8").splitlines()

    replaced = False
    for idx, line in enumerate(env_lines):
        if line.startswith("TELEGRAM_ALLOWED_CHAT_ID="):
            env_lines[idx] = f"TELEGRAM_ALLOWED_CHAT_ID={chat_id}"
            replaced = True
            break

    if not replaced:
        env_lines.append(f"TELEGRAM_ALLOWED_CHAT_ID={chat_id}")

    ENV_FILE.write_text("\n".join(env_lines) + "\n", encoding="utf-8")
    ADMIN_CHAT_IDS.append(chat_id)
    logger.info("[BOT] Chat autorizado automaticamente: %s", chat_id)
    return True


# =============================================================================
# LEITURA DO BOOKING REPORT
# =============================================================================

def _resolve_booking_path() -> Path:
    configured = Path(BOOKING_REPORT_PATH)
    candidates = []
    if configured.exists():
        candidates.append(configured)
    candidates.extend(
        [
            Path(r"C:\capacity_planner\booking_reports\BookingReport.xls"),
            Path(r"C:\capacity_planner\booking_reports\BookingReport.xlsx"),
            Path(r"C:\Users\Portfolio_User\Portfolio Workspace\demo service SERVICE - Database\APP\BookingReport.xls"),
            Path(r"C:\Users\Portfolio_User\Portfolio Workspace\demo service SERVICE - Database\APP\BookingReport.xlsx"),
        ]
    )
    existing = [path for path in candidates if path.exists()]
    if not existing:
        return configured
    return max(existing, key=lambda path: path.stat().st_mtime)


def _ler_booking(path: str) -> pd.DataFrame:
    """
    Le o BookingReport do Demo Export.
    Suporta: HTML disfarado de .xls/.xlsx, XLS real, XLSX real.
    """
    with open(path, "rb") as f:
        raw = f.read()

    sig = raw[:8].hex()

    # HTML disfarado (padrao Demo Export)
    if sig.startswith("3c737479") or b"<html" in raw[:500].lower() or b"<table" in raw[:2000].lower():
        logger.info("[BOOKING] Detectado HTML disfarado de Excel")
        text = raw.decode("utf-8", errors="replace")
        dfs = pd.read_html(io.StringIO(text))
        df_raw = max(dfs, key=len)

        # Montar header combinado (linha 2 = grupo, linha 3 = subcampo)
        r2 = df_raw.iloc[2].tolist()
        r3 = df_raw.iloc[3].tolist()
        cols = []
        for a, b in zip(r2, r3):
            a = str(a).strip() if str(a) != "nan" else ""
            b = str(b).strip() if str(b) != "nan" else ""
            if a and b and a != b:
                cols.append(f"{a}__{b}")
            elif b:
                cols.append(b)
            elif a:
                cols.append(a)
            else:
                cols.append("_col_")

        df = df_raw.iloc[4:].copy()
        df.columns = cols
        df = df.reset_index(drop=True)

    elif sig[:4] == "d0cf":
        logger.info("[BOOKING] Detectado XLS real")
        df = pd.read_excel(path, engine="xlrd", dtype=str, header=2)
    else:
        logger.info("[BOOKING] Detectado XLSX")
        df = pd.read_excel(path, dtype=str, header=2)

    df = df.fillna("")

    if "Vessel / Voyage" in df.columns:
        df = df[~df["Vessel / Voyage"].isin(["Vessel / Voyage", "TOTAL", ""])]
    elif "Vessel" in df.columns:
        df = df[~df["Vessel"].isin(["Vessel", "TOTAL", ""])]

    if "Vessel / Voyage" in df.columns:
        df["VESSEL"] = df["Vessel / Voyage"].astype(str).str.split("/").str[0].str.strip().str.upper()
        df["VOYAGE"] = df["Vessel / Voyage"].astype(str).str.split("/").str[1].fillna("").str.strip()
    else:
        df["VESSEL"] = df.get("Vessel", pd.Series([""] * len(df))).astype(str).str.strip().str.upper()
        df["VOYAGE"] = df.get("Voyage", pd.Series([""] * len(df))).astype(str).str.strip()

    pol_col = next((c for c in df.columns if c.strip().upper() == "POL"), None)
    pod_col = next((c for c in df.columns if c.strip().upper() == "POD"), None)
    df["POL"] = df[pol_col].astype(str).str.strip().str.upper() if pol_col else ""
    df["POD"] = df[pod_col].astype(str).str.strip().str.upper() if pod_col else ""

    col20 = next((c for c in df.columns if "20'" in c and "Bkg" in c), None)
    col40 = next((c for c in df.columns if "40'" in c and "Bkg" in c), None)
    if col20 and col40:
        df["TEU"] = (
            pd.to_numeric(df[col20], errors="coerce").fillna(0)
            + pd.to_numeric(df[col40], errors="coerce").fillna(0) * 2
        )
    else:
        df["TEU"] = 1

    wt_col = next((c for c in df.columns if "Weight" in c or "weight" in c or "KGS" in c), None)
    if wt_col:
        df["WEIGHT_KG"] = pd.to_numeric(df[wt_col], errors="coerce").fillna(0)
    else:
        df["WEIGHT_KG"] = 0

    tipo_col = next((c for c in df.columns if c.endswith("__Type") or c == "Type"), None)
    if tipo_col:
        df["IS_REEFER"] = df[tipo_col].astype(str).str.upper().str.strip().isin(
            ["RF", "REEFER", "REF", "40RF", "20RF", "RH", "HR"]
        )
    else:
        df["IS_REEFER"] = False

    feeder_col = next((c for c in df.columns if "Feeder" in c or "feeder" in c), None)
    if feeder_col:
        df["FEEDER_NAME"] = df[feeder_col].astype(str).str.strip()
    else:
        df["FEEDER_NAME"] = ""

    df = df.reset_index(drop=True)
    logger.info("[BOOKING] Lido: %s bookings | %s navios", len(df), df["VESSEL"].nunique())
    return df


def _calcular_capacidade_booking(df: pd.DataFrame) -> pd.DataFrame:
    resultados = []

    for (vessel, voyage, pol), grp in df.groupby(["VESSEL", "VOYAGE", "POL"]):
        vessel_key = next((k for k in VESSEL_CAPS if k in vessel.upper()), None)
        cap = VESSEL_CAPS.get(vessel_key, {"teu": None, "weight": None, "reefer": None})

        booked_teu = grp["TEU"].sum()
        booked_weight = grp["WEIGHT_KG"].sum() / 1000
        booked_reefer = grp[grp["IS_REEFER"]]["TEU"].sum()

        max_teu = cap["teu"]
        max_weight = cap["weight"]
        max_reefer = cap["reefer"]

        pct_teu = (booked_teu / max_teu * 100) if max_teu else 0
        pct_weight = (booked_weight / max_weight * 100) if max_weight else 0
        pct_reefer = (booked_reefer / max_reefer * 100) if max_reefer else 0

        max_pct = max(pct_teu, pct_weight, pct_reefer)

        if max_pct > 100:
            status = "OVERBOOKED"
        elif max_pct > LIMIAR_CRITICO_PCT:
            status = "CRITICO"
        elif max_pct > LIMIAR_FULL_PCT:
            status = "CHEIO"
        else:
            status = "OK"

        resultados.append(
            {
                "VESSEL": vessel,
                "VOYAGE": voyage,
                "POL": pol,
                "BOOKINGS": len(grp),
                "BOOKED_TEU": int(booked_teu),
                "MAX_TEU": max_teu,
                "PCT_TEU": round(pct_teu, 1),
                "AVAIL_TEU": int(max_teu - booked_teu) if max_teu else None,
                "BOOKED_WT_T": round(booked_weight, 0),
                "MAX_WT_T": max_weight,
                "PCT_WEIGHT": round(pct_weight, 1),
                "BOOKED_REEF": int(booked_reefer),
                "MAX_REEF": max_reefer,
                "PCT_REEFER": round(pct_reefer, 1),
                "STATUS": status,
            }
        )

    return pd.DataFrame(resultados)


def _calcular_capacidade(df: pd.DataFrame) -> pd.DataFrame:
    cap_booking = _calcular_capacidade_booking(df)
    cap_operacional = _carregar_capacidade_relatorio_operacional()

    if cap_operacional.empty:
        return cap_booking

    vessels_override = set(cap_operacional["VESSEL"].dropna().astype(str).tolist())
    cap_booking = cap_booking[~cap_booking["VESSEL"].astype(str).isin(vessels_override)]
    return pd.concat([cap_booking, cap_operacional], ignore_index=True)


def _analisar_conexoes_tiger(df: pd.DataFrame) -> list[dict]:
    resultados = []

    is_tiger = df["VESSEL"].str.contains("TIGER", na=False)
    lam_vessels = ["MERCURY", "VENUS", "NEPTUNE", "CAVAN"]
    is_lam = df["VESSEL"].apply(lambda v: any(lam in v for lam in lam_vessels))

    tigers = df[is_tiger]["VESSEL"].unique()
    lams = df[is_lam]["VESSEL"].unique()

    if len(tigers) == 0:
        return []

    etd_col = next((c for c in df.columns if c.strip().upper() == "ETD"), None)

    for tiger in tigers:
        grp_tiger = df[df["VESSEL"] == tiger]
        etd_tiger_vals = (
            pd.to_datetime(grp_tiger[etd_col], errors="coerce").dropna() if etd_col else pd.Series([], dtype="datetime64[ns]")
        )
        eta_tiger = etd_tiger_vals.min() if len(etd_tiger_vals) > 0 else None

        for lam in lams:
            grp_lam = df[df["VESSEL"] == lam]
            etd_lam_vals = (
                pd.to_datetime(grp_lam[etd_col], errors="coerce").dropna() if etd_col else pd.Series([], dtype="datetime64[ns]")
            )
            eta_lam = etd_lam_vals.min() if len(etd_lam_vals) > 0 else None

            if eta_tiger is None or eta_lam is None:
                gap = None
                risco = "SEM DATA"
            else:
                gap = (eta_lam - eta_tiger).days
                if gap < 0:
                    risco = "PERDEU CONEXAO"
                elif gap < GAP_MINIMO_DIAS:
                    risco = "CONEXAO APERTADA"
                elif gap > GAP_MAXIMO_DIAS:
                    risco = "ESPERA LONGA"
                else:
                    risco = "OK"

            resultados.append(
                {
                    "feeder": tiger,
                    "lam": lam,
                    "eta_feeder": eta_tiger.strftime("%d/%m/%Y") if eta_tiger is not None else "N/A",
                    "eta_lam": eta_lam.strftime("%d/%m/%Y") if eta_lam is not None else "N/A",
                    "gap_dias": gap,
                    "risco": risco,
                }
            )
            break

    return resultados


# =============================================================================
# FORMATACAO DE MENSAGENS
# =============================================================================

STATUS_EMOJI = {
    "OK": "ðŸŸ¢",
    "CHEIO": "ðŸŸ¡",
    "CRITICO": "ðŸ”´",
    "OVERBOOKED": "ðŸš¨",
}

RISCO_EMOJI = {
    "OK": "",
    "CONEXAO APERTADA": "Warning",
    "ESPERA LONGA": "ðŸ•",
    "PERDEU CONEXAO": "X",
    "SEM DATA": "?",
}


def _load_vessel_operation_base() -> list[dict]:
    if not VESSEL_OPERATION_BASE_PATH.exists():
        logger.warning("[BASE] Arquivo de base nao encontrado: %s", VESSEL_OPERATION_BASE_PATH)
        return []

    data = json.loads(VESSEL_OPERATION_BASE_PATH.read_text(encoding="utf-8"))
    return data.get("vessels", [])


VESSEL_OPERATION_BASE = _load_vessel_operation_base()


def _get_report_local_path(entry: dict) -> Path | None:
    """Retorna o caminho local do HTML do relatorio principal do navio."""
    reports = entry.get("reports", {})
    report_key = "departed" if _vessel_has_departed(entry) else "upcoming"
    # Prioridade: caminho especifico do estado -> caminho generico
    local_key = f"local_html_{report_key}" if report_key == "upcoming" else "local_html"
    path_str = reports.get(local_key) or reports.get("local_html")
    if path_str:
        p = Path(path_str)
        if p.exists():
            return p
    # Fallback: tenta derivar da rota local (file:///)
    report_uri = reports.get(report_key) or reports.get("booking", "")
    if report_uri.startswith("file:///"):
        p = Path(report_uri.replace("file:///", "").replace("/", "\\"))
        if p.exists():
            return p
    return None


def _menu_principal_inline() -> InlineKeyboardMarkup:
    """Monta o menu com um botao por navio (callback para envio do relatorio)."""
    keyboard = []
    # Botao de Schedule no TOPO
    keyboard.append([InlineKeyboardButton("ðŸ“… >>> BAIXAR demo service SCHEDULE <<<", callback_data="menu:schedule")])
    
    for entry in VESSEL_OPERATION_BASE:
        key = entry.get("key", "")
        voyage = entry.get("voyage", "")
        label = f"{key.replace(' ', '')}{voyage}"
        cb_data = f"navio:{label}"
        keyboard.append([InlineKeyboardButton(f"ðŸš¢ {label}", callback_data=cb_data)])

    # Fallback se vessel_operation_base estiver vazio
    if not keyboard:
        keyboard = [
            [InlineKeyboardButton("ðŸš¢ MERCURY / VOYAGE 1", callback_data="navio:MERCURY / VOYAGE 1")],
            [InlineKeyboardButton("ðŸš¢ VENUS / VOYAGE 2", callback_data="navio:VENUS / VOYAGE 2")],
            [InlineKeyboardButton("ðŸš¢ NEPTUNE / VOYAGE 3", callback_data="navio:NEPTUNE / VOYAGE 3")],
        ]

    keyboard.append([InlineKeyboardButton("ðŸ”„ Atualizar Menu", callback_data="menu:inicio")])
    return InlineKeyboardMarkup(keyboard)


def _find_vessel_operation_entry(text: str) -> dict | None:
    texto = str(text).upper().strip()
    if not texto:
        return None

    for entry in VESSEL_OPERATION_BASE:
        aliases = [alias.upper() for alias in entry.get("aliases", [])]
        key = str(entry.get("key", "")).upper().strip()
        voyage = str(entry.get("voyage", "")).upper().strip()
        candidatos = [value for value in [key, voyage, *aliases] if value]
        if any(candidato in texto or texto in candidato for candidato in candidatos):
            return entry
    return None


def _get_vessel_operation_entry(vessel_names: list[str]) -> dict | None:
    for vessel_name in vessel_names:
        entry = _find_vessel_operation_entry(vessel_name)
        if entry:
            return entry
    return None


def _vessel_has_departed(entry: dict) -> bool:
    last_port_eta = entry.get("last_port_eta", "").strip()
    if not last_port_eta:
        return False
    try:
        return datetime.strptime(last_port_eta, "%Y-%m-%d").date() <= date.today()
    except ValueError:
        return False


def _relatorio_principal_navio(vessel_names: list[str]) -> str | None:
    entry = _get_vessel_operation_entry(vessel_names)
    if not entry:
        return None

    reports = entry.get("reports", {})
    report_key = "departed" if _vessel_has_departed(entry) else "upcoming"
    report_uri = reports.get(report_key) or reports.get("booking")
    return str(report_uri).strip() if report_uri else None


def _ordenar_portos_por_rotacao(vessel_names: list[str], pols: list[str]) -> list[str]:
    pols_unicos = []
    for pol in pols:
        if pol not in pols_unicos:
            pols_unicos.append(pol)

    entry = _get_vessel_operation_entry(vessel_names)
    if entry:
        rotacao = entry.get("rotation", [])
        ordem = {porto: idx for idx, porto in enumerate(rotacao)}
        return sorted(pols_unicos, key=lambda porto: (ordem.get(porto, 999), porto))

    return sorted(pols_unicos)


def _numero_html_para_float(valor: str) -> float:
    valor_limpo = valor.strip()
    if "." in valor_limpo and "," in valor_limpo:
        valor_limpo = valor_limpo.replace(".", "").replace(",", ".")
    elif "," in valor_limpo:
        partes = valor_limpo.split(",")
        if len(partes[-1]) == 3:
            valor_limpo = valor_limpo.replace(",", "")
        else:
            valor_limpo = valor_limpo.replace(",", ".")
    else:
        valor_limpo = valor_limpo.replace(".", "")
    return float(valor_limpo) if valor_limpo else 0.0


def _montar_linha_capacidade(
    vessel: str,
    voyage: str,
    pol: str,
    units: int,
    teu: int,
    tons: float,
    reefer: int,
) -> dict:
    vessel_key = next((k for k in VESSEL_CAPS if k in vessel.upper()), None)
    cap = VESSEL_CAPS.get(vessel_key, {"teu": None, "weight": None, "reefer": None})

    max_teu = cap["teu"]
    max_weight = cap["weight"]
    max_reefer = cap["reefer"]

    pct_teu = (teu / max_teu * 100) if max_teu else 0
    pct_weight = (tons / max_weight * 100) if max_weight else 0
    pct_reefer = (reefer / max_reefer * 100) if max_reefer else 0
    max_pct = max(pct_teu, pct_weight, pct_reefer)

    if max_pct > 100:
        status = "OVERBOOKED"
    elif max_pct > LIMIAR_CRITICO_PCT:
        status = "CRITICO"
    elif max_pct > LIMIAR_FULL_PCT:
        status = "CHEIO"
    else:
        status = "OK"

    return {
        "VESSEL": vessel,
        "VOYAGE": voyage,
        "POL": pol,
        "BOOKINGS": units,
        "BOOKED_TEU": int(teu),
        "MAX_TEU": max_teu,
        "PCT_TEU": round(pct_teu, 1),
        "AVAIL_TEU": int(max_teu - teu) if max_teu else None,
        "BOOKED_WT_T": round(tons, 0),
        "MAX_WT_T": max_weight,
        "PCT_WEIGHT": round(pct_weight, 1),
        "BOOKED_REEF": int(reefer),
        "MAX_REEF": max_reefer,
        "PCT_REEFER": round(pct_reefer, 1),
        "STATUS": status,
    }


def _carregar_capacidade_relatorio_operacional() -> pd.DataFrame:
    resultados = []

    for entry in VESSEL_OPERATION_BASE:
        reports = entry.get("reports", {})
        operational_html = reports.get("operational_totals_html")
        operational_type = reports.get("operational_totals_type", "").strip()
        if not operational_html:
            continue

        report_path = Path(operational_html)
        if not report_path.exists():
            continue

        html = report_path.read_text(encoding="utf-8", errors="replace")
        if operational_type == "moves_all_ports":
            secao_match = re.search(
                rf'<div class="stitle">{re.escape("Load revenue by POO")}</div>\s*<div class="port-card-grid">(.*)</div>\s*<div class="stitle">',
                html,
                flags=re.DOTALL,
            )
            if not secao_match:
                logger.warning("[OPERACAO] Secao nao encontrada em %s", report_path.name)
                continue

            cards_html = secao_match.group(1)
            card_blocks = re.findall(r"<div class=\"port-card\">(.*)</div>\s*</div>", cards_html, flags=re.DOTALL)

            for card_html in card_blocks:
                titulo = re.search(r"<h3>.*\b([A-Z]{5})</h3>", card_html)
                if not titulo:
                    continue

                pol = titulo.group(1)
                metricas = dict(re.findall(r"<span>([^<]+)</span><strong>([^<]+)</strong>", card_html))
                if "TEUs" not in metricas:
                    continue

                resultados.append(
                    _montar_linha_capacidade(
                        vessel=entry["key"],
                        voyage=entry.get("voyage", ""),
                        pol=pol,
                        units=int(_numero_html_para_float(metricas.get("Units", "0"))),
                        teu=int(_numero_html_para_float(metricas.get("TEUs", "0"))),
                        tons=_numero_html_para_float(metricas.get("Tons", "0")),
                        reefer=int(_numero_html_para_float(metricas.get("Reefer", "0"))),
                    )
                )
        elif operational_type == "disch_load_load_side":
            rows = re.findall(
                r"<tr><td rowspan='2' class='pol-cell'>([A-Z0-9]{5})</td><td class='src-disc'>DISC</td>.*</tr><tr class='row-load'><td class='src-load'>LOAD</td><td>([^<]+)</td><td>([^<]+)</td><td>([^<]+)</td><td>([^<]+)</td><td class='gold-cell'>([^<]+)</td><td class='gold-cell'>([^<]+)</td><td class='gold-cell'>([^<]+)</td><td class='gold-cell'>([^<]+)</td></tr>",
                html,
                flags=re.DOTALL,
            )
            if not rows:
                logger.warning("[OPERACAO] Linhas LOAD nao encontradas em %s", report_path.name)
                continue

            for pol, _dc20, _hc40, _rh40, _fr40, units, teus, tons, plugs in rows:
                resultados.append(
                    _montar_linha_capacidade(
                        vessel=entry["key"],
                        voyage=entry.get("voyage", ""),
                        pol=pol,
                        units=int(_numero_html_para_float(units)),
                        teu=int(_numero_html_para_float(teus)),
                        tons=_numero_html_para_float(tons),
                        reefer=int(_numero_html_para_float(plugs)),
                    )
                )
        else:
            logger.warning("[OPERACAO] Tipo de base nao suportado para %s", entry.get("key", ""))

    return pd.DataFrame(resultados)


def fmt_capacidade_resumo(cap_df: pd.DataFrame) -> str:
    if cap_df.empty:
        return "Nenhum dado de capacidade disponivel."

    linhas = ["*RESUMO DE CAPACIDADE*\n"]
    for _, row in cap_df.iterrows():
        emoji = STATUS_EMOJI.get(row["STATUS"], "OK")
        vessel_short = row["VESSEL"].split()[0]
        teu_str = (
            f"{row['BOOKED_TEU']}/{row['MAX_TEU']} TEU ({row['PCT_TEU']}%)"
            if row["MAX_TEU"]
            else f"{row['BOOKED_TEU']} TEU"
        )
        reef_str = f" | Reefer: {row['BOOKED_REEF']}/{row['MAX_REEF']}" if row["MAX_REEF"] else ""
        linhas.append(f"{emoji} *{vessel_short}/{row['VOYAGE']}* [{row['POL']}]")
        linhas.append(f"   {teu_str}{reef_str}")
        if row["STATUS"] not in ("OK", "CHEIO"):
            linhas.append(f"   Status: *{row['STATUS']}*")
        linhas.append("")

    linhas.append(f"_Atualizado: {datetime.now().strftime('%d/%m %H:%M')}_")
    return "\n".join(linhas)


def fmt_navio_detalhe(df: pd.DataFrame, vessel_query: str) -> str:
    mask = df["VESSEL"].str.contains(vessel_query.upper(), na=False)
    subset = df[mask]

    if subset.empty:
        return f"Nenhum booking encontrado para: *{vessel_query}*"

    vessel_names = subset["VESSEL"].dropna().astype(str).unique().tolist()
    pols = _ordenar_portos_por_rotacao(
        vessel_names,
        [pol for pol in subset["POL"].dropna().astype(str).tolist() if pol.strip()],
    )
    report_uri = _relatorio_principal_navio(vessel_names)

    linhas = [f"*DETALHES: {vessel_query.upper()}*\n"]
    if report_uri:
        linhas.append(f"Relatorio principal: `{report_uri}`")
        linhas.append("")
    linhas.append(f"Portos: {' | '.join(pols)}" if pols else "Portos: N/A")
    return "\n".join(linhas)



def fmt_navio_detalhe(df: pd.DataFrame, vessel_query: str) -> str:
    query_upper = vessel_query.upper().strip()
    entry = _find_vessel_operation_entry(query_upper)

    if entry:
        aliases = [entry.get("key", ""), *entry.get("aliases", [])]
        regex = "|".join(re.escape(str(alias).upper()) for alias in aliases if str(alias).strip())
        mask = df["VESSEL"].astype(str).str.upper().str.contains(regex, na=False, regex=True)
    else:
        mask = df["VESSEL"].astype(str).str.upper().str.contains(query_upper, na=False, regex=False)

    subset = df[mask]

    if subset.empty and not entry:
        return f"Nenhum booking encontrado para: *{vessel_query}*"

    vessel_names = subset["VESSEL"].dropna().astype(str).unique().tolist()
    if entry and not vessel_names:
        vessel_names = [str(entry.get("key", "")).strip()]

    pols = _ordenar_portos_por_rotacao(
        vessel_names,
        [pol for pol in subset["POL"].dropna().astype(str).tolist() if pol.strip()],
    )
    report_uri = _relatorio_principal_navio(vessel_names)
    vessel_label = str(entry.get("key", "")).strip() if entry else query_upper
    voyage = str(entry.get("voyage", "")).strip() if entry else ""

    linhas = [f"*DETALHES: {vessel_label}*\n"]
    if voyage:
        linhas.append(f"Viagem: {vessel_label.replace(' ', '')}{voyage}")
        linhas.append("")
    if report_uri:
        linhas.append(f"Relatorio principal: `{report_uri}`")
        linhas.append("")
    linhas.append(f"Portos: {' | '.join(pols)}" if pols else "Portos: N/A")
    return "\n".join(linhas)


def _texto_menu_inicial(chat_id: str | int, chat_registrado: bool) -> str:
    texto = (
        "*BOT OPERACAO MARITIMA*\n\n"
        "Consulte um navio pelo nome ou pela viagem.\n\n"
        "`/navio <nome ou viagem>`\n\n"
        "*Exemplos*\n"
        "`/navio MERCURY / VOYAGE 1`\n"
        "`/navio VENUS / VOYAGE 2`\n"
        "`/navio NEPTUNE / VOYAGE 3`\n\n"
        "_Use os botoes abaixo para abrir mais rapido._\n\n"
        "_Em producao: bookings e deadlines._\n\n"
        f"_Seu chat ID: {chat_id}_"
    )
    if chat_registrado:
        texto += "\n_Alertas proativos habilitados para este chat._"
    return texto


def fmt_conexoes_tiger(conexoes: list[dict]) -> str:
    if not conexoes:
        return "Nenhuma conexao TIGER encontrada nos dados atuais."

    linhas = ["*CONEXOES TIGER -> demo service*\n"]
    for c in conexoes:
        emoji = RISCO_EMOJI.get(c["risco"], "?")
        gap_str = f"{c['gap_dias']} dias" if c["gap_dias"] is not None else "N/A"
        linhas.append(f"{emoji} *{c['feeder']}* -> *{c['lam']}*")
        linhas.append(f"   Feeder ETA: {c['eta_feeder']} | demo service ETD: {c['eta_lam']}")
        linhas.append(f"   Gap: {gap_str} | Risco: {c['risco']}")
        linhas.append("")

    return "\n".join(linhas)


# =============================================================================
# HANDLERS DOS COMANDOS
# =============================================================================

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    del context
    chat_id = update.effective_chat.id if update.effective_chat else "N/A"
    chat_registrado = _registrar_chat_id(chat_id if isinstance(chat_id, int) else None)
    logger.info("[BOT] /start de chat_id=%s", chat_id)
    texto = _texto_menu_inicial(chat_id, chat_registrado)
    if update.message:
        await update.message.reply_text(
            texto,
            parse_mode=ParseMode.MARKDOWN,
            reply_markup=_menu_principal_inline(),
        )


async def cmd_capacidade(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    del context
    if update.message:
        await update.message.reply_text("Analisando capacidade...", parse_mode=ParseMode.MARKDOWN)
    try:
        df = _ler_booking(str(_resolve_booking_path()))
        cap = _calcular_capacidade(df)
        msg = fmt_capacidade_resumo(cap)
        if update.message:
            await update.message.reply_text(msg, parse_mode=ParseMode.MARKDOWN)
    except Exception as e:
        logger.error("[cmd_capacidade] %s", e)
        if update.message:
            await update.message.reply_text(f"Erro ao ler BookingReport: {e}")


async def cmd_navio(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not context.args:
        if update.message:
            await update.message.reply_text("Uso: /navio <nome>\nEx: /navio MERCURY")
        return

    query = " ".join(context.args).strip()
    try:
        df = _ler_booking(str(_resolve_booking_path()))
        msg = fmt_navio_detalhe(df, query)
        if update.message:
            await update.message.reply_text(msg, parse_mode=ParseMode.MARKDOWN)
    except Exception as e:
        logger.error("[cmd_navio] %s", e)
        if update.message:
            await update.message.reply_text(f"Erro: {e}")


async def cb_menu_navio(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    del context
    query = update.callback_query
    if not query:
        return

    await query.answer()
    data = query.data or ""

    try:
        if data == "menu:schedule":
            print("\n!!! CLIQUE NO SCHEDULE DETECTADO !!!\n")
            logger.info("[BOT] Solicitado demo service Schedule")
            schedule_path = Path(BASE_DIR) / "reports" / "Demo_Service_Schedule.html"
            if not schedule_path.exists():
                # Tenta fallback se o BASE_DIR estiver estranho
                schedule_path = Path(os.getcwd()) / "reports" / "Demo_Service_Schedule.html"
            
            logger.info("[BOT] Caminho Schedule: %s (Existe: %s)", schedule_path, schedule_path.exists())
            
            if schedule_path.exists():
                try:
                    await query.edit_message_text(
                        "*ðŸ“… demo service Schedule*\n\nðŸ“¤ _Enviando cronograma atualizado..._",
                        parse_mode=ParseMode.MARKDOWN,
                        reply_markup=InlineKeyboardMarkup([[
                            InlineKeyboardButton("¸ Voltar ao Menu", callback_data="menu:inicio"),
                        ]])
                    )
                except TgBadRequest as e:
                    if "not modified" not in str(e).lower():
                        raise

                filename = f"demo service_Schedule_{datetime.now().strftime('%d%m%Y')}.html"
                with open(schedule_path, "rb") as f:
                    await query.message.reply_document(
                        document=f,
                        filename=filename,
                        caption="ðŸ“… *demo service Position / Schedule*",
                        parse_mode=ParseMode.MARKDOWN,
                    )
            else:
                try:
                    await query.edit_message_text(
                        "Warning Arquivo de Schedule nao encontrado.",
                        reply_markup=InlineKeyboardMarkup([[
                            InlineKeyboardButton("¸ Voltar ao Menu", callback_data="menu:inicio"),
                        ]])
                    )
                except TgBadRequest as e:
                    if "not modified" not in str(e).lower():
                        raise
            return

        if data == "menu:inicio":
            chat_id = update.effective_chat.id if update.effective_chat else "N/A"
            texto = _texto_menu_inicial(chat_id, False)
            try:
                await query.edit_message_text(
                    texto,
                    parse_mode=ParseMode.MARKDOWN,
                    reply_markup=_menu_principal_inline(),
                )
            except TgBadRequest as e:
                if "not modified" not in str(e).lower():
                    raise
            return

        if not data.startswith("navio:"):
            return

        vessel_query = data.split(":", 1)[1].strip()

        # Busca entrada do navio e arquivo local
        entry = _find_vessel_operation_entry(vessel_query)
        local_path = _get_report_local_path(entry) if entry else None

        if local_path:
            # Monta resumo de texto
            df = _ler_booking(str(_resolve_booking_path()))
            msg = fmt_navio_detalhe(df, vessel_query)

            # Edita a mensagem com resumo + botao voltar
            voltar_markup = InlineKeyboardMarkup([[
                InlineKeyboardButton("¸ Voltar ao Menu", callback_data="menu:inicio"),
            ]])
            try:
                await query.edit_message_text(
                    msg + "\n\nðŸ“¤ _Enviando relatorio..._",
                    parse_mode=ParseMode.MARKDOWN,
                    reply_markup=voltar_markup,
                )
            except TgBadRequest as e:
                if "not modified" not in str(e).lower():
                    raise

            # Envia o HTML como documento para download
            vessel_label = str(entry.get("key", vessel_query)).replace(" ", "_") if entry else vessel_query
            filename = f"Relatorio_{vessel_label}_{datetime.now().strftime('%d%m%Y')}.html"
            with open(local_path, "rb") as f:
                await query.message.reply_document(
                    document=f,
                    filename=filename,
                    caption=f"ðŸ“„ Relatorio *{vessel_label}* - abra no seu browser",
                    parse_mode=ParseMode.MARKDOWN,
                )
        else:
            # Fallback: sem arquivo local, mostra texto apenas
            df = _ler_booking(str(_resolve_booking_path()))
            msg = fmt_navio_detalhe(df, vessel_query)
            msg += "\n\nWarning _Arquivo de relatorio nao encontrado no servidor._"
            voltar_markup = InlineKeyboardMarkup([[
                InlineKeyboardButton("¸ Voltar ao Menu", callback_data="menu:inicio"),
            ]])
            try:
                await query.edit_message_text(
                    msg,
                    parse_mode=ParseMode.MARKDOWN,
                    reply_markup=voltar_markup,
                )
            except TgBadRequest as e:
                if "not modified" not in str(e).lower():
                    raise

    except TgBadRequest as e:
        logger.warning("[cb_menu_navio] BadRequest ignorado: %s", e)
    except Exception as e:
        logger.error("[cb_menu_navio] %s", e)
        try:
            await query.edit_message_text("Erro interno. Tente /ajuda")
        except Exception:
            pass


async def cmd_schedule(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    del context
    chat_id = update.effective_chat.id if update.effective_chat else "N/A"
    logger.info("[BOT] /schedule solicitado por chat_id=%s", chat_id)
    
    schedule_path = Path(BASE_DIR) / "reports" / "Demo_Service_Schedule.html"
    if not schedule_path.exists():
        schedule_path = Path(os.getcwd()) / "reports" / "Demo_Service_Schedule.html"

    if schedule_path.exists():
        if update.message:
            await update.message.reply_text("ðŸ“… *demo service Schedule*\n\nðŸ“¤ _Enviando cronograma..._", parse_mode=ParseMode.MARKDOWN)
            
            filename = f"demo service_Schedule_{datetime.now().strftime('%d%m%Y')}.html"
            with open(schedule_path, "rb") as f:
                await update.message.reply_document(
                    document=f,
                    filename=filename,
                    caption="ðŸ“… *demo service Position / Schedule*",
                    parse_mode=ParseMode.MARKDOWN,
                )
    else:
        if update.message:
            await update.message.reply_text("Warning Arquivo de Schedule nao encontrado no servidor.")

async def cmd_status(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    del context
    if update.message:
        await update.message.reply_text("Verificando navios criticos...", parse_mode=ParseMode.MARKDOWN)
    try:
        df = _ler_booking(str(_resolve_booking_path()))
        cap = _calcular_capacidade(df)
        criticos = cap[cap["STATUS"].isin(["CRITICO", "OVERBOOKED"])]

        if criticos.empty:
            if update.message:
                await update.message.reply_text("Todos os navios estao dentro do limite. Nenhum alerta critico.")
            return

        linhas = ["*NAVIOS EM ALERTA*\n"]
        for _, row in criticos.iterrows():
            emoji = STATUS_EMOJI.get(row["STATUS"], "ðŸ”´")
            avail = row["AVAIL_TEU"] if row["AVAIL_TEU"] is not None else "N/A"
            linhas.append(f"{emoji} *{row['VESSEL']}/{row['VOYAGE']}* [{row['POL']}]")
            linhas.append(f"   TEU: {row['BOOKED_TEU']}/{row['MAX_TEU']} ({row['PCT_TEU']}%) | Disponivel: {avail}")
            linhas.append(f"   Status: *{row['STATUS']}*")
            linhas.append("")

        if update.message:
            await update.message.reply_text("\n".join(linhas), parse_mode=ParseMode.MARKDOWN)
    except Exception as e:
        logger.error("[cmd_status] %s", e)
        if update.message:
            await update.message.reply_text(f"Erro: {e}")


async def cmd_bookings(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not context.args:
        if update.message:
            await update.message.reply_text("Uso: /bookings <navio>\nEx: /bookings EAGLE NORTH")
        return

    query = " ".join(context.args).strip().upper()
    try:
        df = _ler_booking(str(_resolve_booking_path()))
        mask = df["VESSEL"].str.contains(query, na=False)
        sub = df[mask]

        if sub.empty:
            if update.message:
                await update.message.reply_text(f"Nenhum booking para: *{query}*", parse_mode=ParseMode.MARKDOWN)
            return

        cols_show = ["Booking Number", "POL", "POD", "Booking Party", "TEU"]
        cols_exist = [c for c in cols_show if c in sub.columns]
        preview = sub[cols_exist].head(20)

        linhas = [f"*BOOKINGS: {query}* ({len(sub)} total)\n"]
        for _, row in preview.iterrows():
            linha_parts = [f"{c}: {row[c]}" for c in cols_exist if str(row[c]).strip() not in ("", "nan")]
            linhas.append("-¢ " + " | ".join(linha_parts))

        if len(sub) > 20:
            linhas.append(f"\n_...e mais {len(sub) - 20} bookings_")

        if update.message:
            await update.message.reply_text("\n".join(linhas), parse_mode=ParseMode.MARKDOWN)
    except Exception as e:
        logger.error("[cmd_bookings] %s", e)
        if update.message:
            await update.message.reply_text(f"Erro: {e}")


async def cmd_feeder(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    del context
    if update.message:
        await update.message.reply_text("Analisando conexoes TIGER -> demo service...", parse_mode=ParseMode.MARKDOWN)
    try:
        df = _ler_booking(str(_resolve_booking_path()))
        conexoes = _analisar_conexoes_tiger(df)
        msg = fmt_conexoes_tiger(conexoes)
        if update.message:
            await update.message.reply_text(msg, parse_mode=ParseMode.MARKDOWN)
    except Exception as e:
        logger.error("[cmd_feeder] %s", e)
        if update.message:
            await update.message.reply_text(f"Erro: {e}")


async def cmd_atualizar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    del context
    if update.message:
        await update.message.reply_text("Relendo BookingReport...")
    try:
        path = _resolve_booking_path()
        df = _ler_booking(str(path))
        if update.message:
            await update.message.reply_text(
                f"OK - {len(df)} bookings carregados de {df['VESSEL'].nunique()} navios.\nArquivo: {path.name}"
            )
    except Exception as e:
        if update.message:
            await update.message.reply_text(f"Erro ao ler arquivo: {e}")


async def cmd_ajuda(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await cmd_start(update, context)


# =============================================================================
# TRIAGEM DE MENSAGENS LIVRES (FAQ)
# =============================================================================

FAQ = {
    ("status", "ocupacao", "cheio", "vazio", "espaco", "livre"): (
        "Use /capacidade para ver ocupacao de todos os navios, ou /navio <nome> para um especifico."
    ),
    ("booking", "reserva", "carga", "container", "teu"): (
        "Use /bookings <navio> para listar bookings. Ex: /bookings MERCURY"
    ),
    ("tiger", "feeder", "conexao", "brrig", "transbordo"): (
        "Use /feeder para analisar conexoes TIGER -> demo service e identificar riscos."
    ),
    ("vessel_one", "vessel_two", "vessel_three", "mercury voyage one", "venus"): (
        "Use /navio <nome> para detalhes. Ex: /navio MERCURY"
    ),
    ("urgente", "critico", "problema", "atencao", "overbook"): (
        "Use /status para ver navios em situacao critica."
    ),
}

ESCALADA_KEYWORDS = [
    "autoriza",
    "aceita",
    "bloqueia",
    "cancela",
    "decisao",
    "aprovar",
    "reprovar",
    "mudar",
    "alterar",
    "confirmar",
]


async def handle_mensagem(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.message.text:
        return

    msg_lower = update.message.text.lower()

    for kw in ESCALADA_KEYWORDS:
        if kw in msg_lower:
            await update.message.reply_text(
                "Esta solicitacao precisa de decisao do operador.\n"
                "Encaminhando para a equipe... Use os comandos /capacidade ou /status "
                "para informacoes que posso fornecer automaticamente."
            )
            for admin_id in ADMIN_CHAT_IDS:
                try:
                    await context.bot.send_message(
                        chat_id=admin_id,
                        text=f"*ESCALADA DE DEMANDA*\nMensagem: _{update.message.text}_\nDe: {update.effective_user.name}",
                        parse_mode=ParseMode.MARKDOWN,
                    )
                except Exception:
                    pass
            return

    for keywords, resposta in FAQ.items():
        if any(kw in msg_lower for kw in keywords):
            await update.message.reply_text(resposta)
            return

    await update.message.reply_text("Nao entendi. Use /ajuda para ver os comandos disponiveis.")


# =============================================================================
# ALERTAS PROATIVOS
# =============================================================================

async def enviar_alerta_capacidade(bot: Bot) -> None:
    try:
        df = _ler_booking(str(_resolve_booking_path()))
        cap = _calcular_capacidade(df)
        alertas = cap[cap["STATUS"].isin(["CRITICO", "OVERBOOKED", "CHEIO"])]

        if alertas.empty:
            return

        linhas = [f"*ALERTA DE CAPACIDADE* | {datetime.now().strftime('%d/%m %H:%M')}\n"]
        for _, row in alertas.iterrows():
            emoji = STATUS_EMOJI.get(row["STATUS"], "ðŸ”´")
            teu_str = (
                f"{row['BOOKED_TEU']}/{row['MAX_TEU']} TEU ({row['PCT_TEU']}%)"
                if row["MAX_TEU"]
                else f"{row['BOOKED_TEU']} TEU"
            )
            linhas.append(f"{emoji} *{row['VESSEL']}/{row['VOYAGE']}* [{row['POL']}]")
            linhas.append(f"   {teu_str} | {row['STATUS']}")

        msg = "\n".join(linhas)
        for chat_id in ADMIN_CHAT_IDS:
            await bot.send_message(chat_id=chat_id, text=msg, parse_mode=ParseMode.MARKDOWN)
            logger.info("[ALERTA] Enviado para chat_id=%s", chat_id)

    except Exception as e:
        logger.error("[enviar_alerta_capacidade] %s", e)


async def enviar_alerta_conexoes(bot: Bot) -> None:
    try:
        df = _ler_booking(str(_resolve_booking_path()))
        conexoes = _analisar_conexoes_tiger(df)
        riscos = [c for c in conexoes if c["risco"] not in ("OK", "SEM DATA")]

        if not riscos:
            return

        linhas = [f"*ALERTA: CONEXOES FEEDER* | {datetime.now().strftime('%d/%m %H:%M')}\n"]
        for c in riscos:
            emoji = RISCO_EMOJI.get(c["risco"], "Warning")
            gap_str = f"{c['gap_dias']} dias" if c["gap_dias"] is not None else "N/A"
            linhas.append(f"{emoji} *{c['feeder']}* -> {c['lam']}")
            linhas.append(f"   Gap: {gap_str} | {c['risco']}")

        msg = "\n".join(linhas)
        for chat_id in ADMIN_CHAT_IDS:
            await bot.send_message(chat_id=chat_id, text=msg, parse_mode=ParseMode.MARKDOWN)

    except Exception as e:
        logger.error("[enviar_alerta_conexoes] %s", e)


def _rodar_alertas_em_loop(app: Application) -> None:
    async def _enviar() -> None:
        await enviar_alerta_capacidade(app.bot)
        await enviar_alerta_conexoes(app.bot)

    def _job() -> None:
        loop = asyncio.new_event_loop()
        try:
            loop.run_until_complete(_enviar())
        finally:
            loop.close()

    for horario in ALERT_HORARIOS:
        schedule.every().day.at(horario).do(_job)
        logger.info("[AGENDA] Alerta agendado para %s", horario)

    while True:
        schedule.run_pending()
        time.sleep(30)


# =============================================================================
# MAIN
# =============================================================================

def main() -> None:
    logger.info("=" * 60)
    logger.info("BOT TELEGRAM MARITIMO - INICIANDO")
    logger.info("=" * 60)

    if not BOT_TOKEN or BOT_TOKEN == "coloque_seu_token_aqui":
        print("\n[ERRO] Configure o TELEGRAM_BOT_TOKEN antes de rodar!")
        print("  1. Copie .env.example para .env dentro da pasta Telegram_Bot")
        print("  2. Fale com @BotFather no Telegram e gere o token")
        print("  3. Cole o token na variavel TELEGRAM_BOT_TOKEN do arquivo .env\n")
        return

    if not ADMIN_CHAT_IDS:
        logger.warning("[AVISO] TELEGRAM_ALLOWED_CHAT_ID nao configurado; alertas proativos ficarao desativados.")

    booking_path = _resolve_booking_path()
    if not booking_path.exists():
        logger.warning("[AVISO] BookingReport nao encontrado em: %s", booking_path)
        logger.warning("O bot vai rodar, mas os comandos vao falhar ate o arquivo existir.")

    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("capacidade", cmd_capacidade))
    app.add_handler(CommandHandler("navio", cmd_navio))
    app.add_handler(CommandHandler("status", cmd_status))
    app.add_handler(CommandHandler("bookings", cmd_bookings))
    app.add_handler(CommandHandler("feeder", cmd_feeder))
    app.add_handler(CommandHandler("atualizar", cmd_atualizar))
    app.add_handler(CommandHandler("ajuda", cmd_ajuda))
    app.add_handler(CommandHandler("schedule", cmd_schedule))
    app.add_handler(CallbackQueryHandler(cb_menu_navio, pattern="^(navio:|menu:inicio|menu:schedule)"))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_mensagem))

    t = threading.Thread(target=_rodar_alertas_em_loop, args=(app,), daemon=True)
    t.start()
    logger.info("[AGENDA] Thread de alertas iniciada")

    logger.info("[BOT] Rodando... Pressione Ctrl+C para parar.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()




