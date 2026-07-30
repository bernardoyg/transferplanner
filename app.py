from __future__ import annotations

import base64
import hashlib
import io
import json
import math
import re
import sqlite3
import unicodedata
from dataclasses import asdict, dataclass, replace
from datetime import date, datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
DB_PATH = DATA_DIR / "pallet_planner.db"
INITIAL_BASE_PATH = DATA_DIR / "base_inicial.xlsx"
ASSETS_DIR = APP_DIR / "assets"
LOGO_PATH = ASSETS_DIR / "bbm-logistica-logo.png"
FAVICON_PATH = ASSETS_DIR / "bbm-favicon.png"
EXPECTED_COLUMNS = [
    "Data de expedição",
    "Volumes",
    1,
    2,
    3,
    4,
    5,
    "P1",
    "P2",
    "P3",
    "P4",
    "P5",
]
DESTINATION_COLUMNS = [1, 2, 3, 4, 5]
PALLET_COLUMNS = ["P1", "P2", "P3", "P4", "P5"]
WEEKDAY_NAMES = {
    0: "segunda-feira",
    1: "terça-feira",
    2: "quarta-feira",
    3: "quinta-feira",
    4: "sexta-feira",
    5: "sábado",
    6: "domingo",
}
PLOTLY_CONFIG = {
    "scrollZoom": True,
    "displayModeBar": True,
    "displaylogo": False,
    "responsive": True,
}


def image_data_uri(path: Path) -> str:
    if not path.exists():
        return ""
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def apply_brand_theme() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Urbanist:wght@400;500;600;700&display=swap');

        :root {
            --bbm-orange: #E85D04;
            --bbm-orange-light: #FFF0E6;
            --bbm-navy: #1A2744;
            --bbm-blue: #185FA5;
            --bbm-teal: #0B7A75;
            --bbm-surface: #FFFFFF;
            --bbm-background: #F5F7FB;
            --bbm-border: #E2E8F0;
        }

        html, body, [class*="css"], [data-testid="stAppViewContainer"],
        [data-testid="stSidebar"], button, input, textarea, select {
            font-family: 'Urbanist', Arial, sans-serif !important;
        }

        [data-testid="stAppViewContainer"] {
            background: var(--bbm-background);
        }

        [data-testid="stHeader"] {
            background: rgba(245, 247, 251, 0.92);
        }

        [data-testid="stSidebar"] {
            background: #FFFFFF;
            border-right: 1px solid var(--bbm-border);
        }

        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
        [data-testid="stSidebar"] label {
            color: var(--bbm-navy);
        }

        .brand-header {
            display: flex;
            align-items: center;
            gap: 1.5rem;
            padding: 1.4rem 1.6rem;
            margin: 0 0 1.4rem;
            border-radius: 18px;
            background: linear-gradient(135deg, var(--bbm-navy) 0%, #24375F 100%);
            box-shadow: 0 12px 28px rgba(26, 39, 68, 0.14);
        }

        .brand-logo {
            display: flex;
            align-items: center;
            justify-content: center;
            min-width: 230px;
            padding: 0.9rem 1.1rem;
            border-radius: 12px;
            background: #FFFFFF;
        }

        .brand-logo img {
            width: 200px;
            max-width: 100%;
            height: auto;
        }

        .brand-copy h1 {
            margin: 0;
            color: #FFFFFF;
            font-size: 2rem;
            line-height: 1.1;
            font-weight: 700;
            letter-spacing: -0.02em;
        }

        .brand-copy p {
            margin: 0.45rem 0 0;
            color: #DCE5F5;
            font-size: 1rem;
            font-weight: 500;
        }

        .sidebar-app-name {
            color: var(--bbm-navy);
            font-size: 1.15rem;
            font-weight: 700;
            padding: 0.2rem 0 0.65rem;
        }

        div[data-testid="stMetric"] {
            background: var(--bbm-surface);
            border: 1px solid var(--bbm-border);
            border-top: 4px solid var(--bbm-orange);
            border-radius: 14px;
            padding: 1rem 1.1rem;
            box-shadow: 0 5px 16px rgba(26, 39, 68, 0.06);
        }

        div[data-testid="stMetricLabel"] {
            color: #52617A;
            font-weight: 600;
        }

        div[data-testid="stMetricValue"] {
            color: var(--bbm-navy);
            font-weight: 700;
        }

        .stButton > button,
        .stDownloadButton > button {
            border-radius: 9px;
            font-weight: 600;
            border-color: var(--bbm-orange);
        }

        .stButton > button[kind="primary"],
        .stDownloadButton > button:hover {
            background: var(--bbm-orange);
            color: #FFFFFF;
        }

        button[data-baseweb="tab"][aria-selected="true"] {
            color: var(--bbm-navy);
            font-weight: 700;
        }

        button[data-baseweb="tab"][aria-selected="true"] div[data-testid="stMarkdownContainer"] p {
            color: var(--bbm-navy);
        }

        [data-baseweb="tab-highlight"] {
            background-color: var(--bbm-orange);
        }

        div[data-testid="stExpander"] {
            border: 1px solid var(--bbm-border);
            border-radius: 12px;
            background: #FFFFFF;
        }

        div[data-testid="stDataFrame"] {
            border: 1px solid var(--bbm-border);
            border-radius: 12px;
            overflow: hidden;
        }

        .dimension-highlight {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1rem;
            margin: 1rem 0 1.15rem;
            padding: 1rem 1.2rem;
            border: 1px solid #F4B183;
            border-left: 6px solid var(--bbm-orange);
            border-radius: 12px;
            background: linear-gradient(90deg, var(--bbm-orange-light), #FFFFFF);
        }

        .dimension-highlight span {
            display: block;
            color: #9A3E00;
            font-size: 0.76rem;
            font-weight: 700;
            letter-spacing: 0.08em;
            text-transform: uppercase;
        }

        .dimension-highlight strong {
            display: block;
            margin-top: 0.2rem;
            color: var(--bbm-navy);
            font-size: 1.15rem;
        }

        .dimension-highlight .dimension-capacity {
            color: #52617A;
            font-size: 0.92rem;
            font-weight: 600;
            white-space: nowrap;
        }

        @media (max-width: 760px) {
            .brand-header {
                align-items: flex-start;
                flex-direction: column;
                gap: 1rem;
                padding: 1.2rem;
            }

            .brand-logo {
                min-width: 0;
                width: 100%;
                justify-content: flex-start;
            }

            .brand-logo img {
                width: 175px;
            }

            .brand-copy h1 {
                font-size: 1.65rem;
            }

            .dimension-highlight {
                align-items: flex-start;
                flex-direction: column;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def page_header(title: str, subtitle: str) -> None:
    logo_uri = image_data_uri(LOGO_PATH)
    logo_html = (
        f'<div class="brand-logo"><img src="{logo_uri}" alt="BBM Logística"></div>'
        if logo_uri
        else ""
    )
    st.markdown(
        f"""
        <div class="brand-header">
            {logo_html}
            <div class="brand-copy">
                <h1>{title}</h1>
                <p>{subtitle}</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


PARAMETER_HELP = {
    "history_weeks": (
        "Quantidade de semanas anteriores disponíveis para procurar dias comparáveis "
        "e calcular as médias por praça."
    ),
    "initial_band": (
        "Primeira tolerância de volume usada para procurar dias semelhantes. "
        "Com 15%, uma previsão de 70.000 busca inicialmente dias entre 59.500 e 80.500 volumes."
    ),
    "expanded_band": (
        "Tolerância usada quando a faixa inicial não encontra a quantidade mínima "
        "de dias comparáveis."
    ),
    "minimum_comparables": (
        "Quantidade mínima desejada de dias para calcular a média. Se não for atingida, "
        "o modelo amplia a faixa e pode usar todos os registros do mesmo dia da semana."
    ),
    "recency_half_life": (
        "Define a velocidade de perda de peso dos dias antigos. Em 21 dias, um registro "
        "passa a valer metade de um registro recente."
    ),
    "recent_curve": (
        "Período usado para verificar como o modelo teria performado recentemente e "
        "corrigir tendências de aumento ou redução de pallets."
    ),
    "calibration_half_life": (
        "Define quanto os erros mais recentes pesam na curva de correção. Valores menores "
        "tornam o ajuste mais sensível aos últimos dias."
    ),
    "capacity_error_percentile": (
        "Percentil do erro recente acrescentado aos pallets esperados para formar a "
        "capacidade recomendada. Percentis maiores geram mais proteção operacional."
    ),
}


def render_model_parameter_guide() -> None:
    st.markdown(
        """
        #### Como interpretar os parâmetros

        | Parâmetro | Função no cálculo | Ao aumentar |
        |---|---|---|
        | **Janela histórica** | Limita quantas semanas anteriores entram na busca e nas médias por praça. | Usa mais histórico e deixa o modelo mais estável, porém menos sensível a mudanças recentes. |
        | **Faixa inicial** | Define a tolerância de volume para selecionar dias semelhantes. | Inclui mais dias, mas com volumes menos parecidos. |
        | **Faixa ampliada** | Segunda tolerância, utilizada quando a faixa inicial não fornece amostra suficiente. | Reduz a chance de faltar histórico comparável. |
        | **Mínimo de dias comparáveis** | Determina o tamanho mínimo desejado da amostra. | Exige uma base mais ampla e pode acionar a faixa ampliada com maior frequência. |
        | **Meia-vida dos comparáveis** | Controla a perda de peso dos registros antigos. | Dá mais influência ao histórico antigo e suaviza oscilações recentes. |
        | **Curva recente** | Define quantos dias são usados para recalibrar o modelo pelos erros mais atuais. | Avalia um período maior e produz uma correção mais estável. |
        | **Meia-vida da curva recente** | Controla o peso dos erros antigos dentro da curva recente. | Distribui o peso por mais dias e reduz a reação aos últimos resultados. |
        | **Margem pelo erro recente** | Escolhe o percentil do erro histórico acrescentado à capacidade recomendada. | Aumenta a proteção contra falta de capacidade, mas pode elevar a sobra planejada. |
        """
    )
    st.caption(
        "Referência recomendada: mantenha os valores padrão até acumular novos desvios "
        "no Previsto × realizado. A média por praça também recebe uma estabilização "
        "interna equivalente a 20 pallets da média geral, evitando distorções em "
        "unidades com pouco histórico."
    )


@dataclass
class SimulationParameters:
    history_weeks: int = 16
    initial_band: float = 0.15
    expanded_band: float = 0.25
    minimum_comparables: int = 4
    recency_half_life_days: int = 21
    capacity_percentile: float = 0.25
    high_density_percentile: float = 0.90
    recent_calibration_days: int = 14
    calibration_half_life_days: int = 7
    calibration_minimum_days: int = 4
    capacity_error_percentile: float = 0.80
    destination_prior_pallets: int = 20


@dataclass
class SimulationResult:
    operation_date: date
    forecast_volumes: int
    expected_pallets: int
    capacity_pallets: int
    high_density_pallets: int
    expected_density: float
    capacity_density: float
    high_density: float
    comparable_count: int
    selection_rule: str
    comparable_dates: list[str]
    base_density: float
    calibration_factor: float
    calibration_days: int
    capacity_buffer: int


def get_connection() -> sqlite3.Connection:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DB_PATH)
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def initialize_database() -> None:
    with get_connection() as connection:
        connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS loads (
                load_id TEXT PRIMARY KEY,
                operation_date TEXT NOT NULL,
                route TEXT,
                volumes REAL NOT NULL,
                total_pallets REAL NOT NULL,
                source_name TEXT,
                imported_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS load_destinations (
                load_id TEXT NOT NULL,
                destination TEXT NOT NULL,
                pallets REAL NOT NULL,
                PRIMARY KEY (load_id, destination),
                FOREIGN KEY (load_id) REFERENCES loads(load_id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS forecasts (
                forecast_id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                operation_date TEXT NOT NULL,
                forecast_volumes INTEGER NOT NULL,
                expected_pallets INTEGER NOT NULL,
                capacity_pallets INTEGER NOT NULL,
                high_density_pallets INTEGER NOT NULL,
                expected_density REAL NOT NULL,
                comparable_count INTEGER NOT NULL,
                selection_rule TEXT NOT NULL,
                parameters_json TEXT NOT NULL
            );
            """
        )


def normalize_header(value: object) -> object:
    if isinstance(value, str):
        return " ".join(value.replace("\n", " ").split())
    return value


def normalize_destination(value: object) -> str:
    if pd.isna(value) or str(value).strip() == "":
        return ""
    raw = unicodedata.normalize("NFKC", str(value)).upper()
    raw = re.sub(r"[\u200B-\u200D\u2060\uFEFF]", "", raw)
    raw = " ".join(raw.strip().split())
    tokens = [token for token in re.split(r"[\s/\-]+", raw) if token]
    tokens = ["FL2" if token == "LF2" else token for token in tokens]
    token_set = set(tokens)
    if raw == "CWBCW2" or (tokens and token_set <= {"CWB", "CW2"}):
        return "CWB/CW2"
    if tokens and token_set <= {"FLI", "FL2"}:
        return "FLI/FL2"
    return "/".join(tokens)


def read_uploaded_base(file_or_path: object) -> pd.DataFrame:
    dataframe = pd.read_excel(file_or_path, sheet_name=0)
    dataframe.columns = [normalize_header(column) for column in dataframe.columns]
    missing = [column for column in EXPECTED_COLUMNS if column not in dataframe.columns]
    if missing:
        raise ValueError(f"Colunas obrigatórias não encontradas: {missing}")
    return dataframe


def numeric_value(value: object) -> float:
    number = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    return float(number) if pd.notna(number) else 0.0


def dataframe_to_records(dataframe: pd.DataFrame, source_name: str) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    loads: list[dict] = []
    destinations: list[dict] = []
    invalid_rows = 0
    pallet_without_destination = 0
    destination_without_pallet = 0
    imported_at = datetime.now().isoformat(timespec="seconds")
    identity_occurrences: dict[str, int] = {}

    for row_number, row in dataframe.iterrows():
        operation_date = pd.to_datetime(row.get("Data de expedição"), errors="coerce")
        volumes = numeric_value(row.get("Volumes"))
        pallet_values = [numeric_value(row.get(column)) for column in PALLET_COLUMNS]
        total_pallets = sum(value for value in pallet_values if value > 0)
        destination_values = [normalize_destination(row.get(column)) for column in DESTINATION_COLUMNS]

        if pd.isna(operation_date) or volumes <= 0 or total_pallets <= 0:
            invalid_rows += 1
            continue

        pairs: dict[str, float] = {}
        for destination, pallets in zip(destination_values, pallet_values):
            if destination and pallets > 0:
                pairs[destination] = pairs.get(destination, 0.0) + pallets
            elif destination and pallets <= 0:
                destination_without_pallet += 1
            elif not destination and pallets > 0:
                pallet_without_destination += 1

        if not pairs:
            invalid_rows += 1
            continue

        route = "" if pd.isna(row.get("Rota")) else str(row.get("Rota")).strip()
        sequence = "" if pd.isna(row.get("S")) else str(row.get("S")).strip()
        typology = "" if pd.isna(row.get("Tipologia")) else str(row.get("Tipologia")).strip()
        capacity = "" if pd.isna(row.get("Capacidade")) else str(row.get("Capacidade")).strip()
        identity_parts = [operation_date.date().isoformat(), sequence, route, typology, capacity]
        if not sequence and not route:
            identity_parts.extend(
                [
                    f"{volumes:.4f}",
                    f"{total_pallets:.4f}",
                    json.dumps(pairs, sort_keys=True, ensure_ascii=False),
                ]
            )
        identity_base = "|".join(identity_parts)
        occurrence = identity_occurrences.get(identity_base, 0) + 1
        identity_occurrences[identity_base] = occurrence
        signature = f"{identity_base}|{occurrence}"
        load_id = hashlib.sha256(signature.encode("utf-8")).hexdigest()
        loads.append(
            {
                "load_id": load_id,
                "operation_date": operation_date.date().isoformat(),
                "route": route,
                "volumes": volumes,
                "total_pallets": total_pallets,
                "source_name": source_name,
                "imported_at": imported_at,
            }
        )
        for destination, pallets in pairs.items():
            destinations.append({"load_id": load_id, "destination": destination, "pallets": pallets})

    summary = {
        "valid_loads": len(loads),
        "invalid_rows": invalid_rows,
        "destination_without_pallet": destination_without_pallet,
        "pallet_without_destination": pallet_without_destination,
    }
    return pd.DataFrame(loads), pd.DataFrame(destinations), summary


def insert_records(loads: pd.DataFrame, destinations: pd.DataFrame) -> tuple[int, int, int]:
    if loads.empty:
        return 0, 0, 0
    inserted = 0
    updated = 0
    unchanged = 0
    with get_connection() as connection:
        for record in loads.to_dict("records"):
            existing = connection.execute(
                "SELECT route, volumes, total_pallets FROM loads WHERE load_id = ?",
                (record["load_id"],),
            ).fetchone()
            destination_rows = destinations[destinations["load_id"] == record["load_id"]]
            incoming_destinations = {
                row["destination"]: float(row["pallets"])
                for row in destination_rows.to_dict("records")
            }
            if existing is None:
                connection.execute(
                    """
                    INSERT INTO loads
                    (load_id, operation_date, route, volumes, total_pallets, source_name, imported_at)
                    VALUES (:load_id, :operation_date, :route, :volumes, :total_pallets, :source_name, :imported_at)
                    """,
                    record,
                )
                inserted += 1
                connection.executemany(
                    """
                    INSERT OR REPLACE INTO load_destinations (load_id, destination, pallets)
                    VALUES (:load_id, :destination, :pallets)
                    """,
                    destination_rows.to_dict("records"),
                )
            else:
                existing_destinations = dict(
                    connection.execute(
                        "SELECT destination, pallets FROM load_destinations WHERE load_id = ?",
                        (record["load_id"],),
                    ).fetchall()
                )
                changed = (
                    str(existing[0] or "") != str(record["route"] or "")
                    or not math.isclose(float(existing[1]), float(record["volumes"]))
                    or not math.isclose(float(existing[2]), float(record["total_pallets"]))
                    or existing_destinations != incoming_destinations
                )
                if changed:
                    connection.execute(
                        """
                        UPDATE loads SET route = :route, volumes = :volumes,
                        total_pallets = :total_pallets, source_name = :source_name,
                        imported_at = :imported_at WHERE load_id = :load_id
                        """,
                        record,
                    )
                    connection.execute(
                        "DELETE FROM load_destinations WHERE load_id = ?",
                        (record["load_id"],),
                    )
                    connection.executemany(
                        """
                        INSERT INTO load_destinations (load_id, destination, pallets)
                        VALUES (:load_id, :destination, :pallets)
                        """,
                        destination_rows.to_dict("records"),
                    )
                    updated += 1
                else:
                    unchanged += 1
    return inserted, updated, unchanged


def seed_initial_history() -> None:
    with get_connection() as connection:
        count = connection.execute("SELECT COUNT(*) FROM loads").fetchone()[0]
    if count == 0 and INITIAL_BASE_PATH.exists():
        dataframe = read_uploaded_base(INITIAL_BASE_PATH)
        loads, destinations, _ = dataframe_to_records(dataframe, INITIAL_BASE_PATH.name)
        insert_records(loads, destinations)


def normalize_saved_destinations() -> tuple[int, int]:
    with get_connection() as connection:
        rows = connection.execute(
            "SELECT load_id, destination, pallets FROM load_destinations"
        ).fetchall()
        if not rows:
            return 0, 0

        consolidated: dict[tuple[str, str], float] = {}
        renamed = 0
        for load_id, destination, pallets in rows:
            normalized = normalize_destination(destination)
            if not normalized:
                continue
            if normalized != destination:
                renamed += 1
            key = (load_id, normalized)
            consolidated[key] = consolidated.get(key, 0.0) + float(pallets)

        merged = len(rows) - len(consolidated)
        if renamed == 0 and merged == 0:
            return 0, 0

        original_total = sum(float(row[2]) for row in rows)
        consolidated_total = sum(consolidated.values())
        if not math.isclose(original_total, consolidated_total):
            raise ValueError(
                "A normalização das praças alteraria o total de pallets e foi cancelada."
            )

        connection.execute("DELETE FROM load_destinations")
        connection.executemany(
            """
            INSERT INTO load_destinations (load_id, destination, pallets)
            VALUES (?, ?, ?)
            """,
            [
                (load_id, destination, pallets)
                for (load_id, destination), pallets in consolidated.items()
            ],
        )
        return renamed, merged


@st.cache_data(show_spinner=False)
def load_history(db_mtime: float) -> tuple[pd.DataFrame, pd.DataFrame]:
    del db_mtime
    with get_connection() as connection:
        loads = pd.read_sql_query("SELECT * FROM loads", connection)
        destinations = pd.read_sql_query("SELECT * FROM load_destinations", connection)
    if not loads.empty:
        loads["operation_date"] = pd.to_datetime(loads["operation_date"]).dt.date
    return loads, destinations


def get_history() -> tuple[pd.DataFrame, pd.DataFrame]:
    mtime = DB_PATH.stat().st_mtime if DB_PATH.exists() else 0
    return load_history(mtime)


def build_daily_history(loads: pd.DataFrame, operation_date: date, history_weeks: int) -> pd.DataFrame:
    if loads.empty:
        return pd.DataFrame()
    start_date = operation_date - timedelta(weeks=history_weeks)
    eligible = loads[
        (loads["operation_date"] < operation_date)
        & (loads["operation_date"] >= start_date)
    ].copy()
    if eligible.empty:
        return pd.DataFrame()
    daily = (
        eligible.groupby("operation_date", as_index=False)
        .agg(volumes=("volumes", "sum"), pallets=("total_pallets", "sum"), loads=("load_id", "nunique"))
    )
    daily["weekday"] = pd.to_datetime(daily["operation_date"]).dt.weekday
    daily["density"] = daily["volumes"] / daily["pallets"]
    return daily.sort_values("operation_date")


def build_daily_pool(loads: pd.DataFrame, operation_date: date) -> pd.DataFrame:
    if loads.empty:
        return pd.DataFrame()
    eligible = loads[loads["operation_date"] < operation_date].copy()
    if eligible.empty:
        return pd.DataFrame()
    daily = (
        eligible.groupby("operation_date", as_index=False)
        .agg(
            volumes=("volumes", "sum"),
            pallets=("total_pallets", "sum"),
            loads=("load_id", "nunique"),
        )
    )
    daily["weekday"] = pd.to_datetime(daily["operation_date"]).dt.weekday
    daily["density"] = daily["volumes"] / daily["pallets"]
    return daily.sort_values("operation_date")


def build_destination_daily_history(
    loads: pd.DataFrame,
    destinations: pd.DataFrame,
    operation_date: date,
    destination: str,
) -> pd.DataFrame:
    eligible_loads = loads[
        (loads["operation_date"] < operation_date)
        & (loads["total_pallets"] > 0)
    ][["load_id", "operation_date", "volumes", "total_pallets"]].copy()
    selected_destinations = destinations[
        destinations["destination"] == destination
    ][["load_id", "pallets"]].copy()
    merged = selected_destinations.merge(eligible_loads, on="load_id", how="inner")
    merged = merged[merged["pallets"] > 0].copy()
    if merged.empty:
        return pd.DataFrame()

    merged["estimated_volumes"] = (
        merged["volumes"] * merged["pallets"] / merged["total_pallets"]
    )
    daily = (
        merged.groupby("operation_date", as_index=False)
        .agg(
            volumes=("estimated_volumes", "sum"),
            pallets=("pallets", "sum"),
            loads=("load_id", "nunique"),
        )
        .sort_values("operation_date")
    )
    daily["weekday"] = pd.to_datetime(daily["operation_date"]).dt.weekday
    daily["density"] = daily["volumes"] / daily["pallets"]
    return daily


def select_comparables(
    daily: pd.DataFrame,
    operation_date: date,
    forecast_volumes: int,
    parameters: SimulationParameters,
) -> tuple[pd.DataFrame, str]:
    same_weekday = daily[daily["weekday"] == operation_date.weekday()].copy()
    lower = forecast_volumes * (1 - parameters.initial_band)
    upper = forecast_volumes * (1 + parameters.initial_band)
    selected = same_weekday[same_weekday["volumes"].between(lower, upper)]
    rule = f"mesmo dia da semana e faixa de ±{parameters.initial_band:.0%}"

    if len(selected) < parameters.minimum_comparables:
        lower = forecast_volumes * (1 - parameters.expanded_band)
        upper = forecast_volumes * (1 + parameters.expanded_band)
        selected = same_weekday[same_weekday["volumes"].between(lower, upper)]
        rule = f"mesmo dia da semana e faixa ampliada de ±{parameters.expanded_band:.0%}"

    if len(selected) < parameters.minimum_comparables:
        selected = same_weekday
        rule = "todos os registros do mesmo dia da semana"

    if selected.empty:
        selected = daily
        rule = "histórico geral disponível"

    return selected.copy(), rule


def weighted_density(comparables: pd.DataFrame, reference_date: date, half_life_days: int) -> float:
    ages = comparables["operation_date"].apply(lambda value: max(0, (reference_date - value).days))
    weights = np.power(0.5, ages / half_life_days)
    return float((weights * comparables["volumes"]).sum() / (weights * comparables["pallets"]).sum())


def base_estimate_from_daily(
    daily_pool: pd.DataFrame,
    operation_date: date,
    forecast_volumes: int,
    parameters: SimulationParameters,
) -> tuple[float, pd.DataFrame, str]:
    start_date = operation_date - timedelta(weeks=parameters.history_weeks)
    daily = daily_pool[
        (daily_pool["operation_date"] < operation_date)
        & (daily_pool["operation_date"] >= start_date)
    ].copy()
    if daily.empty:
        raise ValueError("Não há histórico anterior suficiente para a data selecionada.")
    comparables, rule = select_comparables(
        daily,
        operation_date,
        forecast_volumes,
        parameters,
    )
    reference_date = max(comparables["operation_date"])
    density = weighted_density(
        comparables,
        reference_date,
        parameters.recency_half_life_days,
    )
    return density, comparables, rule


def recent_calibration(
    daily_pool: pd.DataFrame,
    operation_date: date,
    parameters: SimulationParameters,
) -> tuple[float, int, int]:
    recent_start = operation_date - timedelta(days=parameters.recent_calibration_days)
    recent = daily_pool[
        (daily_pool["operation_date"] < operation_date)
        & (daily_pool["operation_date"] >= recent_start)
    ].copy()
    performance_rows: list[dict] = []
    for row in recent.itertuples(index=False):
        try:
            density, _, _ = base_estimate_from_daily(
                daily_pool,
                row.operation_date,
                int(row.volumes),
                parameters,
            )
        except ValueError:
            continue
        performance_rows.append(
            {
                "operation_date": row.operation_date,
                "actual_pallets": float(row.pallets),
                "base_pallets": float(row.volumes) / density,
            }
        )

    performance = pd.DataFrame(performance_rows)
    if len(performance) < parameters.calibration_minimum_days:
        return 1.0, len(performance), 0

    ages = performance["operation_date"].apply(
        lambda value: max(0, (operation_date - value).days)
    )
    weights = np.power(0.5, ages / parameters.calibration_half_life_days)
    factor = float(
        (weights * performance["actual_pallets"]).sum()
        / (weights * performance["base_pallets"]).sum()
    )
    factor = float(np.clip(factor, 0.80, 1.20))
    adjusted = performance["base_pallets"] * factor
    absolute_errors = (performance["actual_pallets"] - adjusted).abs()
    capacity_buffer = math.ceil(
        float(
            np.percentile(
                absolute_errors,
                parameters.capacity_error_percentile * 100,
            )
        )
    )
    return factor, len(performance), capacity_buffer


def run_simulation(
    loads: pd.DataFrame,
    operation_date: date,
    forecast_volumes: int,
    parameters: SimulationParameters,
) -> tuple[SimulationResult, pd.DataFrame, pd.DataFrame]:
    daily_pool = build_daily_pool(loads, operation_date)
    if daily_pool.empty:
        raise ValueError("Não há histórico anterior suficiente para a data selecionada.")
    base_density, comparables, rule = base_estimate_from_daily(
        daily_pool,
        operation_date,
        forecast_volumes,
        parameters,
    )
    calibration_factor, calibration_days, capacity_buffer = recent_calibration(
        daily_pool,
        operation_date,
        parameters,
    )
    base_pallets = forecast_volumes / base_density
    adjusted_pallets = base_pallets * calibration_factor
    expected_pallets = math.ceil(adjusted_pallets)
    capacity_pallets = math.ceil(adjusted_pallets + capacity_buffer)
    high_density_pallets = max(1, math.ceil(adjusted_pallets - capacity_buffer))
    expected_density = forecast_volumes / adjusted_pallets
    capacity_density = forecast_volumes / capacity_pallets
    high_density = forecast_volumes / high_density_pallets
    daily = daily_pool[
        daily_pool["operation_date"]
        >= operation_date - timedelta(weeks=parameters.history_weeks)
    ].copy()
    if calibration_days >= parameters.calibration_minimum_days:
        rule = (
            f"{rule}; ajuste recente de {format_percent_br(calibration_factor, 1)} "
            f"com {calibration_days} dias"
        )

    result = SimulationResult(
        operation_date=operation_date,
        forecast_volumes=forecast_volumes,
        expected_pallets=expected_pallets,
        capacity_pallets=capacity_pallets,
        high_density_pallets=high_density_pallets,
        expected_density=expected_density,
        capacity_density=capacity_density,
        high_density=high_density,
        comparable_count=len(comparables),
        selection_rule=rule,
        comparable_dates=[value.isoformat() for value in comparables["operation_date"]],
        base_density=base_density,
        calibration_factor=calibration_factor,
        calibration_days=calibration_days,
        capacity_buffer=capacity_buffer,
    )
    return result, comparables, daily


def largest_remainder_allocation(total: int, shares: pd.Series) -> pd.Series:
    if shares.empty or shares.sum() <= 0:
        return pd.Series(dtype=int)
    raw = total * shares / shares.sum()
    allocation = np.floor(raw).astype(int)
    remaining = total - int(allocation.sum())
    if remaining > 0:
        order = (raw - allocation).sort_values(ascending=False).index[:remaining]
        allocation.loc[order] += 1
    return allocation


def destination_density_estimates(
    loads: pd.DataFrame,
    destinations: pd.DataFrame,
    operation_date: date,
    parameters: SimulationParameters,
    global_density: float,
) -> pd.DataFrame:
    start_date = operation_date - timedelta(weeks=parameters.history_weeks)
    eligible_loads = loads[
        (loads["operation_date"] < operation_date)
        & (loads["operation_date"] >= start_date)
    ][["load_id", "operation_date", "volumes", "total_pallets"]].copy()
    if eligible_loads.empty:
        return pd.DataFrame(
            columns=["Praça", "Densidade histórica da praça", "Cargas históricas"]
        )

    eligible_loads["load_density"] = (
        eligible_loads["volumes"] / eligible_loads["total_pallets"]
    )
    eligible_loads["age_days"] = eligible_loads["operation_date"].apply(
        lambda value: max(0, (operation_date - value).days)
    )
    eligible_loads["recency_weight"] = np.power(
        0.5,
        eligible_loads["age_days"] / parameters.recency_half_life_days,
    )
    merged = destinations.merge(eligible_loads, on="load_id", how="inner")
    if merged.empty:
        return pd.DataFrame(
            columns=["Praça", "Densidade histórica da praça", "Cargas históricas"]
        )

    merged["weighted_pallets"] = merged["pallets"] * merged["recency_weight"]
    merged["weighted_density"] = (
        merged["weighted_pallets"] * merged["load_density"]
    )
    grouped = (
        merged.groupby("destination", as_index=False)
        .agg(
            weighted_pallets=("weighted_pallets", "sum"),
            weighted_density=("weighted_density", "sum"),
            historical_loads=("load_id", "nunique"),
        )
        .rename(columns={"destination": "Praça"})
    )
    prior = float(parameters.destination_prior_pallets)
    grouped["Densidade histórica da praça"] = (
        grouped["weighted_density"] + prior * global_density
    ) / (grouped["weighted_pallets"] + prior)
    return grouped[
        ["Praça", "Densidade histórica da praça", "historical_loads"]
    ].rename(columns={"historical_loads": "Cargas históricas"})


def destination_distribution(
    loads: pd.DataFrame,
    destinations: pd.DataFrame,
    comparables: pd.DataFrame,
    result: SimulationResult,
    parameters: SimulationParameters,
) -> pd.DataFrame:
    comparable_dates = set(comparables["operation_date"])
    comparable_loads = loads[loads["operation_date"].isin(comparable_dates)][["load_id", "operation_date"]]
    merged = destinations.merge(comparable_loads, on="load_id", how="inner")
    if merged.empty:
        return pd.DataFrame()
    reference_date = max(comparable_dates)
    merged["age_days"] = merged["operation_date"].apply(lambda value: max(0, (reference_date - value).days))
    merged["weight"] = np.power(0.5, merged["age_days"] / parameters.recency_half_life_days)
    merged["weighted_pallets"] = merged["pallets"] * merged["weight"]
    shares = merged.groupby("destination")["weighted_pallets"].sum().sort_values(ascending=False)
    expected = largest_remainder_allocation(result.expected_pallets, shares)
    capacity = largest_remainder_allocation(result.capacity_pallets, shares)
    output = pd.DataFrame(
        {
            "Praça": shares.index,
            "Participação": shares / shares.sum(),
            "Pallets esperados": expected,
            "Capacidade recomendada": capacity,
        }
    ).reset_index(drop=True)
    density_estimates = destination_density_estimates(
        loads,
        destinations,
        result.operation_date,
        parameters,
        result.expected_density,
    )
    output = output.merge(density_estimates, on="Praça", how="left")
    output["Densidade histórica da praça"] = output[
        "Densidade histórica da praça"
    ].fillna(result.expected_density)
    output["Cargas históricas"] = output["Cargas históricas"].fillna(0).astype(int)

    raw_destination_volumes = (
        output["Pallets esperados"] * output["Densidade histórica da praça"]
    )
    if raw_destination_volumes.sum() > 0:
        estimated_volumes = largest_remainder_allocation(
            result.forecast_volumes,
            raw_destination_volumes,
        )
        reconciliation_factor = (
            result.forecast_volumes / raw_destination_volumes.sum()
        )
    else:
        estimated_volumes = pd.Series(0, index=output.index, dtype=int)
        reconciliation_factor = 1.0

    output["Volumes estimados"] = estimated_volumes
    output["Média vol./pallet da praça"] = (
        output["Densidade histórica da praça"] * reconciliation_factor
    )
    positive_pallets = output["Pallets esperados"] > 0
    output.loc[positive_pallets, "Média vol./pallet da praça"] = (
        output.loc[positive_pallets, "Volumes estimados"]
        / output.loc[positive_pallets, "Pallets esperados"]
    )
    output = output.drop(columns=["Densidade histórica da praça"])
    return output


def save_forecast(result: SimulationResult, parameters: SimulationParameters) -> None:
    with get_connection() as connection:
        connection.execute(
            """
            INSERT INTO forecasts (
                created_at, operation_date, forecast_volumes, expected_pallets,
                capacity_pallets, high_density_pallets, expected_density,
                comparable_count, selection_rule, parameters_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                datetime.now().isoformat(timespec="seconds"),
                result.operation_date.isoformat(),
                result.forecast_volumes,
                result.expected_pallets,
                result.capacity_pallets,
                result.high_density_pallets,
                result.expected_density,
                result.comparable_count,
                result.selection_rule,
                json.dumps(asdict(parameters), ensure_ascii=False),
            ),
        )


def load_forecasts() -> pd.DataFrame:
    with get_connection() as connection:
        dataframe = pd.read_sql_query("SELECT * FROM forecasts ORDER BY operation_date DESC, forecast_id DESC", connection)
    if not dataframe.empty:
        dataframe["operation_date"] = pd.to_datetime(dataframe["operation_date"]).dt.date
        dataframe["created_at"] = pd.to_datetime(dataframe["created_at"], errors="coerce")
    return dataframe


def clear_forecasts() -> int:
    with get_connection() as connection:
        cursor = connection.execute("DELETE FROM forecasts")
        return max(cursor.rowcount, 0)


def save_forecast_and_open_dashboard(
    result: SimulationResult,
    parameters: SimulationParameters,
) -> None:
    save_forecast(result, parameters)
    st.session_state["dashboard_date"] = result.operation_date
    st.session_state["forecast_saved_notice"] = (
        f"Previsão de {format_date_br(result.operation_date)} salva com sucesso."
    )
    st.session_state["navigation"] = "Painel diário"


def actual_for_date(loads: pd.DataFrame, destinations: pd.DataFrame, target_date: date) -> tuple[dict, pd.Series]:
    actual_loads = loads[loads["operation_date"] == target_date]
    if actual_loads.empty:
        return {}, pd.Series(dtype=float)
    load_ids = actual_loads["load_id"]
    actual_destinations = destinations[destinations["load_id"].isin(load_ids)]
    totals = {
        "volumes": float(actual_loads["volumes"].sum()),
        "pallets": float(actual_loads["total_pallets"].sum()),
    }
    totals["density"] = totals["volumes"] / totals["pallets"] if totals["pallets"] else 0
    shares = actual_destinations.groupby("destination")["pallets"].sum().sort_values(ascending=False)
    return totals, shares


def format_number_br(
    value: object,
    decimals: int = 0,
    show_sign: bool = False,
) -> str:
    if value is None or pd.isna(value):
        return "—"
    sign = "+" if show_sign else ""
    formatted = f"{float(value):{sign},.{decimals}f}"
    return (
        formatted.replace(",", "\uFFFF")
        .replace(".", ",")
        .replace("\uFFFF", ".")
    )


def format_integer(value: float | int) -> str:
    return format_number_br(value, decimals=0)


def format_percent_br(value: object, decimals: int = 1, show_sign: bool = False) -> str:
    if value is None or pd.isna(value):
        return "—"
    return f"{format_number_br(float(value) * 100, decimals, show_sign)}%"


def format_date_br(value: object) -> str:
    if value is None or pd.isna(value):
        return "—"
    parsed = pd.to_datetime(value, errors="coerce")
    return parsed.strftime("%d/%m/%Y") if pd.notna(parsed) else str(value)


def format_table_br(
    dataframe: pd.DataFrame,
    *,
    date_columns: tuple[str, ...] = (),
    integer_columns: tuple[str, ...] = (),
    decimal_columns: tuple[str, ...] = (),
    percent_columns: tuple[str, ...] = (),
    signed_columns: tuple[str, ...] = (),
) -> pd.DataFrame:
    output = dataframe.copy()
    for column in date_columns:
        if column in output.columns:
            output[column] = output[column].map(format_date_br)
    for column in integer_columns:
        if column in output.columns:
            output[column] = output[column].map(format_integer)
    for column in decimal_columns:
        if column in output.columns:
            output[column] = output[column].map(
                lambda value: format_number_br(value, decimals=1)
            )
    for column in percent_columns:
        if column in output.columns:
            output[column] = output[column].map(
                lambda value: format_percent_br(
                    value,
                    decimals=1,
                    show_sign=column in signed_columns,
                )
            )
    return output


def dataframe_to_csv_br(
    dataframe: pd.DataFrame,
    date_columns: tuple[str, ...] = (),
) -> bytes:
    output = dataframe.copy()
    for column in date_columns:
        if column in output.columns:
            output[column] = output[column].map(format_date_br)
    return output.to_csv(
        index=False,
        sep=";",
        decimal=",",
    ).encode("utf-8-sig")


def comparables_scatter_chart(
    comparables: pd.DataFrame,
    result: SimulationResult,
    scope_label: str = "Todas as praças",
) -> go.Figure:
    history = comparables.copy()
    customdata = np.column_stack(
        [
            history["operation_date"].map(format_date_br),
            history["volumes"].map(format_integer),
            history["pallets"].map(format_integer),
            history["density"].map(format_integer),
            history["loads"].map(format_integer),
        ]
    )
    load_sizes = history["loads"].astype(float)
    if load_sizes.max() > load_sizes.min():
        marker_sizes = 10 + 14 * (
            (load_sizes - load_sizes.min())
            / (load_sizes.max() - load_sizes.min())
        )
    else:
        marker_sizes = np.full(len(history), 15.0)

    figure = go.Figure()
    figure.add_trace(
        go.Scatter(
            x=history["volumes"],
            y=history["pallets"],
            mode="markers",
            name=f"Dias comparáveis · {scope_label}",
            marker={
                "size": marker_sizes,
                "color": "#1A2744",
                "opacity": 0.72,
                "line": {"color": "#FFFFFF", "width": 1},
            },
            customdata=customdata,
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Volumes: %{customdata[1]}<br>"
                "Pallets: %{customdata[2]}<br>"
                "Vol./pallet: %{customdata[3]}<br>"
                "Carregamentos: %{customdata[4]}<extra></extra>"
            ),
        )
    )
    figure.add_trace(
        go.Scatter(
            x=[result.forecast_volumes],
            y=[result.expected_pallets],
            mode="markers+text",
            name=f"Dimensionamento atual · {scope_label}",
            marker={
                "size": 17,
                "color": "#D62828",
                "line": {"color": "#FFFFFF", "width": 2},
            },
            text=[f"{format_integer(result.expected_pallets)} pallets"],
            textposition="top center",
            textfont={"color": "#D62828", "size": 13, "family": "Urbanist"},
            customdata=[
                [
                    format_date_br(result.operation_date),
                    format_integer(result.forecast_volumes),
                    format_integer(result.expected_pallets),
                ]
            ],
            hovertemplate=(
                "<b>Dimensionamento atual</b><br>"
                "Data: %{customdata[0]}<br>"
                "Volumes previstos: %{customdata[1]}<br>"
                "Pallets dimensionados: %{customdata[2]}<extra></extra>"
            ),
        )
    )
    figure.update_layout(
        title={"text": f"Dias comparáveis — {scope_label}", "x": 0},
        height=420,
        margin={"l": 20, "r": 20, "t": 45, "b": 20},
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font={"family": "Urbanist", "color": "#1A2744"},
        hovermode="closest",
        dragmode="zoom",
        separators=",.",
        legend={"orientation": "h", "y": 1.12, "x": 0},
        xaxis={
            "title": (
                "Volumes"
                if scope_label == "Todas as praças"
                else "Volumes estimados"
            ),
            "showgrid": True,
            "gridcolor": "#E8EDF5",
            "fixedrange": False,
            "tickformat": ",.0f",
        },
        yaxis={
            "title": "Pallets",
            "showgrid": True,
            "gridcolor": "#E8EDF5",
            "fixedrange": False,
            "tickformat": ",.0f",
        },
    )
    return figure


def historical_density_chart(
    daily: pd.DataFrame,
    result: SimulationResult | None = None,
    scope_label: str = "Todas as praças",
) -> go.Figure:
    history = daily.copy().sort_values("operation_date")
    history["operation_date"] = pd.to_datetime(history["operation_date"])
    customdata = np.column_stack(
        [
            history["operation_date"].dt.strftime("%d/%m/%Y"),
            history["density"].map(format_integer),
        ]
    )
    figure = go.Figure()
    figure.add_trace(
        go.Scatter(
            x=history["operation_date"],
            y=history["density"],
            mode="lines",
            name=f"Histórico · {scope_label}",
            line={"color": "#1A2744", "width": 2.5},
            customdata=customdata,
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Vol./pallet: %{customdata[1]}<extra></extra>"
            ),
        )
    )

    key_indices = {
        history["density"].idxmin(),
        history["density"].idxmax(),
        history.index[-1],
    }
    key_points = history.loc[sorted(key_indices)].copy()
    figure.add_trace(
        go.Scatter(
            x=key_points["operation_date"],
            y=key_points["density"],
            mode="markers+text",
            name="Pontos-chave",
            marker={
                "size": 10,
                "color": "#E85D04",
                "line": {"color": "#FFFFFF", "width": 1.5},
            },
            text=key_points["density"].map(format_integer),
            textposition="top center",
            textfont={"color": "#9A3E00", "size": 12, "family": "Urbanist"},
            customdata=np.column_stack(
                [
                    key_points["operation_date"].dt.strftime("%d/%m/%Y"),
                    key_points["density"].map(format_integer),
                ]
            ),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Vol./pallet: %{customdata[1]}<extra></extra>"
            ),
        )
    )

    if result is not None:
        figure.add_trace(
            go.Scatter(
                x=[pd.Timestamp(result.operation_date)],
                y=[result.expected_density],
                mode="markers+text",
                name=f"Dimensionamento atual · {scope_label}",
                marker={
                    "size": 16,
                    "color": "#D62828",
                    "line": {"color": "#FFFFFF", "width": 2},
                },
                text=[format_integer(result.expected_density)],
                textposition="top center",
                textfont={"color": "#D62828", "size": 13, "family": "Urbanist"},
                customdata=[
                    [
                        format_date_br(result.operation_date),
                        format_integer(result.expected_density),
                    ]
                ],
                hovertemplate=(
                    "<b>Dimensionamento atual</b><br>"
                    "Data: %{customdata[0]}<br>"
                    "Vol./pallet: %{customdata[1]}<extra></extra>"
                ),
            )
        )
    figure.update_layout(
        title={"text": f"Histórico diário — {scope_label}", "x": 0},
        height=400,
        margin={"l": 20, "r": 20, "t": 45, "b": 20},
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font={"family": "Urbanist", "color": "#1A2744"},
        hovermode="x unified",
        dragmode="zoom",
        separators=",.",
        legend={"orientation": "h", "y": 1.12, "x": 0},
        xaxis={
            "title": "Data",
            "tickformat": "%d/%m/%Y",
            "showgrid": True,
            "gridcolor": "#E8EDF5",
            "fixedrange": False,
            "rangeslider": {"visible": False},
        },
        yaxis={
            "title": "Volumes por pallet",
            "showgrid": True,
            "gridcolor": "#E8EDF5",
            "fixedrange": False,
            "tickformat": ",.0f",
        },
    )
    return figure


def page_forecast_dashboard(loads: pd.DataFrame, destinations: pd.DataFrame) -> None:
    page_header(
        "Painel diário de previsões",
        "Consulte o dimensionamento salvo para cada data, mesmo antes de carregar a execução.",
    )
    forecasts = load_forecasts()
    if "dashboard_date" not in st.session_state:
        st.session_state["dashboard_date"] = date.today()

    selected_date = st.date_input(
        "Data da previsão",
        format="DD/MM/YYYY",
        key="dashboard_date",
    )

    saved_notice = st.session_state.pop("forecast_saved_notice", None)
    if saved_notice:
        st.success(saved_notice)

    if forecasts.empty:
        st.info("Ainda não existem previsões salvas. Crie a primeira na Simulação diária.")
        return

    day_forecasts = forecasts[forecasts["operation_date"] == selected_date].copy()
    if day_forecasts.empty:
        st.warning(f"Não existe previsão salva para {format_date_br(selected_date)}.")
        recent = (
            forecasts.drop_duplicates(subset=["operation_date"], keep="first")
            .head(10)[
                [
                    "operation_date",
                    "forecast_volumes",
                    "expected_pallets",
                    "capacity_pallets",
                    "created_at",
                ]
            ]
            .rename(
                columns={
                    "operation_date": "Data",
                    "forecast_volumes": "Volumes previstos",
                    "expected_pallets": "Pallets esperados",
                    "capacity_pallets": "Capacidade recomendada",
                    "created_at": "Salva em",
                }
            )
        )
        recent["Salva em"] = recent["Salva em"].map(
            lambda value: value.strftime("%d/%m/%Y %H:%M") if pd.notna(value) else "—"
        )
        st.caption("Previsões mais recentes")
        st.dataframe(
            format_table_br(
                recent,
                date_columns=("Data",),
                integer_columns=(
                    "Volumes previstos",
                    "Pallets esperados",
                    "Capacidade recomendada",
                ),
            ),
            hide_index=True,
            use_container_width=True,
        )
        return

    scenario_options = {}
    for row in day_forecasts.itertuples():
        saved_at = (
            row.created_at.strftime("%d/%m/%Y %H:%M")
            if pd.notna(row.created_at)
            else "horário não disponível"
        )
        label = (
            f"#{row.forecast_id} · {format_integer(row.forecast_volumes)} volumes · "
            f"salva em {saved_at}"
        )
        scenario_options[label] = row.forecast_id

    if len(scenario_options) > 1:
        selected_label = st.selectbox(
            "Cenário salvo",
            list(scenario_options),
            key=f"dashboard_scenario_{selected_date.isoformat()}",
        )
    else:
        selected_label = next(iter(scenario_options))
        st.caption(selected_label)

    selected = day_forecasts[
        day_forecasts["forecast_id"] == scenario_options[selected_label]
    ].iloc[0]

    st.markdown(
        f"""
        <div class="dimension-highlight">
            <div>
                <span>Dimensionamento salvo</span>
                <strong>{format_date_br(selected_date)} ·
                {format_integer(selected["forecast_volumes"])} volumes →
                {format_integer(selected["expected_pallets"])} pallets</strong>
            </div>
            <div class="dimension-capacity">
                Capacidade recomendada:
                {format_integer(selected["capacity_pallets"])} pallets
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    metric1, metric2, metric3, metric4 = st.columns(4)
    metric1.metric("Volumes previstos", format_integer(selected["forecast_volumes"]))
    metric2.metric("Pallets esperados", format_integer(selected["expected_pallets"]))
    metric3.metric(
        "Capacidade recomendada",
        format_integer(selected["capacity_pallets"]),
    )
    metric4.metric(
        "Média vol./pallet",
        format_integer(selected["expected_density"]),
    )

    actual, _ = actual_for_date(loads, destinations, selected_date)
    if actual:
        st.subheader("Execução do dia")
        actual1, actual2, actual3 = st.columns(3)
        actual1.metric(
            "Volumes realizados",
            format_integer(actual["volumes"]),
            delta=format_integer(actual["volumes"] - selected["forecast_volumes"]),
        )
        actual2.metric(
            "Pallets realizados",
            format_integer(actual["pallets"]),
            delta=format_integer(actual["pallets"] - selected["expected_pallets"]),
        )
        actual3.metric(
            "Volumes/pallet real",
            format_integer(actual["density"]),
            delta=format_number_br(
                actual["density"] - selected["expected_density"],
                decimals=0,
                show_sign=True,
            ),
        )
    else:
        st.info("A execução desta data ainda não foi carregada.")

    st.subheader("Distribuição prevista por praça")
    try:
        parameters = SimulationParameters(**json.loads(selected["parameters_json"]))
        reconstructed, comparables, _ = run_simulation(
            loads,
            selected_date,
            int(selected["forecast_volumes"]),
            parameters,
        )
        distribution = destination_distribution(
            loads,
            destinations,
            comparables,
            reconstructed,
            parameters,
        )
    except (TypeError, ValueError, json.JSONDecodeError):
        distribution = pd.DataFrame()

    if distribution.empty:
        st.warning("Não foi possível reconstruir a distribuição desta previsão.")
    else:
        display = format_table_br(
            distribution,
            integer_columns=(
                "Pallets esperados",
                "Capacidade recomendada",
                "Cargas históricas",
                "Volumes estimados",
                "Média vol./pallet da praça",
            ),
            percent_columns=("Participação",),
        )
        st.dataframe(display, hide_index=True, use_container_width=True)
        st.download_button(
            "Baixar distribuição em CSV",
            dataframe_to_csv_br(distribution),
            file_name=f"previsao_salva_{selected_date.strftime('%d-%m-%Y')}.csv",
            mime="text/csv",
        )


def page_simulation(loads: pd.DataFrame, destinations: pd.DataFrame) -> None:
    page_header(
        "Dimensionamento diário de pallets",
        "Informe a data e o volume previsto. O cálculo utiliza somente execuções anteriores.",
    )

    col_date, col_volume = st.columns(2)
    operation_date = col_date.date_input(
        "Data da operação",
        value=date.today(),
        format="DD/MM/YYYY",
    )
    forecast_volumes = int(
        col_volume.number_input("Volumes previstos", min_value=0, value=0, step=500)
    )

    with st.expander("Parâmetros do modelo"):
        col1, col2, col3 = st.columns(3)
        history_weeks = col1.number_input(
            "Janela histórica (semanas)",
            4,
            52,
            16,
            help=PARAMETER_HELP["history_weeks"],
        )
        initial_band = col2.slider(
            "Faixa inicial",
            5,
            30,
            15,
            format="%d%%",
            help=PARAMETER_HELP["initial_band"],
        ) / 100
        expanded_band = col3.slider(
            "Faixa ampliada",
            10,
            50,
            25,
            format="%d%%",
            help=PARAMETER_HELP["expanded_band"],
        ) / 100
        col4, col5, col6 = st.columns(3)
        minimum_comparables = col4.number_input(
            "Mínimo de dias comparáveis",
            2,
            12,
            4,
            help=PARAMETER_HELP["minimum_comparables"],
        )
        half_life = col5.number_input(
            "Meia-vida dos comparáveis (dias)",
            7,
            90,
            21,
            help=PARAMETER_HELP["recency_half_life"],
        )
        calibration_days = col6.number_input(
            "Curva recente (dias)",
            7,
            42,
            14,
            help=PARAMETER_HELP["recent_curve"],
        )
        col7, col8 = st.columns(2)
        calibration_half_life = col7.number_input(
            "Meia-vida da curva recente (dias)",
            3,
            28,
            7,
            help=PARAMETER_HELP["calibration_half_life"],
        )
        capacity_error_percentile = col8.slider(
            "Margem pelo erro recente",
            50,
            95,
            80,
            format="P%d",
            help=PARAMETER_HELP["capacity_error_percentile"],
        ) / 100
        render_model_parameter_guide()

    parameters = SimulationParameters(
        history_weeks=int(history_weeks),
        initial_band=float(initial_band),
        expanded_band=float(expanded_band),
        minimum_comparables=int(minimum_comparables),
        recency_half_life_days=int(half_life),
        recent_calibration_days=int(calibration_days),
        calibration_half_life_days=int(calibration_half_life),
        capacity_error_percentile=float(capacity_error_percentile),
    )

    if forecast_volumes <= 0:
        st.info("Informe os volumes previstos para calcular o dimensionamento.")
        return

    try:
        result, comparables, daily = run_simulation(loads, operation_date, forecast_volumes, parameters)
    except ValueError as exc:
        st.warning(str(exc))
        return

    metric1, metric2, metric3 = st.columns(3)
    metric1.metric("Pallets esperados", format_integer(result.expected_pallets))
    metric2.metric("Capacidade recomendada", format_integer(result.capacity_pallets))
    metric3.metric("Cenário de alta densidade", format_integer(result.high_density_pallets))
    st.caption(
        f"{result.comparable_count} dias comparáveis · {result.selection_rule} · "
        f"{format_integer(result.expected_density)} volumes/pallet esperados · "
        f"margem recente de {format_integer(result.capacity_buffer)} pallets"
    )
    st.markdown(
        f"""
        <div class="dimension-highlight">
            <div>
                <span>Dimensionamento em análise</span>
                <strong>{format_date_br(result.operation_date)} ·
                {format_integer(result.forecast_volumes)} volumes →
                {format_integer(result.expected_pallets)} pallets</strong>
            </div>
            <div class="dimension-capacity">
                Capacidade recomendada: {format_integer(result.capacity_pallets)} pallets
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    distribution = destination_distribution(loads, destinations, comparables, result, parameters)
    destination_options = ["Todas as praças"]
    if not distribution.empty:
        destination_options.extend(sorted(distribution["Praça"].astype(str).unique()))
    chart_destination = st.selectbox(
        "Praça exibida nos gráficos",
        destination_options,
        key="simulation_chart_destination",
    )

    chart_comparables = comparables
    chart_daily = daily
    chart_result = result
    chart_scope = chart_destination
    if chart_destination != "Todas as praças":
        destination_daily = build_destination_daily_history(
            loads,
            destinations,
            operation_date,
            chart_destination,
        )
        comparable_dates = set(comparables["operation_date"])
        chart_comparables = destination_daily[
            destination_daily["operation_date"].isin(comparable_dates)
        ].copy()
        chart_daily = destination_daily

        destination_forecast = distribution[
            distribution["Praça"] == chart_destination
        ].iloc[0]
        destination_volumes = int(destination_forecast["Volumes estimados"])
        destination_pallets = int(destination_forecast["Pallets esperados"])
        destination_capacity = int(destination_forecast["Capacidade recomendada"])
        destination_density = float(
            destination_forecast["Média vol./pallet da praça"]
        )
        chart_result = replace(
            result,
            forecast_volumes=destination_volumes,
            expected_pallets=destination_pallets,
            capacity_pallets=destination_capacity,
            expected_density=destination_density,
            capacity_density=(
                destination_volumes / destination_capacity
                if destination_capacity
                else destination_density
            ),
        )
        st.caption(
            "Na visão por praça, os pallets são os valores executados para a unidade. "
            "Os volumes históricos são estimados proporcionalmente à participação "
            "dos pallets da praça em cada carregamento."
        )

    tab_distribution, tab_comparables, tab_history = st.tabs(
        ["Distribuição por praça", "Dias comparáveis", "Histórico diário"]
    )
    with tab_distribution:
        if distribution.empty:
            st.info("Não há distribuição por praça disponível.")
        else:
            display = format_table_br(
                distribution,
                integer_columns=(
                    "Pallets esperados",
                    "Capacidade recomendada",
                    "Cargas históricas",
                    "Volumes estimados",
                    "Média vol./pallet da praça",
                ),
                percent_columns=("Participação",),
            )
            st.dataframe(display, hide_index=True, use_container_width=True)
            st.download_button(
                "Baixar distribuição em CSV",
                dataframe_to_csv_br(distribution),
                file_name=f"distribuicao_pallets_{operation_date.strftime('%d-%m-%Y')}.csv",
                mime="text/csv",
            )
    with tab_comparables:
        if chart_comparables.empty:
            st.info(f"Não há dias comparáveis com movimentação para {chart_scope}.")
        else:
            volume_label = (
                "Volumes"
                if chart_destination == "Todas as praças"
                else "Volumes estimados"
            )
            comparable_display = chart_comparables[
                ["operation_date", "volumes", "pallets", "density", "loads"]
            ].rename(
                columns={
                    "operation_date": "Data",
                    "volumes": volume_label,
                    "pallets": "Pallets",
                    "density": "Volumes/pallet",
                    "loads": "Carregamentos",
                }
            )
            comparable_display = format_table_br(
                comparable_display,
                date_columns=("Data",),
                integer_columns=(
                    volume_label,
                    "Pallets",
                    "Volumes/pallet",
                    "Carregamentos",
                ),
            )
            st.dataframe(comparable_display, hide_index=True, use_container_width=True)
            st.plotly_chart(
                comparables_scatter_chart(
                    chart_comparables,
                    chart_result,
                    chart_scope,
                ),
                use_container_width=True,
                config=PLOTLY_CONFIG,
            )
    with tab_history:
        if chart_daily.empty:
            st.info(f"Não há histórico diário disponível para {chart_scope}.")
        else:
            st.plotly_chart(
                historical_density_chart(
                    chart_daily,
                    chart_result,
                    chart_scope,
                ),
                use_container_width=True,
                config=PLOTLY_CONFIG,
            )

    st.button(
        "Salvar previsão e abrir painel",
        type="primary",
        on_click=save_forecast_and_open_dashboard,
        args=(result, parameters),
    )


def historical_input_table(uploaded: object | None, edited: pd.DataFrame) -> pd.DataFrame:
    if uploaded is None:
        dataframe = edited.copy()
    elif uploaded.name.lower().endswith(".csv"):
        dataframe = pd.read_csv(uploaded, sep=None, engine="python")
    else:
        dataframe = pd.read_excel(uploaded, sheet_name=0)

    dataframe.columns = [normalize_header(column) for column in dataframe.columns]
    date_candidates = ["Data", "Data da operação", "Data de expedição"]
    volume_candidates = ["Volumes planejados", "Volume planejado", "Volumes", "Volume"]
    date_column = next((column for column in date_candidates if column in dataframe.columns), None)
    volume_column = next((column for column in volume_candidates if column in dataframe.columns), None)
    if date_column is None or volume_column is None:
        raise ValueError("Informe as colunas Data e Volumes planejados.")

    output = dataframe[[date_column, volume_column]].rename(
        columns={date_column: "Data", volume_column: "Volumes planejados"}
    )
    output["Data"] = pd.to_datetime(output["Data"], dayfirst=True, errors="coerce").dt.date
    output["Volumes planejados"] = pd.to_numeric(
        output["Volumes planejados"],
        errors="coerce",
    )
    output = output.dropna(subset=["Data", "Volumes planejados"])
    output = output[output["Volumes planejados"] > 0].copy()
    output["Volumes planejados"] = output["Volumes planejados"].round().astype(int)
    return output.drop_duplicates(subset=["Data"], keep="last").sort_values("Data")


def run_historical_batch(
    loads: pd.DataFrame,
    destinations: pd.DataFrame,
    planned: pd.DataFrame,
    parameters: SimulationParameters,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    daily_rows: list[dict] = []
    destination_rows: list[dict] = []

    for row in planned.itertuples(index=False):
        operation_date = row[0]
        forecast_volumes = int(row[1])
        try:
            result, comparables, _ = run_simulation(
                loads,
                operation_date,
                forecast_volumes,
                parameters,
            )
        except ValueError:
            continue

        actual, actual_destinations = actual_for_date(loads, destinations, operation_date)
        actual_pallets = actual.get("pallets", np.nan)
        pallet_error = (
            result.expected_pallets - actual_pallets
            if pd.notna(actual_pallets)
            else np.nan
        )
        daily_rows.append(
            {
                "Data": operation_date,
                "Dia da semana": WEEKDAY_NAMES[operation_date.weekday()],
                "Volumes planejados": forecast_volumes,
                "Volumes executados": actual.get("volumes", np.nan),
                "Pallets esperados": result.expected_pallets,
                "Capacidade recomendada": result.capacity_pallets,
                "Pallets executados": actual_pallets,
                "Erro em pallets": pallet_error,
                "Erro percentual": (
                    pallet_error / actual_pallets
                    if pd.notna(actual_pallets) and actual_pallets
                    else np.nan
                ),
                "Volumes/pallet esperado": result.expected_density,
                "Volumes/pallet executado": actual.get("density", np.nan),
                "Fator de ajuste recente": result.calibration_factor,
                "Margem de capacidade": result.capacity_buffer,
                "Dias comparáveis": result.comparable_count,
                "Regra de seleção": result.selection_rule,
            }
        )

        forecast_distribution = destination_distribution(
            loads,
            destinations,
            comparables,
            result,
            parameters,
        )[
            [
                "Praça",
                "Pallets esperados",
                "Capacidade recomendada",
                "Volumes estimados",
                "Média vol./pallet da praça",
                "Cargas históricas",
            ]
        ]
        actual_table = actual_destinations.rename("Pallets executados").reset_index()
        if not actual_table.empty:
            actual_table.columns = ["Praça", "Pallets executados"]
        else:
            actual_table = pd.DataFrame(columns=["Praça", "Pallets executados"])
        destination_comparison = forecast_distribution.merge(
            actual_table,
            on="Praça",
            how="outer",
        ).fillna(0)
        destination_comparison.insert(0, "Data", operation_date)
        destination_comparison["Desvio"] = (
            destination_comparison["Pallets esperados"]
            - destination_comparison["Pallets executados"]
        )
        destination_rows.extend(destination_comparison.to_dict("records"))

    return pd.DataFrame(daily_rows), pd.DataFrame(destination_rows)


def page_historical_simulation(loads: pd.DataFrame, destinations: pd.DataFrame) -> None:
    page_header(
        "Simulação histórica",
        "Simule vários dias de uma vez usando apenas execuções anteriores a cada data analisada.",
    )
    initial = pd.DataFrame(
        {
            "Data": [
                date(2026, 7, 23),
                date(2026, 7, 24),
                date(2026, 7, 25),
                date(2026, 7, 28),
            ],
            "Volumes planejados": [43_892, 41_804, 41_535, 69_907],
        }
    )
    edited = st.data_editor(
        initial,
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        column_config={
            "Data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
            "Volumes planejados": st.column_config.NumberColumn(
                "Volumes planejados",
                min_value=1,
                step=100,
                format="%d",
            ),
        },
    )
    uploaded = st.file_uploader(
        "Ou carregue uma lista em Excel/CSV",
        type=["xlsx", "xls", "csv"],
        help="Colunas esperadas: Data e Volumes planejados.",
    )

    with st.expander("Parâmetros do backtest"):
        col1, col2, col3 = st.columns(3)
        history_weeks = col1.number_input(
            "Janela histórica (semanas)",
            4,
            52,
            16,
            key="batch_history_weeks",
            help=PARAMETER_HELP["history_weeks"],
        )
        initial_band = col2.slider(
            "Faixa inicial",
            5,
            30,
            15,
            format="%d%%",
            key="batch_initial_band",
            help=PARAMETER_HELP["initial_band"],
        ) / 100
        expanded_band = col3.slider(
            "Faixa ampliada",
            10,
            50,
            25,
            format="%d%%",
            key="batch_expanded_band",
            help=PARAMETER_HELP["expanded_band"],
        ) / 100
        col4, col5, col6 = st.columns(3)
        minimum_comparables = col4.number_input(
            "Mínimo de dias comparáveis",
            2,
            12,
            4,
            key="batch_minimum_comparables",
            help=PARAMETER_HELP["minimum_comparables"],
        )
        half_life = col5.number_input(
            "Meia-vida dos comparáveis (dias)",
            7,
            90,
            21,
            key="batch_half_life",
            help=PARAMETER_HELP["recency_half_life"],
        )
        calibration_days = col6.number_input(
            "Curva recente (dias)",
            7,
            42,
            14,
            key="batch_calibration_days",
            help=PARAMETER_HELP["recent_curve"],
        )
        col7, col8 = st.columns(2)
        calibration_half_life = col7.number_input(
            "Meia-vida da curva recente (dias)",
            3,
            28,
            7,
            key="batch_calibration_half_life",
            help=PARAMETER_HELP["calibration_half_life"],
        )
        capacity_error_percentile = col8.slider(
            "Margem pelo erro recente",
            50,
            95,
            80,
            format="P%d",
            key="batch_capacity_error_percentile",
            help=PARAMETER_HELP["capacity_error_percentile"],
        ) / 100
        render_model_parameter_guide()

    parameters = SimulationParameters(
        history_weeks=int(history_weeks),
        initial_band=float(initial_band),
        expanded_band=float(expanded_band),
        minimum_comparables=int(minimum_comparables),
        recency_half_life_days=int(half_life),
        recent_calibration_days=int(calibration_days),
        calibration_half_life_days=int(calibration_half_life),
        capacity_error_percentile=float(capacity_error_percentile),
    )

    try:
        planned = historical_input_table(uploaded, edited)
    except Exception as exc:
        st.error(f"Não foi possível validar os dados planejados: {exc}")
        return
    if planned.empty:
        st.info("Inclua ao menos uma data com volume planejado.")
        return

    daily_result, destination_result = run_historical_batch(
        loads,
        destinations,
        planned,
        parameters,
    )
    if daily_result.empty:
        st.warning("Nenhuma das datas possui histórico anterior suficiente.")
        return

    comparable_actuals = daily_result.dropna(subset=["Pallets executados"]).copy()
    metric1, metric2, metric3 = st.columns(3)
    metric1.metric("Dias simulados", format_integer(len(daily_result)))
    if comparable_actuals.empty:
        metric2.metric("Erro médio absoluto", "—")
        metric3.metric("Viés médio", "—")
    else:
        metric2.metric(
            "Erro médio absoluto",
            f"{format_number_br(comparable_actuals['Erro em pallets'].abs().mean(), 1)} pallets",
        )
        metric3.metric(
            "Viés médio",
            f"{format_number_br(comparable_actuals['Erro em pallets'].mean(), 1, True)} pallets",
        )

    tab_daily, tab_destination = st.tabs(["Resultado diário", "Resultado por praça"])
    with tab_daily:
        display = format_table_br(
            daily_result,
            date_columns=("Data",),
            integer_columns=(
                "Volumes planejados",
                "Volumes executados",
                "Pallets esperados",
                "Capacidade recomendada",
                "Pallets executados",
                "Erro em pallets",
                "Volumes/pallet esperado",
                "Volumes/pallet executado",
                "Margem de capacidade",
                "Dias comparáveis",
            ),
            percent_columns=("Erro percentual", "Fator de ajuste recente"),
            signed_columns=("Erro percentual",),
        )
        st.dataframe(display, hide_index=True, use_container_width=True)
        st.download_button(
            "Baixar resultado diário em CSV",
            dataframe_to_csv_br(daily_result, date_columns=("Data",)),
            file_name="simulacao_historica_diaria.csv",
            mime="text/csv",
        )
    with tab_destination:
        selected_date = st.selectbox(
            "Data",
            sorted(destination_result["Data"].unique()),
            format_func=lambda value: value.strftime("%d/%m/%Y"),
        )
        selected_destination = destination_result[
            destination_result["Data"] == selected_date
        ].sort_values("Pallets esperados", ascending=False)
        selected_destination_display = format_table_br(
            selected_destination,
            date_columns=("Data",),
            integer_columns=(
                "Pallets esperados",
                "Capacidade recomendada",
                "Volumes estimados",
                "Média vol./pallet da praça",
                "Cargas históricas",
                "Pallets executados",
                "Desvio",
            ),
        )
        st.dataframe(
            selected_destination_display,
            hide_index=True,
            use_container_width=True,
        )
        st.download_button(
            "Baixar todas as praças em CSV",
            dataframe_to_csv_br(destination_result, date_columns=("Data",)),
            file_name="simulacao_historica_por_praca.csv",
            mime="text/csv",
        )


def page_comparison(loads: pd.DataFrame, destinations: pd.DataFrame) -> None:
    page_header(
        "Previsto × realizado",
        "Compare cenários salvos com volumes, pallets e distribuição efetivamente executados.",
    )
    forecasts = load_forecasts()
    with st.expander("Gerenciar previsões salvas"):
        if forecasts.empty:
            st.caption("Não existem cenários previstos salvos.")
        else:
            st.warning(
                f"Esta ação excluirá permanentemente os {format_integer(len(forecasts))} "
                "cenários previstos salvos."
            )
            confirm_clear = st.checkbox(
                "Confirmo que desejo excluir todos os cenários previstos.",
                key="confirm_clear_forecasts",
            )
            if st.button(
                "Limpar todos os cenários previstos",
                disabled=not confirm_clear,
                type="secondary",
            ):
                deleted = clear_forecasts()
                st.success(f"{deleted} cenários previstos foram excluídos.")
                st.rerun()
    if forecasts.empty:
        st.info("Salve uma previsão na página de simulação para iniciar as comparações.")
        return
    options = {
        f"#{row.forecast_id} · {row.operation_date:%d/%m/%Y} · {format_integer(row.forecast_volumes)} volumes": row.forecast_id
        for row in forecasts.itertuples()
    }
    selected_label = st.selectbox("Previsão", options.keys())
    selected = forecasts[forecasts["forecast_id"] == options[selected_label]].iloc[0]
    target_date = selected["operation_date"]
    actual, actual_destinations = actual_for_date(loads, destinations, target_date)
    if not actual:
        st.warning("Ainda não existe execução carregada para esta data.")
        return

    metric1, metric2, metric3 = st.columns(3)
    metric1.metric(
        "Volumes realizados",
        format_integer(actual["volumes"]),
        delta=format_integer(actual["volumes"] - selected["forecast_volumes"]),
    )
    metric2.metric(
        "Pallets realizados",
        format_integer(actual["pallets"]),
        delta=format_integer(actual["pallets"] - selected["expected_pallets"]),
    )
    metric3.metric(
        "Volumes/pallet real",
        format_integer(actual["density"]),
        delta=format_number_br(
            actual["density"] - selected["expected_density"],
            decimals=0,
            show_sign=True,
        ),
    )

    adjusted_for_volume = math.ceil(actual["volumes"] / selected["expected_density"])
    volume_effect = adjusted_for_volume - selected["expected_pallets"]
    density_effect = actual["pallets"] - adjusted_for_volume
    st.caption(
        f"Efeito do volume: {format_number_br(volume_effect, 0, True)} pallets · "
        f"Efeito de densidade/composição: "
        f"{format_number_br(density_effect, 0, True)} pallets"
    )

    summary = pd.DataFrame(
        {
            "Indicador": ["Volumes", "Pallets", "Volumes/pallet"],
            "Previsto": [
                selected["forecast_volumes"],
                selected["expected_pallets"],
                selected["expected_density"],
            ],
            "Realizado": [actual["volumes"], actual["pallets"], actual["density"]],
        }
    )
    summary["Desvio"] = summary["Realizado"] - summary["Previsto"]
    summary_display = format_table_br(
        summary,
        integer_columns=("Previsto", "Realizado", "Desvio"),
    )
    st.dataframe(summary_display, hide_index=True, use_container_width=True)

    st.subheader("Comparação por praça")
    try:
        parameters = SimulationParameters(**json.loads(selected["parameters_json"]))
        reconstructed, comparables, _ = run_simulation(
            loads,
            target_date,
            int(selected["forecast_volumes"]),
            parameters,
        )
        forecast_distribution = destination_distribution(
            loads,
            destinations,
            comparables,
            reconstructed,
            parameters,
        )[["Praça", "Pallets esperados", "Capacidade recomendada"]]
    except Exception:
        forecast_distribution = pd.DataFrame(
            columns=["Praça", "Pallets esperados", "Capacidade recomendada"]
        )

    actual_table = actual_destinations.rename("Pallets executados").reset_index()
    actual_table.columns = ["Praça", "Pallets executados"]
    comparison = forecast_distribution.merge(actual_table, on="Praça", how="outer").fillna(0)
    for column in ["Pallets esperados", "Capacidade recomendada", "Pallets executados"]:
        comparison[column] = comparison[column].astype(int)
    comparison["Desvio"] = comparison["Pallets executados"] - comparison["Pallets esperados"]
    comparison = comparison.sort_values(
        ["Pallets esperados", "Pallets executados"],
        ascending=False,
    )
    comparison_display = format_table_br(
        comparison,
        integer_columns=(
            "Pallets esperados",
            "Capacidade recomendada",
            "Pallets executados",
            "Desvio",
        ),
    )
    st.dataframe(comparison_display, hide_index=True, use_container_width=True)


def page_information() -> None:
    page_header(
        "Informações do modelo",
        "Entenda os parâmetros manipuláveis e como interpretar o dimensionamento.",
    )

    st.subheader("Parâmetros da Simulação diária")
    st.write(
        "Esses controles definem quais dias históricos entram no cálculo, quanto "
        "peso é dado às execuções recentes e qual margem de segurança será aplicada."
    )
    render_model_parameter_guide()

    st.subheader("Configuração padrão")
    defaults = pd.DataFrame(
        {
            "Parâmetro": [
                "Janela histórica",
                "Faixa inicial",
                "Faixa ampliada",
                "Mínimo de dias comparáveis",
                "Meia-vida dos comparáveis",
                "Curva recente",
                "Meia-vida da curva recente",
                "Margem pelo erro recente",
            ],
            "Valor padrão": [
                "16 semanas",
                "15%",
                "25%",
                "4 dias",
                "21 dias",
                "14 dias",
                "7 dias",
                "P80",
            ],
            "Uso recomendado": [
                "Equilíbrio entre histórico e comportamento recente",
                "Primeira busca por volumes semelhantes",
                "Busca alternativa quando a amostra é pequena",
                "Amostra mínima para formar a referência",
                "Redução gradual do peso dos registros antigos",
                "Período usado para recalibrar os desvios",
                "Maior peso para os erros da última semana",
                "Proteção operacional equilibrada",
            ],
        }
    )
    st.dataframe(defaults, hide_index=True, use_container_width=True)

    st.subheader("Como o cálculo seleciona os dias")
    st.markdown(
        """
        1. Considera somente execuções anteriores à data simulada.
        2. Procura o mesmo dia da semana dentro da **faixa inicial** de volume.
        3. Se não encontrar o **mínimo de dias comparáveis**, utiliza a
           **faixa ampliada**.
        4. Se a amostra continuar insuficiente, utiliza os registros disponíveis
           do mesmo dia da semana.
        5. Aplica maior peso aos dias recentes e corrige o resultado pela
           performance observada na **curva recente**.
        """
    )
    st.info(
        "Exemplo: para 70.000 volumes, a faixa inicial de 15% procura dias entre "
        "59.500 e 80.500 volumes. A faixa ampliada de 25% procura dias entre "
        "52.500 e 87.500 volumes."
    )

    st.subheader("Como interpretar os resultados")
    st.markdown(
        """
        | Resultado | Interpretação |
        |---|---|
        | **Pallets esperados** | Estimativa central do que deverá ser carregado, considerando histórico, dia da semana, volume e curva recente. |
        | **Capacidade recomendada** | Pallets esperados acrescidos da margem calculada pelos erros recentes. É a referência mais segura para contratação ou reserva de capacidade. |
        | **Cenário de alta densidade** | Cenário em que entram mais volumes por pallet e, consequentemente, são necessários menos pallets. |
        | **Média vol./pallet** | Relação estimada entre volumes e pallets. Na distribuição, essa média é calculada separadamente para cada praça. |
        """
    )
    st.warning(
        "Altere os parâmetros preferencialmente após analisar o Previsto × realizado. "
        "Mudar vários controles ao mesmo tempo dificulta identificar qual ajuste "
        "melhorou ou piorou a estimativa."
    )


def page_upload() -> None:
    page_header(
        "Atualização da base executada",
        "Acrescente registros novos ou atualize cargas corrigidas preservando o histórico.",
    )
    uploaded = st.file_uploader("Base executada", type=["xlsx", "xls"])
    if uploaded is None:
        st.info("Selecione uma planilha com a mesma estrutura da base limpa.")
        return
    try:
        dataframe = read_uploaded_base(uploaded)
        loads, destinations, summary = dataframe_to_records(dataframe, uploaded.name)
    except Exception as exc:
        st.error(f"Não foi possível validar o arquivo: {exc}")
        return

    metric1, metric2, metric3, metric4 = st.columns(4)
    metric1.metric("Carregamentos válidos", format_integer(summary["valid_loads"]))
    metric2.metric("Linhas ignoradas", format_integer(summary["invalid_rows"]))
    metric3.metric("Destino sem pallets", format_integer(summary["destination_without_pallet"]))
    metric4.metric("Pallets sem destino", format_integer(summary["pallet_without_destination"]))

    if not loads.empty:
        st.write(
            f"Período: **{format_date_br(loads['operation_date'].min())}** a "
            f"**{format_date_br(loads['operation_date'].max())}** · "
            f"{format_integer(loads['volumes'].sum())} volumes · "
            f"{format_integer(loads['total_pallets'].sum())} pallets"
        )
        preview = loads[["operation_date", "route", "volumes", "total_pallets"]].head(20)
        preview = preview.rename(
            columns={
                "operation_date": "Data",
                "route": "Rota",
                "volumes": "Volumes",
                "total_pallets": "Pallets",
            }
        )
        preview = format_table_br(
            preview,
            date_columns=("Data",),
            integer_columns=("Volumes", "Pallets"),
        )
        st.dataframe(preview, hide_index=True, use_container_width=True)

    if st.button("Incorporar registros válidos", type="primary", disabled=loads.empty):
        inserted, updated, unchanged = insert_records(loads, destinations)
        load_history.clear()
        st.success(
            f"{format_integer(inserted)} carregamentos novos, "
            f"{format_integer(updated)} atualizados e "
            f"{format_integer(unchanged)} registros sem alteração."
        )


def page_history(loads: pd.DataFrame) -> None:
    page_header(
        "Histórico consolidado",
        "Acompanhe a evolução diária de volumes, pallets, carregamentos e densidade.",
    )
    if loads.empty:
        st.info("Não há histórico carregado.")
        return
    daily = (
        loads.groupby("operation_date", as_index=False)
        .agg(volumes=("volumes", "sum"), pallets=("total_pallets", "sum"), carregamentos=("load_id", "nunique"))
    )
    daily["volumes_por_pallet"] = daily["volumes"] / daily["pallets"]
    col1, col2, col3 = st.columns(3)
    col1.metric(
        "Período",
        f"{format_date_br(daily['operation_date'].min())} a "
        f"{format_date_br(daily['operation_date'].max())}",
    )
    col2.metric("Volumes", format_integer(daily["volumes"].sum()))
    col3.metric("Pallets", format_integer(daily["pallets"].sum()))
    chart_daily = daily.rename(columns={"volumes_por_pallet": "density"})
    st.plotly_chart(
        historical_density_chart(chart_daily),
        use_container_width=True,
        config=PLOTLY_CONFIG,
    )
    history_display = daily.sort_values("operation_date", ascending=False).rename(
        columns={
            "operation_date": "Data",
            "volumes": "Volumes",
            "pallets": "Pallets",
            "carregamentos": "Carregamentos",
            "volumes_por_pallet": "Volumes/pallet",
        }
    )
    history_display = format_table_br(
        history_display,
        date_columns=("Data",),
        integer_columns=(
            "Volumes",
            "Pallets",
            "Carregamentos",
            "Volumes/pallet",
        ),
    )
    st.dataframe(
        history_display,
        hide_index=True,
        use_container_width=True,
    )


def main() -> None:
    st.set_page_config(
        page_title="Dimensionador de Pallets | BBM Logística",
        page_icon=str(FAVICON_PATH) if FAVICON_PATH.exists() else "📦",
        layout="wide",
    )
    apply_brand_theme()
    initialize_database()
    seed_initial_history()
    normalize_saved_destinations()
    loads, destinations = get_history()

    if LOGO_PATH.exists():
        st.sidebar.image(str(LOGO_PATH), width=205)
    st.sidebar.markdown(
        '<div class="sidebar-app-name">Dimensionador de Pallets</div>',
        unsafe_allow_html=True,
    )
    page = st.sidebar.radio(
        "Navegação",
        [
            "Painel diário",
            "Simulação diária",
            "Simulação histórica",
            "Previsto × realizado",
            "Atualizar base",
            "Histórico",
            "Informações",
        ],
        key="navigation",
    )
    st.sidebar.caption(
        f"{len(loads):,} carregamentos · "
        f"{loads['operation_date'].min():%d/%m/%Y} a {loads['operation_date'].max():%d/%m/%Y}"
        if not loads.empty
        else "Sem histórico carregado"
    )

    if page == "Painel diário":
        page_forecast_dashboard(loads, destinations)
    elif page == "Simulação diária":
        page_simulation(loads, destinations)
    elif page == "Simulação histórica":
        page_historical_simulation(loads, destinations)
    elif page == "Previsto × realizado":
        page_comparison(loads, destinations)
    elif page == "Atualizar base":
        page_upload()
    elif page == "Histórico":
        page_history(loads)
    else:
        page_information()


if __name__ == "__main__":
    main()
