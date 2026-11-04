from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

import matplotlib
import numpy as np
import pandas as pd
from flask import Flask, abort, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

matplotlib.use("Agg")
import matplotlib.pyplot as plt


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
AMOUNT_COLUMN = "IMPORTE (€)"
CACHE_VERSION = 10
BANK_PROFILE_SLUGS = frozenset({"isa", "hipoteca"})
CALCULATED_BALANCE_COLUMN = "SALDO CALCULADO"
BALANCE_CHECK_COLUMN = "COMPROBACIÓN"
ISA_MOVEMENT_COLUMNS = (
    "FECHA OPERACIÓN",
    "FECHA VALOR",
    "CONCEPTO",
    "IMPORTE EUR",
    "SALDO",
)


@dataclass(frozen=True)
class Profile:
    slug: str
    label: str
    workbook_name: str
    sheet_indices: tuple[int, ...]
    accounts: tuple[tuple[str, str, float], ...]
    start_date: str | None

    @property
    def directory(self) -> Path:
        return DATA_DIR / self.slug

    @property
    def workbook_path(self) -> Path:
        return DATA_DIR / self.workbook_name

    @property
    def balances_path(self) -> Path:
        return self.directory / f"{self.slug}_balances.csv"

    @property
    def summary_path(self) -> Path:
        return self.directory / f"{self.slug}_summary.csv"

    @property
    def total_plot_path(self) -> Path:
        return self.directory / f"{self.slug}_saldo_total.png"

    @property
    def monthly_plot_path(self) -> Path:
        return self.directory / f"{self.slug}_mensual.png"

    @property
    def monthly_state_path(self) -> Path:
        return self.directory / f"{self.slug}_monthly_plot_state.json"

    @property
    def source_csv_directory(self) -> Path:
        return self.directory

    def source_csv_path(self, sheet_name: str) -> Path:
        filename = secure_filename(sheet_name).lower().replace("_", "-")
        return self.source_csv_directory / f"{self.slug}_{filename}.csv"


PROFILES = {
    "warein": Profile(
        slug="warein",
        label="Warein",
        workbook_name="Excel_ing.xlsx",
        sheet_indices=(0, 2),
        accounts=(
            ("saldo_ing", "ING", 5063.16),
            ("saldo_naranja", "Naranja", 0.0),
            ("saldo_trade", "Trade", 0.0),
        ),
        start_date=None,
    ),
    "isa": Profile(
        slug="isa",
        label="Isa",
        workbook_name="Movs Isa.xlsx",
        sheet_indices=(0,),
        accounts=(("saldo_isa", "Isa", 1915.89),),
        start_date=None,
    ),
    "hipoteca": Profile(
        slug="hipoteca",
        label="Hipoteca",
        workbook_name="Movs Hipoteca.xlsx",
        sheet_indices=(0,),
        accounts=(("saldo_hipoteca", "Hipoteca", 0.0),),
        start_date=None,
    ),
    "total": Profile(
        slug="total",
        label="Total",
        workbook_name="",
        sheet_indices=(),
        accounts=(("saldo_total", "Total", 0.0),),
        start_date=None,
    ),
}

WAREIN_MOVEMENT_SOURCES = {
    "ing-nomina": {
        "label": "ING nómina",
        "balance_column": "saldo_ing",
        "filename": "warein_ing-nomina.csv",
    },
    "ing-naranja": {
        "label": "ING Naranja",
        "balance_column": "saldo_naranja",
        "filename": "warein_ing-naranja.csv",
    },
    "trade": {
        "label": "Trade",
        "balance_column": "saldo_trade",
        "filename": "warein_trade.csv",
    },
}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024
app.config["SECRET_KEY"] = os.environ.get("FLASK_SECRET_KEY", "local-private-dashboard")


class WorkbookError(ValueError):
    """Raised when an uploaded workbook cannot be processed."""


def get_profile(slug: str) -> Profile:
    profile = PROFILES.get(slug)
    if profile is None:
        abort(404)
    return profile


def active_workbook(profile: Profile) -> Path:
    if profile.workbook_path.exists():
        return profile.workbook_path
    raise WorkbookError(f"No se encuentra {profile.workbook_name}. Sube un archivo para continuar.")


def validate_workbook(profile: Profile, path: Path) -> None:
    try:
        workbook = pd.ExcelFile(path)
    except Exception as error:
        raise WorkbookError("El archivo no es un libro Excel .xlsx valido.") from error

    if len(workbook.sheet_names) <= max(profile.sheet_indices):
        raise WorkbookError("El libro no contiene las pestañas requeridas.")

    for sheet_index in profile.sheet_indices:
        if profile.slug in BANK_PROFILE_SLUGS:
            frame = pd.read_excel(workbook, sheet_name=sheet_index, header=None, nrows=1)
            if len(frame.columns) < len(ISA_MOVEMENT_COLUMNS):
                raise WorkbookError("La primera hoja debe tener al menos cinco columnas.")
        else:
            frame = pd.read_excel(workbook, sheet_name=sheet_index, nrows=1)
            if AMOUNT_COLUMN not in frame.columns:
                raise WorkbookError(
                    f"La hoja {sheet_index + 1} debe incluir la columna '{AMOUNT_COLUMN}'."
                )


def source_sheets(profile: Profile) -> dict[int, pd.DataFrame]:
    workbook_path = active_workbook(profile)
    validate_workbook(profile, workbook_path)
    options: dict[str, object] = {"sheet_name": list(profile.sheet_indices)}
    if profile.slug in BANK_PROFILE_SLUGS:
        options["header"] = None
    return pd.read_excel(workbook_path, **options)


def source_exports_are_stale(profile: Profile) -> bool:
    workbook = active_workbook(profile)
    try:
        sheet_count = len(pd.ExcelFile(workbook).sheet_names)
    except (OSError, ValueError) as error:
        raise WorkbookError("No se ha podido leer el libro Excel.") from error
    return any(
        not profile.source_csv_path(sheet_name).exists()
        or workbook.stat().st_mtime > profile.source_csv_path(sheet_name).stat().st_mtime
        for sheet_name in pd.ExcelFile(workbook).sheet_names
    )


def export_source_sheets(profile: Profile) -> list[dict[str, object]]:
    workbook_path = active_workbook(profile)
    try:
        workbook = pd.ExcelFile(workbook_path)
        profile.source_csv_directory.mkdir(parents=True, exist_ok=True)
        sheets: list[dict[str, object]] = []
        for sheet_index, sheet_name in enumerate(workbook.sheet_names):
            frame = pd.read_excel(workbook, sheet_name=sheet_index)
            frame.to_csv(profile.source_csv_path(sheet_name), index=False)
            sheets.append(
                {
                    "index": sheet_index,
                    "name": sheet_name,
                    "rows": len(frame),
                    "columns": len(frame.columns),
                }
            )
    except (OSError, ValueError) as error:
        raise WorkbookError("No se han podido convertir las pestañas del Excel a CSV.") from error
    return sheets


def source_sheet_metadata(profile: Profile) -> list[dict[str, object]]:
    if profile.slug == "warein":
        sheets: list[dict[str, object]] = []
        for index, (account_key, account) in enumerate(WAREIN_MOVEMENT_SOURCES.items()):
            try:
                frame = pd.read_csv(warein_movement_path(account_key), nrows=0)
            except (OSError, pd.errors.EmptyDataError) as error:
                raise WorkbookError("No se ha podido leer uno de los CSV de movimientos ING.") from error
            sheets.append(
                {
                    "index": index,
                    "name": str(account["label"]),
                    "account_key": account_key,
                    "columns": len(frame.columns),
                }
            )
        return sheets

    if profile.slug in BANK_PROFILE_SLUGS:
        frame = ensure_bank_movement_csv(profile)
        return [{"index": 0, "name": profile.label, "columns": len(frame.columns)}]

    if source_exports_are_stale(profile):
        return export_source_sheets(profile)

    try:
        workbook = pd.ExcelFile(active_workbook(profile))
    except (OSError, ValueError) as error:
        raise WorkbookError("No se ha podido leer el libro Excel.") from error
    sheets: list[dict[str, object]] = []
    for sheet_index, sheet_name in enumerate(workbook.sheet_names):
        try:
            frame = pd.read_csv(profile.source_csv_path(sheet_name), nrows=0)
        except (OSError, pd.errors.EmptyDataError) as error:
            raise WorkbookError("No se ha podido leer uno de los CSV generados.") from error
        sheets.append(
            {
                "index": sheet_index,
                "name": sheet_name,
                "columns": len(frame.columns),
            }
        )
    return sheets


def load_source_sheet(profile: Profile, sheet_index: int) -> pd.DataFrame:
    sheets = source_sheet_metadata(profile)
    if sheet_index < 0 or sheet_index >= len(sheets):
        abort(404)
    try:
        if profile.slug == "warein":
            return ensure_warein_movement_csv(str(sheets[sheet_index]["account_key"]))
        if profile.slug in BANK_PROFILE_SLUGS:
            frame = ensure_bank_movement_csv(profile).copy()
            for column in ("FECHA OPERACIÓN", "FECHA VALOR"):
                frame[column] = frame[column].dt.strftime("%d/%m/%Y")
            return frame
        return pd.read_csv(profile.source_csv_path(str(sheets[sheet_index]["name"])), keep_default_na=False)
    except (OSError, pd.errors.EmptyDataError) as error:
        raise WorkbookError("No se ha podido leer el CSV de la pestaña seleccionada.") from error


def parse_amounts(values: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(values):
        return pd.to_numeric(values, errors="coerce").fillna(0)
    numeric_values = pd.to_numeric(values, errors="coerce")
    if numeric_values.notna().eq(values.notna()).all():
        return numeric_values.fillna(0)
    normalized = values.astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(normalized, errors="coerce").fillna(0)


def warein_movement_path(account_key: str) -> Path:
    try:
        filename = WAREIN_MOVEMENT_SOURCES[account_key]["filename"]
    except KeyError as error:
        raise WorkbookError("Selecciona la cuenta ING que quieres actualizar.") from error
    return PROFILES["warein"].source_csv_directory / str(filename)


def save_warein_movements(account_key: str, movements: pd.DataFrame) -> None:
    source_path = warein_movement_path(account_key)
    source_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = source_path.with_suffix(".tmp")
    try:
        add_balance_check_columns(
            movements, AMOUNT_COLUMN, "SALDO (€)"
        ).to_csv(temporary_path, index=False, date_format="%Y-%m-%d")
        temporary_path.replace(source_path)
    except OSError as error:
        temporary_path.unlink(missing_ok=True)
        raise WorkbookError("No se ha podido guardar el CSV de movimientos ING.") from error


def ensure_warein_movement_csv(account_key: str) -> pd.DataFrame:
    source_path = warein_movement_path(account_key)
    try:
        existing = pd.read_csv(source_path)
    except (OSError, pd.errors.EmptyDataError) as error:
        raise WorkbookError("No se ha podido leer el CSV de movimientos ING.") from error
    normalized = normalize_movement_frame(existing)
    if not has_balance_check_columns(normalized):
        save_warein_movements(account_key, normalized)
    return add_balance_check_columns(normalized, AMOUNT_COLUMN, "SALDO (€)")


def normalize_movement_frame(frame: pd.DataFrame, require_valid_dates: bool = False) -> pd.DataFrame:
    frame = frame.dropna(axis=1, how="all").copy()
    frame.columns = [str(column).strip() for column in frame.columns]
    frame = frame.loc[:, ~frame.columns.str.startswith("Unnamed:")]
    if "F. VALOR" not in frame.columns or AMOUNT_COLUMN not in frame.columns:
        raise WorkbookError("El archivo no contiene las columnas 'F. VALOR' e 'IMPORTE (€)'.")
    frame = frame.dropna(how="all").copy()
    frame["F. VALOR"] = pd.to_datetime(frame["F. VALOR"], errors="coerce").dt.normalize()
    if require_valid_dates and frame["F. VALOR"].isna().any():
        raise WorkbookError("El archivo contiene movimientos con una fecha no válida.")
    frame[AMOUNT_COLUMN] = parse_amounts(frame[AMOUNT_COLUMN])
    return frame


def add_balance_check_columns(
    frame: pd.DataFrame, amount_column: str, balance_column: str
) -> pd.DataFrame:
    if frame.empty:
        frame[CALCULATED_BALANCE_COLUMN] = pd.Series(dtype=float)
        frame[BALANCE_CHECK_COLUMN] = pd.Series(dtype=str)
        return frame

    balances = parse_amounts(frame[balance_column])
    amounts = parse_amounts(frame[amount_column])
    calculated_balances = np.empty(len(frame))
    calculated_balances[-1] = balances.iloc[-1]
    if len(frame) > 1:
        calculated_balances[:-1] = (
            balances.iloc[-1]
            + amounts.iloc[:-1].iloc[::-1].to_numpy().cumsum()[::-1]
        )

    frame = frame.copy()
    frame[CALCULATED_BALANCE_COLUMN] = np.round(calculated_balances, 2)
    frame[BALANCE_CHECK_COLUMN] = np.where(
        np.isclose(balances.to_numpy(), calculated_balances, rtol=0, atol=0.01),
        "OK",
        "ERROR",
    )
    return frame


def has_balance_check_columns(frame: pd.DataFrame) -> bool:
    return {CALCULATED_BALANCE_COLUMN, BALANCE_CHECK_COLUMN}.issubset(frame.columns)


def read_bank_movements(path: Path) -> pd.DataFrame:
    try:
        raw = pd.read_excel(path, header=None)
    except (OSError, ValueError, ImportError) as error:
        raise WorkbookError("No se ha podido leer el Excel bancario.") from error

    header_rows = raw.index[
        raw.apply(
            lambda row: "F. VALOR" in row.astype(str).str.strip().tolist()
            and AMOUNT_COLUMN in row.astype(str).str.strip().tolist(),
            axis=1,
        )
    ]
    if header_rows.empty:
        raise WorkbookError("No se ha encontrado la fila de encabezados de movimientos ING.")

    header_row = int(header_rows[0])
    frame = raw.iloc[header_row + 1 :].copy()
    frame.columns = raw.iloc[header_row].tolist()
    return normalize_movement_frame(frame, require_valid_dates=True)


def movement_fingerprints(frame: pd.DataFrame, columns: list[str]) -> pd.Series:
    values = frame.loc[:, columns].copy()
    for column in columns:
        if column in {"F. VALOR", "FECHA OPERACIÓN", "FECHA VALOR"}:
            values[column] = pd.to_datetime(values[column], errors="coerce").dt.strftime("%Y-%m-%d")
    return values.fillna("").astype(str).agg("\x1f".join, axis=1)


def replace_warein_movements(
    existing: pd.DataFrame, uploaded: pd.DataFrame
) -> tuple[pd.DataFrame, int]:
    fingerprint_columns = ["F. VALOR", "DESCRIPCIÓN", AMOUNT_COLUMN, "SALDO (€)"]
    missing_columns = set(fingerprint_columns).difference(existing.columns).union(
        set(fingerprint_columns).difference(uploaded.columns)
    )
    if missing_columns:
        raise WorkbookError(
            "El archivo no contiene las columnas necesarias para validar los movimientos."
        )

    existing_fingerprints = set(movement_fingerprints(existing, fingerprint_columns))
    uploaded_fingerprints = movement_fingerprints(uploaded, fingerprint_columns)
    if not uploaded_fingerprints.isin(existing_fingerprints).any():
        raise WorkbookError(
            "Excel de importación incorrecto: no coincide ningún movimiento existente "
            "(fecha, descripción, importe y saldo)."
        )

    return uploaded.copy(), len(uploaded)


def bank_movement_path(profile: Profile) -> Path:
    return profile.directory / f"{profile.slug}_movements.csv"


def normalize_bank_movements(frame: pd.DataFrame, require_valid_dates: bool = False) -> pd.DataFrame:
    frame = frame.dropna(axis=1, how="all").copy()
    frame.columns = [str(column).strip() for column in frame.columns]
    missing_columns = set(ISA_MOVEMENT_COLUMNS).difference(frame.columns)
    if missing_columns:
        raise WorkbookError("El archivo no contiene las columnas de movimientos esperadas.")

    audit_columns = [
        column
        for column in (CALCULATED_BALANCE_COLUMN, BALANCE_CHECK_COLUMN)
        if column in frame.columns
    ]
    frame = frame.loc[:, [*ISA_MOVEMENT_COLUMNS, *audit_columns]].dropna(how="all").copy()
    for column in ("FECHA OPERACIÓN", "FECHA VALOR"):
        values = frame[column]
        iso_dates = values.astype(str).str.fullmatch(r"\d{4}-\d{2}-\d{2}")
        frame.loc[iso_dates, column] = pd.to_datetime(
            values.loc[iso_dates], errors="coerce", format="%Y-%m-%d"
        )
        frame.loc[~iso_dates, column] = pd.to_datetime(
            values.loc[~iso_dates], errors="coerce", dayfirst=True
        )
        frame[column] = pd.to_datetime(frame[column], errors="coerce").dt.normalize()
    if require_valid_dates and frame["FECHA OPERACIÓN"].isna().any():
        raise WorkbookError("El archivo contiene movimientos con una fecha de operación no válida.")
    frame["CONCEPTO"] = frame["CONCEPTO"].fillna("").astype(str).str.strip()
    frame["IMPORTE EUR"] = parse_amounts(frame["IMPORTE EUR"])
    frame["SALDO"] = parse_amounts(frame["SALDO"])
    return frame


def read_bank_account_movements(path: Path) -> pd.DataFrame:
    try:
        raw = pd.read_excel(path, header=None)
    except (OSError, ValueError, ImportError) as error:
        raise WorkbookError("No se ha podido leer el Excel bancario.") from error

    expected_columns = set(ISA_MOVEMENT_COLUMNS)
    header_rows = raw.index[
        raw.apply(
            lambda row: expected_columns.issubset(
                {str(value).strip() for value in row.dropna().tolist()}
            ),
            axis=1,
        )
    ]
    if header_rows.empty:
        if raw.shape[1] < len(ISA_MOVEMENT_COLUMNS):
            raise WorkbookError("El Excel no contiene las columnas de movimientos esperadas.")
        frame = raw.iloc[:, : len(ISA_MOVEMENT_COLUMNS)].copy()
        frame.columns = ISA_MOVEMENT_COLUMNS
    else:
        header_row = int(header_rows[0])
        frame = raw.iloc[header_row + 1 :].copy()
        frame.columns = raw.iloc[header_row].tolist()
    return normalize_bank_movements(frame, require_valid_dates=True)


def read_legacy_bank_movements(profile: Profile) -> pd.DataFrame:
    try:
        raw = pd.read_excel(active_workbook(profile), sheet_name=0, header=None)
    except (OSError, ValueError, ImportError) as error:
        raise WorkbookError(f"No se ha podido leer el historial de movimientos de {profile.label}.") from error
    if raw.shape[1] < len(ISA_MOVEMENT_COLUMNS):
        raise WorkbookError("El historial de Isa no tiene las columnas de movimientos requeridas.")
    frame = raw.iloc[:, : len(ISA_MOVEMENT_COLUMNS)].copy()
    frame.columns = ISA_MOVEMENT_COLUMNS
    if set(ISA_MOVEMENT_COLUMNS).issubset(set(frame.iloc[0].astype(str).str.strip())):
        frame = frame.iloc[1:].copy()
    return normalize_bank_movements(frame, require_valid_dates=True)


def save_bank_movements(profile: Profile, movements: pd.DataFrame) -> None:
    source_path = bank_movement_path(profile)
    source_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = source_path.with_suffix(".tmp")
    try:
        add_balance_check_columns(movements, "IMPORTE EUR", "SALDO").to_csv(
            temporary_path, index=False, date_format="%Y-%m-%d"
        )
        temporary_path.replace(source_path)
    except OSError as error:
        temporary_path.unlink(missing_ok=True)
        raise WorkbookError(f"No se ha podido guardar el CSV de movimientos de {profile.label}.") from error


def ensure_bank_movement_csv(profile: Profile) -> pd.DataFrame:
    source_path = bank_movement_path(profile)
    try:
        existing = pd.read_csv(source_path)
    except FileNotFoundError:
        existing = read_legacy_bank_movements(profile)
    except (OSError, pd.errors.EmptyDataError) as error:
        raise WorkbookError(f"No se ha podido leer el CSV de movimientos de {profile.label}.") from error
    normalized = normalize_bank_movements(existing, require_valid_dates=True)
    if not source_path.exists() or not has_balance_check_columns(normalized):
        save_bank_movements(profile, normalized)
    return add_balance_check_columns(normalized, "IMPORTE EUR", "SALDO")


def replace_bank_movements(
    existing: pd.DataFrame, uploaded: pd.DataFrame
) -> tuple[pd.DataFrame, int]:
    fingerprint_columns = [
        "FECHA OPERACIÓN",
        "CONCEPTO",
        "IMPORTE EUR",
        "SALDO",
    ]
    existing_fingerprints = set(movement_fingerprints(existing, fingerprint_columns))
    uploaded_fingerprints = movement_fingerprints(uploaded, fingerprint_columns)
    if not uploaded_fingerprints.isin(existing_fingerprints).any():
        raise WorkbookError(
            "Excel de importación incorrecto: no coincide ningún movimiento existente "
            "(fecha, descripción, importe y saldo)."
        )

    return uploaded.copy(), len(uploaded)


def movement_check_status(profile: Profile) -> tuple[str, bool]:
    if profile.slug == "warein":
        sources = [
            (
                ensure_warein_movement_csv(account_key),
                "F. VALOR",
                AMOUNT_COLUMN,
                "SALDO (€)",
            )
            for account_key in WAREIN_MOVEMENT_SOURCES
        ]
    elif profile.slug in BANK_PROFILE_SLUGS:
        sources = [
            (
                ensure_bank_movement_csv(profile),
                "FECHA OPERACIÓN",
                "IMPORTE EUR",
                "SALDO",
            )
        ]
    else:
        return "", False

    latest_date: pd.Timestamp | None = None
    error_count = 0
    for frame, date_column, amount_column, balance_column in sources:
        audited = add_balance_check_columns(frame, amount_column, balance_column)
        source_latest_date = pd.to_datetime(audited[date_column]).max()
        latest_date = (
            source_latest_date
            if latest_date is None
            else max(latest_date, source_latest_date)
        )
        error_count += int(audited[BALANCE_CHECK_COLUMN].eq("ERROR").sum())

    if latest_date is None:
        raise WorkbookError("No hay movimientos para comprobar.")
    message = (
        f"Saldo calculado hasta el {latest_date.strftime('%d %b %Y')} "
        "a partir de los CSV de movimientos."
    )
    if error_count:
        return f"{message} {error_count} errores en movimientos.", True
    return f"{message} Sin errores en movimientos.", False


def refresh_bank_artifacts(profile: Profile) -> None:
    for generated_file in (
        profile.balances_path,
        profile.summary_path,
        profile.total_plot_path,
        profile.monthly_plot_path,
        profile.monthly_state_path,
    ):
        generated_file.unlink(missing_ok=True)
    balances = calculate_balances(profile)
    create_total_plot(profile, balances)
    create_monthly_plot(profile, balances)


def parse_manual_amount(value: str) -> float:
    if not value.strip():
        raise WorkbookError("Indica el importe del movimiento.")

    amount = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.isna(amount):
        normalized = value.replace(".", "").replace(",", ".")
        amount = pd.to_numeric(pd.Series([normalized]), errors="coerce").iloc[0]
    if pd.isna(amount) or not np.isfinite(float(amount)):
        raise WorkbookError("El importe debe ser un número válido.")
    return float(amount)


def append_manual_trade_movement(
    movement_date: date, category: str, subcategory: str, description: str, amount: float
) -> None:
    target_path = warein_movement_path("trade")
    try:
        existing = normalize_movement_frame(pd.read_csv(target_path), require_valid_dates=True)
    except (OSError, pd.errors.EmptyDataError) as error:
        raise WorkbookError("No se ha podido leer el CSV de movimientos Trade.") from error

    new_row = {column: "" for column in existing.columns}
    new_row.update(
        {
            "F. VALOR": pd.Timestamp(movement_date),
            "CATEGORÍA": category,
            "SUBCATEGORÍA": subcategory,
            "DESCRIPCIÓN": description,
            AMOUNT_COLUMN: amount,
            "SALDO (€)": "",
        }
    )
    updated = pd.concat([pd.DataFrame([new_row]), existing], ignore_index=True)

    save_warein_movements("trade", updated)


def refresh_warein_artifacts(profile: Profile) -> None:
    for generated_file in (
        profile.balances_path,
        profile.summary_path,
        profile.total_plot_path,
        profile.monthly_plot_path,
        profile.monthly_state_path,
    ):
        generated_file.unlink(missing_ok=True)
    balances = calculate_balances(profile)
    create_total_plot(profile, balances)
    create_monthly_plot(profile, balances)


def calculate_total_balances(profile: Profile) -> pd.DataFrame:
    source_profiles = [PROFILES[slug] for slug in ("warein", "isa", "hipoteca")]
    source_balances = [cached_balances(source) for source in source_profiles]
    first_date = min(frame["fecha"].min() for frame in source_balances)
    last_date = max(frame["fecha"].max() for frame in source_balances)
    days = pd.date_range(first_date, last_date, freq="D")
    balances = pd.DataFrame({"fecha": days})
    balances["saldo_total"] = np.zeros(len(days))
    for frame in source_balances:
        daily_balance = frame.set_index("fecha")["saldo_total"]
        balances["saldo_total"] += daily_balance.reindex(days).ffill().fillna(0).to_numpy()
    balances["saldo_total"] = balances["saldo_total"].round(2)
    balances.to_csv(profile.balances_path, index=False, date_format="%Y-%m-%d")
    latest = balances.iloc[-1]
    pd.DataFrame(
        [
            {
                "cache_version": CACHE_VERSION,
                "fecha": latest["fecha"].strftime("%Y-%m-%d"),
                "saldo_total": latest["saldo_total"],
            }
        ]
    ).to_csv(profile.summary_path, index=False)
    return balances


def calculate_balances(profile: Profile) -> pd.DataFrame:
    profile.directory.mkdir(parents=True, exist_ok=True)
    if profile.slug == "total":
        return calculate_total_balances(profile)
    prepared_sheets: list[pd.DataFrame] = []
    first_date: pd.Timestamp | None = None
    last_date: pd.Timestamp | None = None

    if profile.slug == "warein":
        source_frames = [
            ensure_warein_movement_csv(account_key)
            for account_key in WAREIN_MOVEMENT_SOURCES
        ]
    else:
        source_frames = [ensure_bank_movement_csv(profile)]

    for frame in source_frames:
        if profile.slug in BANK_PROFILE_SLUGS:
            frame["fecha"] = frame["FECHA OPERACIÓN"]
            frame[AMOUNT_COLUMN] = frame["IMPORTE EUR"]
            frame = add_balance_check_columns(frame, AMOUNT_COLUMN, "SALDO")
        else:
            frame = normalize_movement_frame(frame)
            frame["fecha"] = frame["F. VALOR"]
            frame[AMOUNT_COLUMN] = parse_amounts(frame[AMOUNT_COLUMN])
            frame = add_balance_check_columns(frame, AMOUNT_COLUMN, "SALDO (€)")
        frame = frame.dropna(subset=["fecha"])
        prepared_sheets.append(frame)
        if not frame.empty:
            sheet_start = frame["fecha"].min()
            first_date = sheet_start if first_date is None else min(first_date, sheet_start)
            sheet_end = frame["fecha"].max()
            last_date = sheet_end if last_date is None else max(last_date, sheet_end)

    if first_date is None or last_date is None:
        raise WorkbookError("El Excel no contiene movimientos con una fecha valida.")

    start = pd.Timestamp(profile.start_date) if profile.start_date else first_date
    end = max(pd.Timestamp.today().normalize(), last_date)
    days = pd.date_range(start, end, freq="D")
    balances = pd.DataFrame({"fecha": days})
    account_columns: list[str] = []

    for account, frame in zip(profile.accounts, prepared_sheets):
        column_name, _label, initial_balance = account
        closing_balances = (
            frame.drop_duplicates(subset=["fecha"], keep="first")
            .set_index("fecha")[CALCULATED_BALANCE_COLUMN]
        )
        balances[column_name] = closing_balances.reindex(days).ffill().bfill().to_numpy()
        account_columns.append(column_name)

    balances["saldo_total"] = balances[account_columns].sum(axis=1).round(2)
    balances.to_csv(profile.balances_path, index=False, date_format="%Y-%m-%d")
    latest = balances.iloc[-1]
    summary = {
        "cache_version": CACHE_VERSION,
        "fecha": latest["fecha"].strftime("%Y-%m-%d"),
        "saldo_total": latest["saldo_total"],
    }
    summary.update({column: latest[column] for column in account_columns})
    pd.DataFrame([summary]).to_csv(profile.summary_path, index=False)
    return balances


def workbook_is_newer(profile: Profile, generated_file: Path) -> bool:
    workbook = active_workbook(profile)
    return not generated_file.exists() or workbook.stat().st_mtime > generated_file.stat().st_mtime


def cache_is_stale(profile: Profile) -> bool:
    if profile.slug == "total":
        if not profile.balances_path.exists() or not profile.summary_path.exists():
            return True
        return any(
            cache_is_stale(source)
            or source.balances_path.stat().st_mtime > profile.balances_path.stat().st_mtime
            for source in (PROFILES["warein"], PROFILES["isa"], PROFILES["hipoteca"])
        )
    if profile.slug in BANK_PROFILE_SLUGS:
        source_path = bank_movement_path(profile)
        if not source_path.exists():
            return True
        if (
            not profile.balances_path.exists()
            or not profile.summary_path.exists()
            or source_path.stat().st_mtime > profile.balances_path.stat().st_mtime
            or source_path.stat().st_mtime > profile.summary_path.stat().st_mtime
        ):
            return True
    elif profile.slug == "warein":
        source_paths = [warein_movement_path(account_key) for account_key in WAREIN_MOVEMENT_SOURCES]
        if (
            not profile.balances_path.exists()
            or not profile.summary_path.exists()
            or any(
                not source_path.exists()
                or source_path.stat().st_mtime > profile.balances_path.stat().st_mtime
                or source_path.stat().st_mtime > profile.summary_path.stat().st_mtime
                for source_path in source_paths
            )
        ):
            return True
    elif workbook_is_newer(profile, profile.balances_path) or workbook_is_newer(
        profile, profile.summary_path
    ):
        return True
    try:
        cache_version = pd.read_csv(profile.summary_path, nrows=1).iloc[0].get("cache_version")
        return int(cache_version) != CACHE_VERSION
    except (OSError, ValueError, TypeError, pd.errors.EmptyDataError):
        return True


def refresh_cached_data(profile: Profile) -> pd.DataFrame:
    balances = calculate_balances(profile)
    create_total_plot(profile, balances)
    create_monthly_plot(profile, balances)
    return balances


def cached_balances(profile: Profile) -> pd.DataFrame:
    if cache_is_stale(profile):
        return refresh_cached_data(profile)
    return pd.read_csv(profile.balances_path, parse_dates=["fecha"])


def cached_summary(profile: Profile) -> dict[str, object]:
    if cache_is_stale(profile):
        refresh_cached_data(profile)
    summary = pd.read_csv(profile.summary_path, parse_dates=["fecha"]).iloc[0]
    accounts = [
        {"label": label, "value": float(summary[column])}
        for column, label, _initial_balance in profile.accounts
    ]
    return {
        "date": summary["fecha"].strftime("%d %b %Y"),
        "total": float(summary["saldo_total"]),
        "accounts": accounts,
    }


def create_total_plot(profile: Profile, balances: pd.DataFrame) -> Path:
    rolling_average = balances["saldo_total"].rolling(window=50, min_periods=1).mean()
    fig, axis = plt.subplots(figsize=(15, 9), dpi=100)
    axis.plot(balances["fecha"], balances["saldo_total"], alpha=0.2, linewidth=1)
    axis.plot(balances["fecha"], rolling_average, linewidth=2, color=profile_color(profile))
    axis.set_title(
        f"Saldo {profile.label}: {balances['saldo_total'].iloc[-1]:,.2f} EUR - "
        f"Actualizado: {datetime.now():%d/%m/%y %H:%M}",
        fontsize=14,
    )
    axis.set_xlabel("Fecha")
    axis.set_ylabel("Saldo (EUR)")
    axis.set_ylim(bottom=0)
    axis.set_xlim(balances["fecha"].iloc[0], balances["fecha"].iloc[-1])
    axis.grid(True, linestyle="--", linewidth=0.9, alpha=0.3)
    fig.savefig(profile.total_plot_path, bbox_inches="tight", dpi=100)
    plt.close(fig)
    return profile.total_plot_path


def profile_color(profile: Profile) -> str:
    return {
        "warein": "#4c7a9e",
        "isa": "#b54a3a",
        "hipoteca": "#4d805b",
        "total": "#5d6f96",
    }[profile.slug]


def latest_completed_month() -> date:
    return date.today().replace(day=1) - timedelta(days=1)


def monthly_plot_is_current(profile: Profile) -> bool:
    if not profile.monthly_plot_path.exists() or not profile.monthly_state_path.exists():
        return False
    try:
        state = json.loads(profile.monthly_state_path.read_text())
        return state.get("through_month") == latest_completed_month().isoformat()
    except (OSError, json.JSONDecodeError):
        return False


def create_monthly_plot(profile: Profile, balances: pd.DataFrame) -> Path:
    completed_month = latest_completed_month()
    monthly = balances.loc[balances["fecha"] <= pd.Timestamp(completed_month)].copy()
    monthly = monthly.loc[monthly["fecha"].dt.is_month_end].copy()
    if len(monthly) < 2:
        raise WorkbookError("Aun no hay dos meses completos para generar el grafico mensual.")

    monthly["ganancia"] = monthly["saldo_total"].diff()
    monthly = monthly.iloc[1:].copy()
    labels = monthly["fecha"].dt.strftime("%b %Y").str.capitalize()
    positions = np.arange(len(monthly))
    colors = np.where(monthly["ganancia"] >= 0, "#237a57", "#bb3e3e")
    tick_step = max(1, int(np.ceil(len(positions) / 10)))
    fig, axis = plt.subplots(figsize=(12, 6), dpi=130)
    axis.bar(positions, monthly["ganancia"], color=colors, edgecolor="#283334", width=0.8)
    axis.set_xticks(positions[::tick_step], labels.iloc[::tick_step], rotation=45, ha="right")
    axis.set_title(f"Ganancia mensual {profile.label} - Actualizado: {datetime.now():%d/%m/%y %H:%M}")
    axis.set_xlabel("Mes completado")
    axis.set_ylabel("Ganancia (EUR)")
    axis.set_ylim(-5000, 5000)
    axis.grid(axis="y", linestyle="--", alpha=0.45)
    fig.tight_layout()
    fig.savefig(profile.monthly_plot_path, dpi=200)
    plt.close(fig)
    profile.monthly_state_path.write_text(json.dumps({"through_month": completed_month.isoformat()}))
    return profile.monthly_plot_path


def chart_version(path: Path) -> int:
    return int(path.stat().st_mtime) if path.exists() else 0


@app.route("/")
def index():
    return redirect(url_for("dashboard", profile_slug="warein"))


@app.route("/<profile_slug>")
def dashboard(profile_slug: str):
    profile = get_profile(profile_slug)
    try:
        source = (
            "CSV de movimientos de ING"
            if profile.slug == "warein"
            else f"CSV de movimientos de {profile.label}"
            if profile.slug in BANK_PROFILE_SLUGS
            else None
        )
        return render_template(
            "dashboard.html",
            profile=profile,
            summary=cached_summary(profile),
            source=source,
            movement_check=movement_check_status(profile),
        )
    except WorkbookError as error:
        return render_template("dashboard.html", profile=profile, summary=None, workbook=None, error=str(error)), 400


@app.route("/<profile_slug>/upload", methods=["POST"])
def upload_workbook(profile_slug: str):
    profile = get_profile(profile_slug)
    if profile.slug == "total":
        abort(404)
    upload = request.files.get("workbook")
    if not upload or not upload.filename:
        flash("Selecciona un archivo Excel.", "error")
        return redirect(url_for("dashboard", profile_slug=profile.slug))
    suffix = Path(secure_filename(upload.filename)).suffix.lower()
    allowed_extensions = {".xls", ".xlsx"}
    if suffix not in allowed_extensions:
        flash("Solo se admiten archivos Excel válidos.", "error")
        return redirect(url_for("dashboard", profile_slug=profile.slug))

    profile.directory.mkdir(parents=True, exist_ok=True)
    temporary_path = profile.directory / f"upload-check{suffix}"
    upload.save(temporary_path)
    try:
        if profile.slug == "warein":
            account_key = request.form.get("account")
            uploaded_movements = read_bank_movements(temporary_path)
            existing_movements = ensure_warein_movement_csv(account_key or "")

            updated_movements, movement_count = replace_warein_movements(
                existing_movements, uploaded_movements
            )
            save_warein_movements(account_key or "", updated_movements)
            temporary_path.unlink(missing_ok=True)
            refresh_warein_artifacts(profile)
            account_label = WAREIN_MOVEMENT_SOURCES[account_key]["label"]
            flash(
                f"{account_label} actualizado: {movement_count} movimientos importados "
                "en el orden original del Excel. Los saldos diarios, el resumen y los "
                "gráficos se han regenerado.",
                "success",
            )
            return redirect(url_for("dashboard", profile_slug=profile.slug))

        if profile.slug in BANK_PROFILE_SLUGS:
            uploaded_movements = read_bank_account_movements(temporary_path)
            existing_movements = ensure_bank_movement_csv(profile)
            updated_movements, movement_count = replace_bank_movements(
                existing_movements, uploaded_movements
            )
            save_bank_movements(profile, updated_movements)
            temporary_path.unlink(missing_ok=True)
            refresh_bank_artifacts(profile)
            flash(
                f"{profile.label} actualizado: {movement_count} movimientos importados "
                "en el orden original del Excel. Los saldos, el resumen y los gráficos "
                "se han regenerado.",
                "success",
            )
            return redirect(url_for("dashboard", profile_slug=profile.slug))

        validate_workbook(profile, temporary_path)
        temporary_path.replace(profile.workbook_path)
        export_source_sheets(profile)
        for generated_file in (
            profile.balances_path,
            profile.summary_path,
            profile.total_plot_path,
        ):
            generated_file.unlink(missing_ok=True)
        balances = calculate_balances(profile)
        create_total_plot(profile, balances)
        if not monthly_plot_is_current(profile):
            create_monthly_plot(profile, balances)
        flash(f"{profile.workbook_name} actualizado. El resumen y su grafico se han regenerado.", "success")
    except WorkbookError as error:
        temporary_path.unlink(missing_ok=True)
        flash(str(error), "error")
    return redirect(url_for("dashboard", profile_slug=profile.slug))


@app.route("/<profile_slug>/trade/movimientos", methods=["POST"])
def add_trade_movement(profile_slug: str):
    profile = get_profile(profile_slug)
    if profile.slug != "warein":
        abort(404)

    try:
        date_value = request.form.get("movement_date", "").strip()
        try:
            movement_date = datetime.strptime(date_value, "%Y-%m-%d").date()
        except ValueError as error:
            raise WorkbookError("Indica una fecha válida para el movimiento.") from error

        category = request.form.get("category", "").strip()
        subcategory = request.form.get("subcategory", "").strip()
        description = request.form.get("description", "").strip()
        if not description:
            raise WorkbookError("Indica una descripción para el movimiento.")
        if len(category) > 120 or len(subcategory) > 120 or len(description) > 500:
            raise WorkbookError("El texto del movimiento supera la longitud permitida.")

        amount = parse_manual_amount(request.form.get("amount", ""))
        append_manual_trade_movement(movement_date, category, subcategory, description, amount)
        refresh_warein_artifacts(profile)
        flash(
            "Movimiento añadido a Trade. Los saldos, el resumen y los gráficos se han regenerado.",
            "success",
        )
    except WorkbookError as error:
        flash(str(error), "error")
    return redirect(url_for("total_page", profile_slug=profile.slug, sheet=2))


@app.route("/<profile_slug>/saldos")
def balances_page(profile_slug: str):
    profile = get_profile(profile_slug)
    try:
        balances = cached_balances(profile).iloc[::-1].copy()
        balances["fecha"] = balances["fecha"].dt.strftime("%d/%m/%Y")
        columns = [{"key": key, "label": label} for key, label, _initial in profile.accounts]
        return render_template("balances.html", profile=profile, rows=balances.head(180).to_dict("records"), columns=columns)
    except WorkbookError as error:
        flash(str(error), "error")
        return redirect(url_for("dashboard", profile_slug=profile.slug))


@app.route("/<profile_slug>/total")
def total_page(profile_slug: str):
    profile = get_profile(profile_slug)
    if profile.slug == "total":
        abort(404)
    try:
        sheets = source_sheet_metadata(profile)
        selected_sheet = request.args.get("sheet", default=0, type=int)
        frame = load_source_sheet(profile, selected_sheet)
        selected = sheets[selected_sheet]
        return render_template(
            "total.html",
            profile=profile,
            sheets=sheets,
            selected_sheet=selected_sheet,
            selected=selected,
            columns=frame.columns.tolist(),
            rows=frame.to_dict("records"),
            current_date=date.today().isoformat(),
        )
    except WorkbookError as error:
        flash(str(error), "error")
        return redirect(url_for("dashboard", profile_slug=profile.slug))


@app.route("/<profile_slug>/total/<int:sheet_index>.csv")
def source_sheet_csv(profile_slug: str, sheet_index: int):
    profile = get_profile(profile_slug)
    sheets = source_sheet_metadata(profile)
    if sheet_index < 0 or sheet_index >= len(sheets):
        abort(404)
    sheet_name = str(sheets[sheet_index]["name"])
    load_source_sheet(profile, sheet_index)
    source_path = (
        warein_movement_path(str(sheets[sheet_index]["account_key"]))
        if profile.slug == "warein"
        else bank_movement_path(profile)
    )
    return send_file(
        source_path,
        mimetype="text/csv",
        as_attachment=True,
        download_name=source_path.name,
        max_age=0,
    )


@app.route("/<profile_slug>/plot-total.png")
def total_plot(profile_slug: str):
    profile = get_profile(profile_slug)
    if not profile.total_plot_path.exists():
        return "El grafico se genera al actualizar el archivo Excel.", 404
    return send_file(profile.total_plot_path, mimetype="image/png", max_age=0)


@app.route("/<profile_slug>/mensual")
def monthly_page(profile_slug: str):
    profile = get_profile(profile_slug)
    return render_template(
        "monthly.html",
        profile=profile,
        has_plot=profile.monthly_plot_path.exists(),
        is_current=monthly_plot_is_current(profile),
        version=chart_version(profile.monthly_plot_path),
    )


@app.route("/<profile_slug>/generar-mensual", methods=["POST"])
def generate_monthly_plot(profile_slug: str):
    profile = get_profile(profile_slug)
    if monthly_plot_is_current(profile):
        flash("No hay ningun mes completo nuevo desde el ultimo grafico.", "info")
    else:
        try:
            create_monthly_plot(profile, cached_balances(profile))
            flash("El grafico mensual se ha generado.", "success")
        except WorkbookError as error:
            flash(str(error), "error")
    return redirect(url_for("monthly_page", profile_slug=profile.slug))


@app.route("/<profile_slug>/plot-mensual.png")
def monthly_plot(profile_slug: str):
    profile = get_profile(profile_slug)
    if not profile.monthly_plot_path.exists():
        return "Aun no se ha generado el grafico mensual.", 404
    return send_file(profile.monthly_plot_path, mimetype="image/png", max_age=0)


@app.errorhandler(413)
def file_too_large(_error):
    flash("El archivo supera el limite de 20 MB.", "error")
    return redirect(url_for("dashboard", profile_slug="warein"))


if __name__ == "__main__":
    DATA_DIR.mkdir(exist_ok=True)
    app.run(host="127.0.0.1", port=8081, debug=False)

'''
from __future__ import annotations

import json
import os
from datetime import date, datetime, timedelta
from pathlib import Path

import matplotlib
import numpy as np
import pandas as pd
from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

matplotlib.use("Agg")
import matplotlib.pyplot as plt


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DEFAULT_WORKBOOK = BASE_DIR / "Movs_ing.xlsx"
WORKBOOK_PATH = DATA_DIR / "Movs_ing.xlsx"
BALANCES_PATH = DATA_DIR / "Saldo_Total.xlsx"
TOTAL_PLOT_PATH = DATA_DIR / "Movs_ing_total.png"
MONTHLY_PLOT_PATH = DATA_DIR / "Movs_ing_mensual.png"
MONTHLY_STATE_PATH = DATA_DIR / "monthly_plot_state.json"
REQUIRED_SHEET_INDICES = (0, 2, 3, 4)
AMOUNT_COLUMN = "IMPORTE (€)"
ACCOUNT_NAMES = ("ING", "Naranja", "Deposito", "Tesoro")
INITIAL_BALANCES = (5063.16, 0.0, 0.0, 0.0)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024
app.config["SECRET_KEY"] = os.environ.get("FLASK_SECRET_KEY", "local-private-dashboard")


class WorkbookError(ValueError):
    """Raised when the uploaded workbook cannot be processed."""


def active_workbook() -> Path:
    if WORKBOOK_PATH.exists():
        return WORKBOOK_PATH
    if DEFAULT_WORKBOOK.exists():
        return DEFAULT_WORKBOOK
    raise WorkbookError("No se encuentra Movs_ing.xlsx. Sube un archivo para continuar.")


def validate_workbook(path: Path) -> None:
    try:
        workbook = pd.ExcelFile(path)
    except Exception as error:
        raise WorkbookError("El archivo no es un libro Excel .xlsx valido.") from error

    if len(workbook.sheet_names) <= max(REQUIRED_SHEET_INDICES):
        raise WorkbookError("El libro debe incluir al menos cinco hojas.")

    for sheet_index in REQUIRED_SHEET_INDICES:
        frame = pd.read_excel(workbook, sheet_name=sheet_index, nrows=1)
        if frame.empty and not list(frame.columns):
            raise WorkbookError(f"La hoja {sheet_index + 1} no contiene encabezados.")
        if AMOUNT_COLUMN not in frame.columns:
            raise WorkbookError(
                f"La hoja {sheet_index + 1} debe incluir la columna '{AMOUNT_COLUMN}'."
            )


def calculate_balances() -> pd.DataFrame:
    DATA_DIR.mkdir(exist_ok=True)
    workbook_path = active_workbook()
    validate_workbook(workbook_path)
    sheets = pd.read_excel(workbook_path, sheet_name=list(REQUIRED_SHEET_INDICES))
    start = pd.Timestamp("2023-01-01")
    today = pd.Timestamp.today().normalize()
    days = pd.date_range(start, today, freq="D")
    balances = pd.DataFrame({"fecha": days})

    account_columns: list[str] = []
    for account_name, initial_balance, sheet_index in zip(
        ACCOUNT_NAMES, INITIAL_BALANCES, REQUIRED_SHEET_INDICES
    ):
        frame = sheets[sheet_index].copy()
        date_column = frame.columns[0]
        frame["fecha"] = pd.to_datetime(frame[date_column], errors="coerce").dt.normalize()
        frame[AMOUNT_COLUMN] = pd.to_numeric(frame[AMOUNT_COLUMN], errors="coerce").fillna(0)
        movements = frame.dropna(subset=["fecha"]).groupby("fecha")[AMOUNT_COLUMN].sum()
        daily_movements = movements.reindex(days, fill_value=0.0).to_numpy()

        # This matches the original script: the opening balance is fixed on 01/01/2023.
        account_balance = np.empty(len(days), dtype=float)
        account_balance[0] = initial_balance
        for position in range(1, len(days)):
            account_balance[position] = account_balance[position - 1] + daily_movements[position]

        column_name = f"saldo_{account_name.lower()}"
        balances[column_name] = account_balance
        account_columns.append(column_name)

    balances["saldo_total"] = balances[account_columns].sum(axis=1)
    balances.to_excel(BALANCES_PATH, index=False)
    return balances


def load_balances() -> pd.DataFrame:
    return calculate_balances()


def create_total_plot(balances: pd.DataFrame) -> Path:
    rolling_average = balances["saldo_total"].rolling(window=50, min_periods=1).mean()
    fig, axis = plt.subplots(figsize=(15, 9), dpi=100)
    axis.plot(balances["fecha"], balances["saldo_total"], alpha=0.2, linewidth=1)
    axis.plot(balances["fecha"], rolling_average, linewidth=2)
    axis.set_title(
        f"Saldo Total: {balances['saldo_total'].iloc[-1]:,.2f} EUR - "
        f"Actualizado: {datetime.now():%d/%m/%y %H:%M}",
        fontsize=14,
    )
    axis.set_xlabel("Fecha")
    axis.set_ylabel("Saldo (EUR)")
    axis.set_ylim(bottom=0)
    axis.set_xlim(pd.Timestamp("2023-10-01"), balances["fecha"].iloc[-1])
    axis.grid(True, linestyle="--", linewidth=0.9, alpha=0.3)
    for years_ago in range(1, 6):
        axis.axvline(datetime.now() - timedelta(days=365 * years_ago), linestyle="--", linewidth=1, alpha=0.6)
    fig.savefig(TOTAL_PLOT_PATH, bbox_inches="tight", dpi=100)
    plt.close(fig)
    return TOTAL_PLOT_PATH


def latest_completed_month() -> date:
    first_of_this_month = date.today().replace(day=1)
    return first_of_this_month - timedelta(days=1)


def monthly_plot_is_current() -> bool:
    if not MONTHLY_PLOT_PATH.exists() or not MONTHLY_STATE_PATH.exists():
        return False
    try:
        state = json.loads(MONTHLY_STATE_PATH.read_text())
        return state.get("through_month") == latest_completed_month().isoformat()
    except (OSError, json.JSONDecodeError):
        return False


def create_monthly_plot(balances: pd.DataFrame) -> Path:
    completed_month = latest_completed_month()
    monthly = balances.loc[balances["fecha"] <= pd.Timestamp(completed_month)].copy()
    monthly = monthly.loc[monthly["fecha"].dt.is_month_end].copy()
    if len(monthly) < 2:
        raise WorkbookError("Aun no hay dos meses completos para generar el grafico mensual.")

    monthly["ganancia"] = monthly["saldo_total"].diff()
    monthly = monthly.iloc[1:].copy()
    labels = monthly["fecha"].dt.strftime("%b %Y").str.capitalize()
    positions = np.arange(len(monthly))
    colors = np.where(monthly["ganancia"] >= 0, "#237a57", "#bb3e3e")

    fig, axis = plt.subplots(figsize=(12, 6), dpi=130)
    axis.bar(positions, monthly["ganancia"], color=colors, edgecolor="#283334", width=0.8)
    axis.set_xticks(positions, labels, rotation=45, ha="right")
    axis.set_title(f"Ganancia mensual - Actualizado: {datetime.now():%d/%m/%y %H:%M}")
    axis.set_xlabel("Mes completado")
    axis.set_ylabel("Ganancia (EUR)")
    axis.grid(axis="y", linestyle="--", alpha=0.45)
    fig.tight_layout()
    fig.savefig(MONTHLY_PLOT_PATH, dpi=200)
    plt.close(fig)
    MONTHLY_STATE_PATH.write_text(json.dumps({"through_month": completed_month.isoformat()}))
    return MONTHLY_PLOT_PATH


def chart_version(path: Path) -> int:
    return int(path.stat().st_mtime) if path.exists() else 0


@app.route("/")
def dashboard():
    try:
        balances = load_balances()
        latest = balances.iloc[-1]
        summary = {
            "date": latest["fecha"].strftime("%d %b %Y"),
            "total": latest["saldo_total"],
            "ing": latest["saldo_ing"],
            "naranja": latest["saldo_naranja"],
            "deposito": latest["saldo_deposito"],
            "tesoro": latest["saldo_tesoro"],
        }
        return render_template("dashboard.html", summary=summary, workbook=active_workbook().name)
    except WorkbookError as error:
        return render_template("dashboard.html", summary=None, workbook=None, error=str(error)), 400


@app.route("/upload", methods=["POST"])
def upload_workbook():
    upload = request.files.get("workbook")
    if not upload or not upload.filename:
        flash("Selecciona un archivo Movs_ing.xlsx.", "error")
        return redirect(url_for("dashboard"))
    if Path(secure_filename(upload.filename)).suffix.lower() != ".xlsx":
        flash("Solo se admiten archivos .xlsx.", "error")
        return redirect(url_for("dashboard"))

    DATA_DIR.mkdir(exist_ok=True)
    temporary_path = DATA_DIR / "upload-check.xlsx"
    upload.save(temporary_path)
    try:
        validate_workbook(temporary_path)
        temporary_path.replace(WORKBOOK_PATH)
        for generated_file in (BALANCES_PATH, TOTAL_PLOT_PATH, MONTHLY_PLOT_PATH, MONTHLY_STATE_PATH):
            generated_file.unlink(missing_ok=True)
        flash("Movs_ing.xlsx actualizado. Los saldos se han recalculado.", "success")
    except WorkbookError as error:
        temporary_path.unlink(missing_ok=True)
        flash(str(error), "error")
    return redirect(url_for("dashboard"))


@app.route("/saldos")
def balances_page():
    try:
        balances = load_balances().iloc[::-1].copy()
        balances["fecha"] = balances["fecha"].dt.strftime("%d/%m/%Y")
        return render_template("balances.html", rows=balances.head(180).to_dict("records"))
    except WorkbookError as error:
        flash(str(error), "error")
        return redirect(url_for("dashboard"))


@app.route("/generar-total", methods=["POST"])
def generate_total_plot():
    try:
        create_total_plot(load_balances())
        flash("El grafico de saldo total se ha generado.", "success")
    except WorkbookError as error:
        flash(str(error), "error")
    return redirect(url_for("dashboard"))


@app.route("/plot-total.png")
def total_plot():
    if not TOTAL_PLOT_PATH.exists():
        return "Genera primero el grafico de saldo total.", 404
    return send_file(TOTAL_PLOT_PATH, mimetype="image/png", max_age=0)


@app.route("/mensual")
def monthly_page():
    return render_template(
        "monthly.html",
        has_plot=MONTHLY_PLOT_PATH.exists(),
        is_current=monthly_plot_is_current(),
        version=chart_version(MONTHLY_PLOT_PATH),
    )


@app.route("/generar-mensual", methods=["POST"])
def generate_monthly_plot():
    if monthly_plot_is_current():
        flash("No hay ningun mes completo nuevo desde el ultimo grafico.", "info")
    else:
        try:
            create_monthly_plot(load_balances())
            flash("El grafico mensual se ha generado.", "success")
        except WorkbookError as error:
            flash(str(error), "error")
    return redirect(url_for("monthly_page"))


@app.route("/plot-mensual.png")
def monthly_plot():
    if not MONTHLY_PLOT_PATH.exists():
        return "Aun no se ha generado el grafico mensual.", 404
    return send_file(MONTHLY_PLOT_PATH, mimetype="image/png", max_age=0)


@app.errorhandler(413)
def file_too_large(_error):
    flash("El archivo supera el limite de 20 MB.", "error")
    return redirect(url_for("dashboard"))


if __name__ == "__main__":
    DATA_DIR.mkdir(exist_ok=True)
    app.run(host="127.0.0.1", port=8081, debug=False)
'''