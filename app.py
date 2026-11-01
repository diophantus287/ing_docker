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
CACHE_VERSION = 4


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
        return DATA_DIR

    @property
    def workbook_path(self) -> Path:
        return self.directory / self.workbook_name

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
        return self.directory / f"{self.slug}_full"

    def source_csv_path(self, sheet_name: str) -> Path:
        filename = secure_filename(sheet_name).lower().replace("_", "-")
        return self.source_csv_directory / f"{filename}.csv"


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
        start_date="2023-01-01",
    ),
    "isa": Profile(
        slug="isa",
        label="Isa",
        workbook_name="Excel_isa.xlsx",
        sheet_indices=(0,),
        accounts=(("saldo_isa", "Isa", 1915.89),),
        start_date=None,
    ),
}

WAREIN_MOVEMENT_SOURCES = {
    "ing-nomina": {
        "label": "ING nómina",
        "balance_column": "saldo_ing",
        "filename": "ing-nomina.csv",
    },
    "ing-naranja": {
        "label": "ING Naranja",
        "balance_column": "saldo_naranja",
        "filename": "ing-naranja.csv",
    },
    "trade": {
        "label": "Trade",
        "balance_column": "saldo_trade",
        "filename": "trade.csv",
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
        if profile.slug == "isa":
            frame = pd.read_excel(workbook, sheet_name=sheet_index, header=None, nrows=1)
            if len(frame.columns) < 4:
                raise WorkbookError("La primera hoja de Isa debe tener al menos cuatro columnas.")
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
    if profile.slug == "isa":
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
            return pd.read_csv(warein_movement_path(str(sheets[sheet_index]["account_key"])), keep_default_na=False)
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
        if column == "F. VALOR":
            values[column] = pd.to_datetime(values[column], errors="coerce").dt.strftime("%Y-%m-%d")
    return values.fillna("").astype(str).agg("\x1f".join, axis=1)


def merge_new_movements(existing: pd.DataFrame, uploaded: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    all_columns = list(dict.fromkeys([*existing.columns, *uploaded.columns]))
    existing = existing.reindex(columns=all_columns)
    uploaded = uploaded.reindex(columns=all_columns)
    fingerprint_columns = [column for column in all_columns if column != "SALDO (€)"]
    existing_counts = movement_fingerprints(existing, fingerprint_columns).value_counts()
    uploaded_fingerprints = movement_fingerprints(uploaded, fingerprint_columns)
    seen_counts: dict[str, int] = {}
    new_rows: list[int] = []
    for index, fingerprint in uploaded_fingerprints.items():
        occurrence = seen_counts.get(fingerprint, 0)
        seen_counts[fingerprint] = occurrence + 1
        if occurrence >= existing_counts.get(fingerprint, 0):
            new_rows.append(index)

    merged = (
        pd.concat([existing, uploaded.loc[new_rows]], ignore_index=True)
        if new_rows
        else existing.copy()
    )
    merged = normalize_movement_frame(merged).sort_values("F. VALOR", ascending=False, kind="stable")
    return merged, len(new_rows)


def calculate_balances(profile: Profile) -> pd.DataFrame:
    profile.directory.mkdir(parents=True, exist_ok=True)
    prepared_sheets: list[pd.DataFrame] = []
    first_date: pd.Timestamp | None = None

    if profile.slug == "warein":
        source_frames = [
            pd.read_csv(warein_movement_path(account_key))
            for account_key in WAREIN_MOVEMENT_SOURCES
        ]
    else:
        sheets = source_sheets(profile)
        source_frames = [sheets[sheet_index].copy() for sheet_index in profile.sheet_indices]

    for frame in source_frames:
        if profile.slug == "isa":
            frame["fecha"] = pd.to_datetime(frame.iloc[:, 0], errors="coerce", dayfirst=True).dt.normalize()
            frame[AMOUNT_COLUMN] = parse_amounts(frame.iloc[:, 3])
        else:
            frame = normalize_movement_frame(frame)
            frame["fecha"] = frame["F. VALOR"]
            frame[AMOUNT_COLUMN] = parse_amounts(frame[AMOUNT_COLUMN])
        frame = frame.dropna(subset=["fecha"])
        prepared_sheets.append(frame)
        if not frame.empty:
            sheet_start = frame["fecha"].min()
            first_date = sheet_start if first_date is None else min(first_date, sheet_start)

    if first_date is None:
        raise WorkbookError("El Excel no contiene movimientos con una fecha valida.")

    start = pd.Timestamp(profile.start_date) if profile.start_date else first_date
    today = pd.Timestamp.today().normalize()
    days = pd.date_range(start, today, freq="D")
    balances = pd.DataFrame({"fecha": days})
    account_columns: list[str] = []

    for account, frame in zip(profile.accounts, prepared_sheets):
        column_name, _label, initial_balance = account
        if profile.slug == "warein":
            if "SALDO (€)" not in frame.columns:
                raise WorkbookError("El CSV de ING debe incluir la columna 'SALDO (€)'.")
            frame["saldo_banco"] = parse_amounts(frame["SALDO (€)"])
            closing_balances = (
                frame.drop_duplicates(subset=["fecha"], keep="first")
                .set_index("fecha")["saldo_banco"]
            )
            balances[column_name] = closing_balances.reindex(days).ffill().bfill().to_numpy()
        else:
            movements = frame.groupby("fecha")[AMOUNT_COLUMN].sum()
            daily_movements = movements.reindex(days, fill_value=0.0).to_numpy()
            balances[column_name] = initial_balance + np.cumsum(daily_movements)
        account_columns.append(column_name)

    balances["saldo_total"] = balances[account_columns].sum(axis=1)
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
    if workbook_is_newer(profile, profile.balances_path) or workbook_is_newer(profile, profile.summary_path):
        return True
    try:
        cache_version = pd.read_csv(profile.summary_path, nrows=1).iloc[0].get("cache_version")
        return int(cache_version) != CACHE_VERSION
    except (OSError, ValueError, TypeError, pd.errors.EmptyDataError):
        return True


def refresh_cached_data(profile: Profile) -> pd.DataFrame:
    balances = calculate_balances(profile)
    create_total_plot(profile, balances)
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
    return "#216c57" if profile.slug == "warein" else "#b54a3a"


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
        return render_template(
            "dashboard.html",
            profile=profile,
            summary=cached_summary(profile),
            workbook=active_workbook(profile).name,
        )
    except WorkbookError as error:
        return render_template("dashboard.html", profile=profile, summary=None, workbook=None, error=str(error)), 400


@app.route("/<profile_slug>/upload", methods=["POST"])
def upload_workbook(profile_slug: str):
    profile = get_profile(profile_slug)
    upload = request.files.get("workbook")
    if not upload or not upload.filename:
        flash("Selecciona un archivo Excel.", "error")
        return redirect(url_for("dashboard", profile_slug=profile.slug))
    suffix = Path(secure_filename(upload.filename)).suffix.lower()
    allowed_extensions = {".xls", ".xlsx"} if profile.slug == "warein" else {".xlsx"}
    if suffix not in allowed_extensions:
        flash("Solo se admiten archivos Excel válidos.", "error")
        return redirect(url_for("dashboard", profile_slug=profile.slug))

    profile.directory.mkdir(parents=True, exist_ok=True)
    temporary_path = profile.directory / f"upload-check{suffix}"
    upload.save(temporary_path)
    try:
        if profile.slug == "warein":
            account_key = request.form.get("account")
            target_path = warein_movement_path(account_key or "")
            uploaded_movements = read_bank_movements(temporary_path)
            try:
                existing_movements = normalize_movement_frame(pd.read_csv(target_path))
            except (OSError, pd.errors.EmptyDataError) as error:
                raise WorkbookError("No se ha podido leer el CSV de movimientos existente.") from error

            updated_movements, new_count = merge_new_movements(existing_movements, uploaded_movements)
            target_path.parent.mkdir(parents=True, exist_ok=True)
            temporary_csv_path = target_path.with_suffix(".tmp")
            updated_movements.to_csv(temporary_csv_path, index=False, date_format="%Y-%m-%d")
            temporary_csv_path.replace(target_path)
            temporary_path.unlink(missing_ok=True)
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
            account_label = WAREIN_MOVEMENT_SOURCES[account_key]["label"]
            flash(
                f"{account_label} actualizado: {new_count} movimientos nuevos. "
                "Los saldos, el resumen y los gráficos se han regenerado.",
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
    if profile.slug != "warein":
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
        )
    except WorkbookError as error:
        flash(str(error), "error")
        return redirect(url_for("dashboard", profile_slug=profile.slug))


@app.route("/<profile_slug>/total/<int:sheet_index>.csv")
def source_sheet_csv(profile_slug: str, sheet_index: int):
    profile = get_profile(profile_slug)
    if profile.slug != "warein":
        abort(404)
    sheets = source_sheet_metadata(profile)
    if sheet_index < 0 or sheet_index >= len(sheets):
        abort(404)
    sheet_name = str(sheets[sheet_index]["name"])
    load_source_sheet(profile, sheet_index)
    source_path = (
        warein_movement_path(str(sheets[sheet_index]["account_key"]))
        if profile.slug == "warein"
        else profile.source_csv_path(sheet_name)
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