#!/usr/bin/env python3
from __future__ import annotations

import argparse
import subprocess
import sys
from copy import copy
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DROPBOX_UPLOADER = Path("/home/robin/scripts/dropboxup/dropbox_uploader.sh")
DROPBOX_DESTINATION = "/Mi Drop/0006 Excels Bancarios/"
WAREIN_SOURCES = (
    ("ING Nomina", DATA_DIR / "warein" / "warein_ing-nomina.csv"),
    ("ING Naranja", DATA_DIR / "warein" / "warein_ing-naranja.csv"),
    ("Trade", DATA_DIR / "warein" / "warein_trade.csv"),
)


def read_csv(path: Path) -> pd.DataFrame:
    try:
        return pd.read_csv(path, keep_default_na=False)
    except FileNotFoundError as error:
        raise RuntimeError(f"No se encuentra el CSV requerido: {path}") from error
    except (OSError, pd.errors.EmptyDataError, pd.errors.ParserError) as error:
        raise RuntimeError(f"No se ha podido leer el CSV: {path}") from error


def add_balance_check_columns(
    frame: pd.DataFrame, amount_column: str, balance_column: str
) -> pd.DataFrame:
    if frame.empty:
        frame["SALDO CALCULADO"] = pd.Series(dtype=float)
        frame["COMPROBACIÓN"] = pd.Series(dtype=str)
        return frame

    balances = pd.to_numeric(frame[balance_column], errors="raise")
    amounts = pd.to_numeric(frame[amount_column], errors="raise")
    calculated_balances = np.empty(len(frame))
    calculated_balances[-1] = balances.iloc[-1]
    if len(frame) > 1:
        calculated_balances[:-1] = (
            balances.iloc[-1] + amounts.iloc[:-1].iloc[::-1].to_numpy().cumsum()[::-1]
        )

    frame = frame.copy()
    frame["SALDO CALCULADO"] = np.round(calculated_balances, 2)
    frame["COMPROBACIÓN"] = np.where(
        np.isclose(balances.to_numpy(), calculated_balances, rtol=0, atol=0.01),
        "OK",
        "ERROR",
    )
    return frame


def load_bank_movements(filename: str) -> pd.DataFrame:
    return add_balance_check_columns(
        read_csv(DATA_DIR / filename), "IMPORTE EUR", "SALDO"
    )


def fit_column_widths(workbook_path: Path) -> None:
    workbook = load_workbook(workbook_path)
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                font = copy(cell.font)
                font.sz = 8
                cell.font = font
        for column in worksheet.columns:
            header = column[0].value
            worksheet.column_dimensions[column[0].column_letter].width = (
                70 if header in {"DESCRIPCIÓN", "CONCEPTO"} else 12
            )
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
    workbook.save(workbook_path)


def export_movements(output_directory: Path) -> Path:
    output_directory.mkdir(parents=True, exist_ok=True)
    output_path = output_directory / f"movimientos_export.xlsx"
    sheets = {
        sheet_name: add_balance_check_columns(
            read_csv(path), "IMPORTE (€)", "SALDO (€)"
        )
        for sheet_name, path in WAREIN_SOURCES
    } | {
        "Isa": load_bank_movements("isa/isa_movements.csv"),
        "Hipoteca": load_bank_movements("hipoteca/hipoteca_movements.csv"),
    }
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for sheet_name, movements in sheets.items():
                movements.to_excel(writer, sheet_name=sheet_name, index=False)
        fit_column_widths(output_path)
    except OSError as error:
        raise RuntimeError(f"No se ha podido crear el Excel: {output_path}") from error
    return output_path


def upload_to_dropbox(workbook_path: Path) -> None:
    if not DROPBOX_UPLOADER.is_file():
        raise RuntimeError(f"No se encuentra el cliente de Dropbox: {DROPBOX_UPLOADER}")
    try:
        subprocess.run(
            [str(DROPBOX_UPLOADER), "upload", str(workbook_path), DROPBOX_DESTINATION],
            check=True,
        )
    except subprocess.CalledProcessError as error:
        raise RuntimeError("No se ha podido subir el Excel a Dropbox.") from error


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Exporta los movimientos de Warein, Isa e Hipoteca a un único Excel."
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=BASE_DIR / "exports",
        help="Directorio en el que se sobrescribirá movimientos_export.xlsx.",
    )
    parser.add_argument(
        "--skip-upload",
        action="store_true",
        help="Genera el Excel sin subirlo a Dropbox.",
    )
    return parser.parse_args()


def main() -> int:
    arguments = parse_arguments()
    try:
        workbook_path = export_movements(arguments.output_dir)
        if not arguments.skip_upload:
            upload_to_dropbox(workbook_path)
    except RuntimeError as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1

    if arguments.skip_upload:
        print(f"Excel generado: {workbook_path}")
    else:
        print("Gráfico y Excel subidos a Dropbox.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
