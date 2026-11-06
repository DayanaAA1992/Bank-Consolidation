"""Genera el reporte final de conciliación bancaria.

Este archivo está documentado paso a paso para que el proceso sea mantenible
por otra persona del equipo financiero.
"""

from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# Paso 1: definir las carpetas de entrada y salida.
# Los CSV fueron generados por 02_normalize_inputs.py y 03_reconcile_billings.py.
BASE_DIR = Path(__file__).resolve().parent
RESULTS_DIR = BASE_DIR / "results"
REPORT_PATH = RESULTS_DIR / "reporte_conciliacion_bancaria.xlsx"


def load_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Paso 2: cargar los resultados de conciliación y créditos no identificados."""
    reconciliation = pd.read_csv(RESULTS_DIR / "billing_reconciliation.csv")
    unmatched = pd.read_csv(RESULTS_DIR / "unmatched_bank_credits.csv")
    return reconciliation, unmatched


def build_summary(reconciliation: pd.DataFrame, unmatched: pd.DataFrame) -> pd.DataFrame:
    """Paso 3: construir indicadores para la hoja Resumen."""
    # Contamos las billings por estado.
    status_counts = reconciliation["estado"].value_counts()

    # Sumamos los montos de billing en USD por estado.
    amount_by_status = reconciliation.groupby("estado", dropna=False)["billing_usd"].sum()

    # Creamos una tabla simple de indicadores para lectura ejecutiva.
    rows = [
        ["Indicador", "Valor"],
        ["Total de billings", len(reconciliation)],
        ["Billings pagadas", int(status_counts.get("Pagada", 0))],
        ["Billings pendientes", int(status_counts.get("Pendiente", 0))],
        ["Billings con diferencia", int(status_counts.get("Diferencia", 0))],
        ["Monto total facturado (USD)", float(reconciliation["billing_usd"].sum())],
        ["Monto pagado (USD)", float(amount_by_status.get("Pagada", 0))],
        ["Monto pendiente (USD)", float(amount_by_status.get("Pendiente", 0))],
        ["Créditos bancarios no identificados", len(unmatched)],
    ]
    return pd.DataFrame(rows[1:], columns=rows[0])


def prepare_for_excel(reconciliation: pd.DataFrame, unmatched: pd.DataFrame):
    """Paso 4: preparar nombres y formatos legibles para el reporte."""
    reconciliation = reconciliation.rename(columns={
        "billing_id": "ID billing",
        "cliente": "Código cliente",
        "descripcion": "Descripción",
        "billing_usd": "Billing USD",
        "estado": "Estado",
        "movimiento_id": "ID movimiento bancario",
        "fecha_pago": "Fecha pago",
        "moneda_banco": "Moneda banco",
        "monto_banco_original": "Monto banco original",
        "monto_banco_usd": "Monto banco USD",
        "diferencia_usd": "Diferencia USD",
    })
    unmatched = unmatched.rename(columns={
        "movimiento_id": "ID movimiento bancario",
        "fecha": "Fecha",
        "cliente": "Código cliente",
        "descripcion": "Descripción",
        "moneda_banco": "Moneda banco",
        "monto_original": "Monto original",
        "monto_usd": "Monto USD",
    })
    return reconciliation, unmatched


def format_workbook(path: Path) -> None:
    """Paso 5: aplicar formato financiero y filtros al archivo Excel."""
    workbook = load_workbook(path)
    navy = "003B70"
    blue = "0B5A8F"
    light_blue = "D9EAF7"
    white_font = Font(color="FFFFFF", bold=True)
    thin_blue = Side(style="thin", color="B4C7E7")

    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        sheet.sheet_view.showGridLines = False
        sheet.auto_filter.ref = sheet.dimensions

        # Formateamos la primera fila como encabezado de tabla.
        for cell in sheet[1]:
            cell.fill = PatternFill("solid", fgColor=blue)
            cell.font = white_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Ajustamos anchos y bordes para mejorar legibilidad.
        for column_cells in sheet.columns:
            column_letter = get_column_letter(column_cells[0].column)
            max_length = max(len(str(cell.value or "")) for cell in column_cells)
            sheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 42)
            for cell in column_cells:
                cell.border = Border(bottom=thin_blue)
                cell.alignment = Alignment(vertical="center")

        # Aplicamos formatos numéricos según el contenido de cada hoja.
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, (int, float)) and cell.column > 1:
                    cell.number_format = '$#,##0.00'

        # Las columnas de códigos e IDs son enteros, no montos monetarios.
        for header in ["ID billing", "Código cliente", "ID movimiento bancario", "Créditos bancarios no identificados"]:
            for cell in sheet[1]:
                if cell.value == header:
                    for data_cell in sheet.iter_cols(min_col=cell.column, max_col=cell.column, min_row=2):
                        for item in data_cell:
                            item.number_format = '0'

        # El monto original debe conservar la moneda del banco de origen.
        # Los montos convertidos y las diferencias siempre se muestran en USD.
        header_map = {cell.value: cell.column for cell in sheet[1]}
        original_col = header_map.get("Monto banco original") or header_map.get("Monto original")
        usd_col = header_map.get("Monto banco USD") or header_map.get("Monto USD")
        difference_col = header_map.get("Diferencia USD")
        currency_col = header_map.get("Moneda banco")
        if original_col and currency_col:
            for row in range(2, sheet.max_row + 1):
                currency = sheet.cell(row, currency_col).value
                sheet.cell(row, original_col).number_format = (
                    '"₡"#,##0.00' if currency == "CRC" else '$#,##0.00'
                )
        for col in [usd_col, difference_col]:
            if col:
                for row in range(2, sheet.max_row + 1):
                    sheet.cell(row, col).number_format = '$#,##0.00'

        # En la hoja Resumen los indicadores de conteo son números enteros,
        # mientras que únicamente los indicadores expresamente marcados como
        # USD deben tener formato monetario.
        if sheet.title == "Resumen":
            for row in range(2, sheet.max_row + 1):
                indicator = str(sheet.cell(row, 1).value or "")
                value_cell = sheet.cell(row, 2)
                if "USD" in indicator:
                    value_cell.number_format = '$#,##0.00'
                else:
                    value_cell.number_format = '0'

    # Resaltamos estados para facilitar la revisión financiera.
    if "Conciliación" in workbook.sheetnames:
        sheet = workbook["Conciliación"]
        headers = {cell.value: cell.column for cell in sheet[1]}
        state_col = headers.get("Estado")
        if state_col:
            for row in range(2, sheet.max_row + 1):
                state = sheet.cell(row, state_col).value
                if state == "Pagada":
                    sheet.cell(row, state_col).fill = PatternFill("solid", fgColor="C6EFCE")
                elif state == "Pendiente":
                    sheet.cell(row, state_col).fill = PatternFill("solid", fgColor="FFF2CC")
                elif state == "Diferencia":
                    sheet.cell(row, state_col).fill = PatternFill("solid", fgColor="F4CCCC")

    workbook.save(path)


def main() -> None:
    # Paso 6: leer los datos ya conciliados.
    reconciliation, unmatched = load_data()

    # Paso 7: construir el resumen y renombrar campos para el usuario final.
    summary = build_summary(reconciliation, unmatched)
    reconciliation, unmatched = prepare_for_excel(reconciliation, unmatched)

    # Paso 8: exportar cuatro hojas a un único archivo Excel.
    # Cada hoja sirve para una revisión distinta del proceso.
    with pd.ExcelWriter(REPORT_PATH, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Resumen", index=False)
        reconciliation.to_excel(writer, sheet_name="Conciliación", index=False)
        reconciliation[reconciliation["Estado"] == "Pendiente"].to_excel(
            writer, sheet_name="Pendientes", index=False
        )
        unmatched.to_excel(writer, sheet_name="Créditos no identificados", index=False)

    # Paso 9: aplicar formato después de escribir los datos.
    format_workbook(REPORT_PATH)
    print(f"Reporte generado: {REPORT_PATH}")


if __name__ == "__main__":
    main()

