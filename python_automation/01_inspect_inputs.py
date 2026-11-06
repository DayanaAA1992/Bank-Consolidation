from pathlib import Path
import pandas as pd


INPUT_DIR = Path(r"C:\Users\Usuario\OneDrive\Escritorio")

FILES = {
    "billing": INPUT_DIR / "Monthly_Billing.xlsx",
    "bank_crc": INPUT_DIR / "EstadoBancario_BAC_colones.xlsx",
    "bank_usd": INPUT_DIR / "EstadoBancario_BAC_dolares.xlsx",
    "fx": INPUT_DIR / "Tipo de Cambio.xlsx",
}


def inspect_file(label: str, path: Path) -> None:
    print(f"\n{'=' * 70}\n{label}: {path}")

    if not path.exists():
        print("ERROR: archivo no encontrado")
        return

    workbook = pd.ExcelFile(path)
    print(f"Hojas: {workbook.sheet_names}")

    for sheet_name in workbook.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet_name)
        print(f"\nHoja: {sheet_name}")
        print(f"Filas: {len(df)} | Columnas: {len(df.columns)}")
        print(f"Columnas: {list(df.columns)}")
        print("Tipos de datos:")
        print(df.dtypes.to_string())
        print("Primeras filas:")
        print(df.head(5).to_string(index=False))

        missing = df.isna().sum()
        missing = missing[missing > 0]
        if not missing.empty:
            print("Valores vacíos:")
            print(missing.to_string())


def main() -> None:
    print("Inspección inicial de archivos para conciliación bancaria")
    print("Regla de negocio confirmada: Monthly Billing está completamente en USD.")

    for label, path in FILES.items():
        inspect_file(label, path)


if __name__ == "__main__":
    main()

