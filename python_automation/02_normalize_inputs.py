from pathlib import Path
import pandas as pd


INPUT_DIR = Path(r"C:\Users\Usuario\OneDrive\Escritorio")
WORK_DIR = Path(__file__).resolve().parent / "normalized"


def read_billing() -> pd.DataFrame:
    df = pd.read_excel(INPUT_DIR / "Monthly_Billing.xlsx")
    df = df.rename(columns={
        "Cliente": "cliente",
        "Descripción": "descripcion",
        "Total billing": "billing_usd",
    })
    df["cliente"] = pd.to_numeric(df["cliente"], errors="coerce").astype("Int64")
    df["billing_usd"] = pd.to_numeric(df["billing_usd"], errors="coerce")
    df["moneda"] = "USD"
    df["billing_id"] = range(1, len(df) + 1)
    return df[["billing_id", "cliente", "descripcion", "billing_usd", "moneda"]]


def read_fx() -> pd.DataFrame:
    df = pd.read_excel(INPUT_DIR / "Tipo de Cambio.xlsx", header=1)
    df = df.rename(columns={"Fecha": "fecha", "Tipo de cambio": "tipo_cambio"})
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce").dt.normalize()
    df["tipo_cambio"] = pd.to_numeric(df["tipo_cambio"], errors="coerce")
    # Si existe más de una tasa para un mismo día, se conserva la última registrada.
    return df.dropna(subset=["fecha", "tipo_cambio"]).drop_duplicates("fecha", keep="last")


def read_bank(path: Path, moneda_banco: str) -> pd.DataFrame:
    # Los estados tienen el título del banco en la primera fila y los encabezados en la segunda.
    df = pd.read_excel(path, header=1)
    df = df.rename(columns={
        "Fecha de transacción": "fecha",
        "Código del cliente": "cliente",
        "Descripción": "descripcion",
        "Debito": "debito",
        "Crédito": "credito",
    })
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce").dt.normalize()
    df["cliente"] = pd.to_numeric(df["cliente"], errors="coerce").astype("Int64")
    df["credito"] = pd.to_numeric(df["credito"], errors="coerce")
    # Regla confirmada: únicamente los créditos son ingresos conciliables.
    df = df[df["credito"].notna()].copy()
    df["moneda_banco"] = moneda_banco
    df["monto_original"] = df["credito"]
    df["monto_usd"] = df["credito"]
    return df[["fecha", "cliente", "descripcion", "moneda_banco", "monto_original", "monto_usd"]]


def main() -> None:
    WORK_DIR.mkdir(parents=True, exist_ok=True)

    billing = read_billing()
    fx = read_fx()
    bank_usd = read_bank(INPUT_DIR / "EstadoBancario_BAC_dolares.xlsx", "USD")
    bank_crc = read_bank(INPUT_DIR / "EstadoBancario_BAC_colones.xlsx", "CRC")

    bank_crc = bank_crc.merge(fx, on="fecha", how="left", validate="many_to_one")
    missing_fx = bank_crc[bank_crc["tipo_cambio"].isna()]
    if not missing_fx.empty:
        raise ValueError(f"Faltan tipos de cambio para estas fechas: {missing_fx['fecha'].dt.date.tolist()}")
    bank_crc["monto_usd"] = bank_crc["monto_original"] / bank_crc["tipo_cambio"]
    bank_crc = bank_crc.drop(columns=["tipo_cambio"])

    bank = pd.concat([bank_usd, bank_crc], ignore_index=True)
    bank["movimiento_id"] = range(1, len(bank) + 1)
    bank = bank[["movimiento_id", "fecha", "cliente", "descripcion", "moneda_banco", "monto_original", "monto_usd"]]

    billing.to_csv(WORK_DIR / "billing_normalized.csv", index=False, encoding="utf-8-sig")
    bank.to_csv(WORK_DIR / "bank_credits_normalized.csv", index=False, encoding="utf-8-sig")
    fx.to_csv(WORK_DIR / "fx_normalized.csv", index=False, encoding="utf-8-sig")

    print(f"Billings USD: {len(billing)}")
    print(f"Créditos bancarios USD: {len(bank_usd)}")
    print(f"Créditos bancarios CRC convertidos: {len(bank_crc)}")
    print(f"Total de créditos conciliables: {len(bank)}")
    print("Archivos normalizados:", WORK_DIR)


if __name__ == "__main__":
    main()

