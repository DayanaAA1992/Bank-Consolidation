from pathlib import Path
import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "normalized"
OUTPUT_DIR = BASE_DIR / "results"
TOLERANCE_USD = 0.05


def reconcile() -> tuple[pd.DataFrame, pd.DataFrame]:
    billings = pd.read_csv(DATA_DIR / "billing_normalized.csv")
    bank = pd.read_csv(DATA_DIR / "bank_credits_normalized.csv")

    billings["billing_usd"] = pd.to_numeric(billings["billing_usd"], errors="coerce")
    bank["monto_usd"] = pd.to_numeric(bank["monto_usd"], errors="coerce")
    billings["cliente"] = pd.to_numeric(billings["cliente"], errors="coerce").astype("Int64")
    bank["cliente"] = pd.to_numeric(bank["cliente"], errors="coerce").astype("Int64")
    bank["usado"] = False

    results = []
    for _, billing in billings.iterrows():
        candidates = bank[
            (~bank["usado"])
            & (bank["cliente"] == billing["cliente"])
        ].copy()
        candidates["diferencia_usd"] = (candidates["monto_usd"] - billing["billing_usd"]).abs()
        candidates = candidates[candidates["diferencia_usd"] <= TOLERANCE_USD]

        if candidates.empty:
            results.append({
                **billing.to_dict(),
                "estado": "Pendiente",
                "movimiento_id": None,
                "fecha_pago": None,
                "moneda_banco": None,
                "monto_banco_original": None,
                "monto_banco_usd": None,
                "diferencia_usd": None,
            })
            continue

        best_index = candidates["diferencia_usd"].idxmin()
        matched = bank.loc[best_index]
        bank.loc[best_index, "usado"] = True
        difference = float(matched["monto_usd"] - billing["billing_usd"])
        results.append({
            **billing.to_dict(),
            "estado": "Pagada" if abs(difference) <= TOLERANCE_USD else "Diferencia",
            "movimiento_id": int(matched["movimiento_id"]),
            "fecha_pago": matched["fecha"],
            "moneda_banco": matched["moneda_banco"],
            "monto_banco_original": matched["monto_original"],
            "monto_banco_usd": matched["monto_usd"],
            "diferencia_usd": round(difference, 2),
        })

    reconciled = pd.DataFrame(results)
    unmatched_bank = bank[~bank["usado"]].drop(columns=["usado"])
    return reconciled, unmatched_bank


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    reconciled, unmatched_bank = reconcile()

    reconciled.to_csv(OUTPUT_DIR / "billing_reconciliation.csv", index=False, encoding="utf-8-sig")
    unmatched_bank.to_csv(OUTPUT_DIR / "unmatched_bank_credits.csv", index=False, encoding="utf-8-sig")

    print("Resumen de conciliación")
    print(reconciled["estado"].value_counts(dropna=False).to_string())
    print(f"Créditos bancarios no identificados: {len(unmatched_bank)}")
    print(f"Resultados guardados en: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()

