"""Ejecuta toda la automatización financiera en el orden correcto.

Flujo completo:
1. Inspeccionar y validar los archivos de entrada.
2. Normalizar billings, créditos bancarios y tipo de cambio.
3. Conciliar billings contra créditos bancarios.
4. Generar el reporte final en Excel.
"""

from pathlib import Path
import subprocess
import sys


# Obtenemos la carpeta donde están todos los scripts del proceso.
SCRIPT_DIR = Path(__file__).resolve().parent

# Definimos el orden obligatorio de ejecución.
STEPS = [
    "01_inspect_inputs.py",
    "02_normalize_inputs.py",
    "03_reconcile_billings.py",
    "04_generate_excel_report.py",
]


def main() -> None:
    # Ejecutamos cada script usando el mismo intérprete de Python.
    for step_number, script_name in enumerate(STEPS, start=1):
        script_path = SCRIPT_DIR / script_name
        print(f"\nPaso {step_number}/{len(STEPS)}: ejecutando {script_name}")

        # check=True detiene todo el proceso si un paso genera un error.
        subprocess.run(
            [sys.executable, str(script_path)],
            check=True,
        )

        print(f"Paso {step_number} completado correctamente.")

    print("\nAutomatización completa. El reporte está en la carpeta results.")


if __name__ == "__main__":
    main()

