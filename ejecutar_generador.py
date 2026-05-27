#!/usr/bin/env python3
"""
Script de ejecución del generador de informes de posgrados
Ejecuta el análisis y genera el informe Word automáticamente
"""

import sys
import os
import subprocess
from pathlib import Path


def main():
    """Ejecuta el generador de informes"""

    # Rutas de archivos
    repo_root = Path(__file__).parent
    script_path = repo_root / "scripts" / "generar_informe_posgrados.py"
    excel_path = repo_root / "ENCUESTA SEGUIMIENTO EGRESADOS 30-04-2026.xlsx"
    plantilla_path = repo_root / "INTERES EN APERTURA DE PROGRAMAS ACADEMICOS UNIPUTUMAYO.docx"
    output_path = repo_root / "informes_generados" / "INFORME_POSGRADOS_2026.docx"

    # Verificar que existen los archivos
    if not excel_path.exists():
        print(f"❌ Error: No se encontró el archivo Excel en: {excel_path}")
        return 1

    if not plantilla_path.exists():
        print(f"❌ Error: No se encontró la plantilla Word en: {plantilla_path}")
        return 1

    if not script_path.exists():
        print(f"❌ Error: No se encontró el script en: {script_path}")
        return 1

    print("✅ Archivos verificados correctamente")
    print(f"📊 Excel: {excel_path.name}")
    print(f"📄 Plantilla: {plantilla_path.name}")
    print()

    # Crear carpeta de salida si no existe
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Ejecutar el generador
    print("🚀 Ejecutando generador de informes...")
    print()

    cmd = [
        sys.executable,
        str(script_path),
        "--excel", str(excel_path),
        "--plantilla", str(plantilla_path),
        "--output", str(output_path),
    ]

    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print(result.stdout)

        if output_path.exists():
            print()
            print("=" * 60)
            print("✅ ¡INFORME GENERADO EXITOSAMENTE!")
            print("=" * 60)
            print()
            print(f"📁 Ubicación: {output_path}")
            print(f"📏 Tamaño: {output_path.stat().st_size / 1024:.2f} KB")
            print()
            print("🎯 Próximos pasos:")
            print("   1. Descarga el archivo desde 'informes_generados/'")
            print("   2. Abre el documento Word")
            print("   3. Revisa tablas, gráficos y conclusiones")
            print("   4. Edita manualmente si es necesario")
            print()
            return 0
        else:
            print("❌ Error: El informe no se generó correctamente")
            return 1

    except subprocess.CalledProcessError as e:
        print("❌ Error al ejecutar el generador:")
        print(e.stderr)
        return 1
    except Exception as e:
        print(f"❌ Error inesperado: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
