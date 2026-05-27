# acreditacion-uniputumayo
"Análisis de lineamientos CESU 2025 y generación de encuestas institucionales"

## 📊 Generación Automática de Informes

Este repositorio incluye un sistema automatizado para generar informes institucionales.

### 🚀 Generar Informe de Posgrados

**Opción 1: Ejecución automática (recomendada)**

1. Ve a [Actions](../../actions/workflows/generar-informe.yml)
2. Haz clic en **Run workflow**
3. Descarga el informe generado desde **Artifacts**

**Opción 2: Ejecución local**

```bash
git clone https://github.com/juliancourse07/acreditacion-uniputumayo.git
cd acreditacion-uniputumayo
pip install -r scripts/requirements.txt
python ejecutar_generador.py
```

El informe se generará en: `informes_generados/INFORME_POSGRADOS_2026.docx`

### 📥 Descargar Último Informe

- [Ver informes generados](./informes_generados/)
- [Ver historial de ejecuciones](../../actions/workflows/generar-informe.yml)
