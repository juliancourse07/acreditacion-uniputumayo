# Generador Automatizado de Informes de Posgrados

## Instalación

```bash
pip install -r scripts/requirements.txt
```

## Uso

### Opción 1: Línea de comandos

```bash
python scripts/generar_informe_posgrados.py \
  --excel "ruta/al/excel.xlsx" \
  --plantilla "ruta/a/plantilla.docx" \
  --output "INFORME_FINAL.docx"
```

### Opción 2: Configuración por defecto

```bash
python scripts/generar_informe_posgrados.py
```

Toma valores desde `scripts/config.yaml`. También intenta detectar archivos con nombres similares si no se pasan rutas explícitas.

## Estructura

- `scripts/generar_informe_posgrados.py` - Script principal
- `scripts/requirements.txt` - Dependencias Python
- `scripts/config.yaml` - Configuración base
- `README_GENERADOR_INFORMES.md` - Esta guía

## Salidas

- `informes_generados/INFORME_POSGRADOS_YYYY-MM-DD.docx` - Informe final
- `temp/grafico_posgrados.png` - Gráfico temporal
- `logs/ejecucion_YYYYMMDD.log` - Log de ejecución

## Secciones generadas automáticamente

1. Portada de análisis (respetando plantilla institucional)
2. Resumen ejecutivo
3. Caracterización de encuestados
4. Frecuencia de posgrados y gráfico Top 15
5. Análisis por programa de origen
6. Respuestas abiertas del campo "Otro"
7. Análisis geográfico
8. Conclusiones y recomendaciones

## Manejo de errores incluido

- Validación de existencia de archivos de entrada
- Validación de columnas requeridas (lógica por alias)
- Relleno de valores faltantes con `No especificado`
- Logging de advertencias y trazabilidad del proceso
