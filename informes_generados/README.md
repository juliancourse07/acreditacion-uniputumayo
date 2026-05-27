# 📊 Informes Generados Automáticamente

Esta carpeta contiene los informes generados automáticamente por GitHub Actions.

## 📥 Cómo Descargar el Último Informe

### Método 1: Desde GitHub Actions (Recomendado)

1. Ve a la pestaña **Actions** del repositorio
2. Haz clic en el workflow **"Generar Informe de Posgrados"**
3. Selecciona la ejecución más reciente (con ✅)
4. En la sección **Artifacts**, descarga: `informe-posgrados-XXX`
5. Descomprime el ZIP y abre el archivo Word

### Método 2: Desde el repositorio

Si el informe fue commiteado automáticamente:

1. Navega a esta carpeta en GitHub
2. Haz clic en `INFORME_POSGRADOS_2026.docx`
3. Haz clic en **Download**

## 🚀 Cómo Generar un Nuevo Informe

### Opción A: Ejecución Manual

1. Ve a **Actions** > **Generar Informe de Posgrados**
2. Haz clic en **Run workflow**
3. Selecciona la rama `main`
4. Haz clic en **Run workflow**
5. Espera 2-3 minutos
6. Descarga el artefacto generado

### Opción B: Automático

El informe se regenera automáticamente cuando:
- Se modifica el archivo Excel de encuestas
- Se actualiza el script de generación
- Se hace push a la rama `main`

## 📋 Contenido del Informe

El informe incluye:

- ✅ **Portada** con logos institucionales
- ✅ **Tabla de contenido**
- ✅ **Introducción** (contexto, objetivos, metodología)
- ✅ **Marco teórico** con citas académicas
- ✅ **Caracterización de encuestados** (701 egresados)
- ✅ **Análisis de demanda de posgrados** con gráficos
- ✅ **Análisis cruzado** (programa origen vs posgrado solicitado)
- ✅ **Análisis geográfico**
- ✅ **Respuestas abiertas** (campo "Otro")
- ✅ **Discusión** académica
- ✅ **Conclusiones** basadas en datos
- ✅ **Recomendaciones** por horizonte temporal
- ✅ **Referencias bibliográficas** (formato APA)
- ✅ **Anexos**

## 🔄 Historial de Generación

Cada ejecución de GitHub Actions crea un artefacto numerado que se conserva por 90 días.

Para ver el historial completo: **Actions** > **Generar Informe de Posgrados**

## ⚙️ Configuración

- **Script principal**: `scripts/generar_informe_posgrados.py`
- **Ejecutor**: `ejecutar_generador.py`
- **Configuración**: `scripts/config.yaml`
- **Excel de entrada**: `ENCUESTA SEGUIMIENTO EGRESADOS 30-04-2026.xlsx`
- **Plantilla Word**: `INTERES EN APERTURA DE PROGRAMAS ACADEMICOS UNIPUTUMAYO.docx`

---

**Última actualización**: Generado automáticamente por GitHub Actions

