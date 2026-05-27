#!/usr/bin/env python3
"""Generador automatizado de informe de interés en programas de posgrado."""

from __future__ import annotations

import argparse
import logging
import os
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
import yaml
from docx import Document
from docx.shared import Inches


COL_KEYS = {
    "cedula": ["Número de cédula", "Numero de cedula"],
    "genero": ["Género", "Genero"],
    "municipio": ["Municipio de residencia Actual", "Municipio de residencia"],
    "sede": [
        "Indique Sede, subsede o Ampliación de egreso de formación",
        "Sede",
    ],
    "titulo": ["Título obtenido en la UNIPUTUMAYO", "Titulo obtenido en la UNIPUTUMAYO"],
    "ciclo": ["Ciclo Profesional", "Ciclo"],
    "programa_origen": [
        "Seleccione el programa del que es egresado",
        "Programa del que es egresado",
    ],
    "posgrado": [
        "¿Qué posgrados considera que debe ofrecer la Uniputumayo, en concordancia a su programa de formación?",
    ],
    "posgrado_otro": [
        "¿Qué posgrados considera que debe ofrecer la Uniputumayo, en concordancia a su programa de formación? [Otro]",
    ],
}

CRITICAL_KEYS = ["genero", "municipio", "sede", "ciclo", "programa_origen", "posgrado"]


@dataclass
class Configuracion:
    excel_file: str | None = None
    plantilla_word: str | None = None
    output_folder: str = "informes_generados"
    output_name: str = "INFORME_POSGRADOS_{fecha}.docx"
    top_n_posgrados: int = 15


def limpiar_texto(valor: Any) -> str:
    if pd.isna(valor):
        return ""
    texto = str(valor)
    texto = texto.replace("\xa0", " ").replace("&nbsp;", " ")
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def normalizar_columna(nombre: str) -> str:
    nombre = limpiar_texto(nombre)
    nombre = re.sub(r"\s+", " ", nombre)
    return nombre


def resolver_columna(df: pd.DataFrame, opciones: list[str]) -> str | None:
    columnas = {normalizar_columna(c): c for c in df.columns}
    for opcion in opciones:
        n_opcion = normalizar_columna(opcion)
        if n_opcion in columnas:
            return columnas[n_opcion]

    for opcion in opciones:
        base = re.sub(r"\s+", " ", opcion).strip().lower()
        for columna in df.columns:
            c_norm = re.sub(r"\s+", " ", limpiar_texto(columna)).lower()
            if base in c_norm:
                return columna
    return None


def detectar_archivo_por_patron(patrones: list[str]) -> str | None:
    candidatos = []
    for patron in patrones:
        candidatos.extend(Path(".").glob(patron))
    if not candidatos:
        return None
    candidatos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return str(candidatos[0])


def cargar_configuracion(config_path: str | None) -> Configuracion:
    cfg = Configuracion()
    if config_path and os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
        encuesta = data.get("encuesta", {})
        graficos = data.get("graficos", {})
        output = data.get("output", {})
        cfg.excel_file = encuesta.get("excel_file")
        cfg.plantilla_word = encuesta.get("plantilla_word")
        cfg.top_n_posgrados = int(graficos.get("top_n_posgrados", cfg.top_n_posgrados))
        cfg.output_folder = output.get("carpeta", cfg.output_folder)
        cfg.output_name = output.get("nombre_archivo", cfg.output_name)
    return cfg


def cargar_datos(excel_path: str) -> pd.DataFrame:
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"No se encontró el archivo: {excel_path}")

    logging.info("Leyendo archivo Excel: %s", excel_path)
    df = pd.read_excel(excel_path)
    df.columns = [normalizar_columna(c) for c in df.columns]

    mapping = {}
    for key, aliases in COL_KEYS.items():
        col = resolver_columna(df, aliases)
        if col:
            mapping[key] = col

    required = ["genero", "municipio", "posgrado"]
    faltantes = [k for k in required if k not in mapping]
    if faltantes:
        faltantes_txt = ", ".join(faltantes)
        raise ValueError(f"No se encontraron columnas requeridas (lógicas): {faltantes_txt}")

    df = df.dropna(how="all").copy()

    for key in CRITICAL_KEYS:
        if key in mapping:
            col = mapping[key]
            vacios = df[col].isna().sum()
            if vacios:
                logging.warning("Columna %s tiene %s vacíos, se rellena con 'No especificado'", col, vacios)
            df[col] = df[col].apply(limpiar_texto)
            df[col] = df[col].replace("", "No especificado")

    df.attrs["column_mapping"] = mapping
    logging.info("Total de registros procesados: %s", len(df))
    return df


def _split_posgrados(valor: str) -> list[str]:
    if not valor:
        return []
    partes = re.split(r"\s*(?:\||;|,|/|\\n|\\r|\\t)\s*", valor)
    return [limpiar_texto(p) for p in partes if limpiar_texto(p)]


def extraer_posgrados(df: pd.DataFrame) -> list[str]:
    mapping = df.attrs.get("column_mapping", {})
    col_pos = mapping.get("posgrado")
    col_otro = mapping.get("posgrado_otro")
    if not col_pos:
        return []

    salida: list[str] = []
    for _, row in df.iterrows():
        principal = limpiar_texto(row.get(col_pos, ""))
        otro = limpiar_texto(row.get(col_otro, "")) if col_otro else ""

        if principal.lower() == "otro" and otro:
            salida.extend(_split_posgrados(otro))
        elif principal and principal != "No especificado":
            salida.extend(_split_posgrados(principal))
            if otro and otro != "No especificado":
                salida.extend(_split_posgrados(otro))
    return salida


def analizar_demografia(df: pd.DataFrame) -> dict[str, Any]:
    mapping = df.attrs["column_mapping"]
    total = len(df)

    genero_col = mapping.get("genero")
    municipio_col = mapping.get("municipio")
    sede_col = mapping.get("sede")
    ciclo_col = mapping.get("ciclo")
    programa_col = mapping.get("programa_origen")

    genero_counts = df[genero_col].value_counts(dropna=False) if genero_col else pd.Series(dtype=int)

    hombres = int(sum(v for k, v in genero_counts.items() if "hombre" in limpiar_texto(k).lower()))
    mujeres = int(sum(v for k, v in genero_counts.items() if "mujer" in limpiar_texto(k).lower()))

    def porcentaje(n: int) -> float:
        return (n / total * 100) if total else 0.0

    return {
        "total_encuestados": total,
        "hombres": hombres,
        "mujeres": mujeres,
        "hombres_pct": porcentaje(hombres),
        "mujeres_pct": porcentaje(mujeres),
        "municipios_top": (df[municipio_col].value_counts().head(5).to_dict() if municipio_col else {}),
        "sedes": (df[sede_col].value_counts().to_dict() if sede_col else {}),
        "ciclos": (df[ciclo_col].value_counts().to_dict() if ciclo_col else {}),
        "programas_origen_top": (df[programa_col].value_counts().head(5).to_dict() if programa_col else {}),
    }


def analizar_posgrados(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str], dict[str, int]]:
    posgrados = extraer_posgrados(df)
    total = len(posgrados)

    if total == 0:
        tabla = pd.DataFrame(columns=["Programa de Posgrado", "Frecuencia", "Porcentaje"])
        return tabla, [], {}

    conteo = Counter(posgrados)
    tabla = pd.DataFrame(
        [{"Programa de Posgrado": k, "Frecuencia": v} for k, v in conteo.items()]
    ).sort_values(["Frecuencia", "Programa de Posgrado"], ascending=[False, True])
    tabla["Porcentaje"] = (tabla["Frecuencia"] / total * 100).round(2)
    tabla = tabla.reset_index(drop=True)

    otro_col = df.attrs.get("column_mapping", {}).get("posgrado_otro")
    otros = []
    if otro_col:
        for val in df[otro_col].tolist():
            t = limpiar_texto(val)
            if t and t != "No especificado":
                otros.extend(_split_posgrados(t))

    return tabla, otros, dict(conteo)


def crear_graficos(df_posgrados: pd.DataFrame, output_dir: str, top_n: int = 15) -> str | None:
    if df_posgrados.empty:
        return None

    os.makedirs(output_dir, exist_ok=True)
    top = df_posgrados.head(top_n).copy().iloc[::-1]

    sns.set_theme(style="whitegrid")
    plt.figure(figsize=(12, 8))
    ax = sns.barplot(data=top, x="Frecuencia", y="Programa de Posgrado", color="#1f4e79")
    ax.set_title(f"Top {min(top_n, len(df_posgrados))} posgrados más solicitados")
    ax.set_xlabel("Frecuencia")
    ax.set_ylabel("Programa de Posgrado")

    for i, (_, row) in enumerate(top.iterrows()):
        ax.text(row["Frecuencia"] + 0.05, i, str(int(row["Frecuencia"])), va="center")

    plt.tight_layout()
    ruta = os.path.join(output_dir, "grafico_posgrados.png")
    plt.savefig(ruta, dpi=200)
    plt.close()
    return ruta


def analizar_por_programa_origen(df: pd.DataFrame) -> pd.DataFrame:
    mapping = df.attrs["column_mapping"]
    programa_col = mapping.get("programa_origen")
    pos_col = mapping.get("posgrado")
    otro_col = mapping.get("posgrado_otro")

    if not programa_col or not pos_col:
        return pd.DataFrame(columns=["Programa de Egreso", "Posgrados más solicitados", "Frecuencia"])

    acumulado: dict[str, Counter[str]] = defaultdict(Counter)

    for _, row in df.iterrows():
        programa = limpiar_texto(row.get(programa_col, "")) or "No especificado"
        principal = limpiar_texto(row.get(pos_col, ""))
        otro = limpiar_texto(row.get(otro_col, "")) if otro_col else ""

        if principal.lower() == "otro" and otro:
            for p in _split_posgrados(otro):
                acumulado[programa][p] += 1
        else:
            for p in _split_posgrados(principal):
                if p and p != "No especificado":
                    acumulado[programa][p] += 1

    filas = []
    for programa, counter in sorted(acumulado.items(), key=lambda x: (-sum(x[1].values()), x[0])):
        top = counter.most_common(2)
        if not top:
            continue
        pos_txt = "\n".join(f"{idx}. {nombre}" for idx, (nombre, _) in enumerate(top, start=1))
        freq_txt = "\n".join(str(freq) for _, freq in top)
        filas.append(
            {
                "Programa de Egreso": programa,
                "Posgrados más solicitados": pos_txt,
                "Frecuencia": freq_txt,
            }
        )

    return pd.DataFrame(filas)


def analizar_geografico(df: pd.DataFrame) -> pd.DataFrame:
    mapping = df.attrs["column_mapping"]
    municipio_col = mapping.get("municipio")
    pos_col = mapping.get("posgrado")
    otro_col = mapping.get("posgrado_otro")

    if not municipio_col or not pos_col:
        return pd.DataFrame(columns=["Municipio", "Total Encuestados", "Posgrados más solicitados"])

    municipio_data: dict[str, Counter[str]] = defaultdict(Counter)
    municipio_totales = Counter()

    for _, row in df.iterrows():
        municipio = limpiar_texto(row.get(municipio_col, "")) or "No especificado"
        municipio_totales[municipio] += 1

        principal = limpiar_texto(row.get(pos_col, ""))
        otro = limpiar_texto(row.get(otro_col, "")) if otro_col else ""

        if principal.lower() == "otro" and otro:
            for p in _split_posgrados(otro):
                municipio_data[municipio][p] += 1
        else:
            for p in _split_posgrados(principal):
                if p and p != "No especificado":
                    municipio_data[municipio][p] += 1

    filas = []
    for municipio, total in municipio_totales.most_common():
        top = municipio_data[municipio].most_common(3)
        top_txt = ", ".join(f"{nombre} ({freq})" for nombre, freq in top) if top else "No especificado"
        filas.append(
            {
                "Municipio": municipio,
                "Total Encuestados": total,
                "Posgrados más solicitados": top_txt,
            }
        )

    return pd.DataFrame(filas)


def analizar_otro_temas(otros: list[str]) -> dict[str, int]:
    temas = {
        "Ambiental": ["ambient", "sosten", "recursos naturales"],
        "Gerencia y negocios": ["gerencia", "administr", "mercadeo", "marketing", "finanzas"],
        "Infraestructura": ["pavimento", "vías", "civil", "constru"],
        "Educación": ["docencia", "pedagog", "educ"],
        "Tecnología": ["datos", "software", "sistemas", "ia", "tecnolog"],
    }
    conteo = Counter()

    for respuesta in otros:
        r = limpiar_texto(respuesta).lower()
        if not r:
            continue
        etiquetado = False
        for tema, patrones in temas.items():
            if any(p in r for p in patrones):
                conteo[tema] += 1
                etiquetado = True
                break
        if not etiquetado:
            conteo["Otros"] += 1

    return dict(conteo)


def generar_conclusiones(stats: dict[str, Any]) -> dict[str, list[str]]:
    total_solicitudes = stats.get("total_solicitudes_posgrado", 0) or 1
    top_pos = stats.get("top_posgrados", [])
    geo = stats.get("geo_top", [])

    area_dominante = "interdisciplinar"
    if top_pos:
        top_nombre = top_pos[0]["programa"].lower()
        if "ambient" in top_nombre:
            area_dominante = "ambiental"
        elif any(k in top_nombre for k in ["gerencia", "mercadeo", "administr", "negocio"]):
            area_dominante = "gerencia y negocios"

    porcentaje_top = (top_pos[0]["frecuencia"] / total_solicitudes * 100) if top_pos else 0

    conclusiones = [
        f"El {porcentaje_top:.1f}% de las solicitudes se concentra en el área de {area_dominante}.",
        "Los programas más demandados son: "
        + "; ".join(
            f"{item['programa']} ({item['frecuencia']} solicitudes)" for item in top_pos[:2]
        )
        if top_pos
        else "No se identificaron programas con frecuencia suficiente.",
        "La demanda regional se concentra en: " + ", ".join(geo[:3]) if geo else "No hay concentración geográfica clara.",
    ]

    recomendaciones = [
        "Priorizar la apertura de: "
        + ", ".join(item["programa"] for item in top_pos[:3])
        if top_pos
        else "Priorizar análisis adicional de demanda.",
        "Considerar alianzas estratégicas para programas con menor demanda relativa.",
        "Ampliar oferta y socialización en municipios con mayor participación.",
        "Realizar estudios de factibilidad para propuestas del campo 'Otro'.",
    ]

    return {"conclusiones": conclusiones, "recomendaciones": recomendaciones}


def _add_table(doc: Document, headers: list[str], rows: list[list[str]]) -> None:
    tabla = doc.add_table(rows=1, cols=len(headers))
    tabla.style = "Table Grid"
    for i, h in enumerate(headers):
        tabla.rows[0].cells[i].text = h
    for row in rows:
        celdas = tabla.add_row().cells
        for i, val in enumerate(row):
            celdas[i].text = limpiar_texto(val)


def _add_section_heading(doc: Document, texto: str) -> None:
    """Agrega un encabezado de sección usando párrafo en negrita (compatible con todas las plantillas)."""
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(texto)
    run.bold = True
    run.font.size = Pt(13)


def _add_numbered_paragraph(doc: Document, numero: int, texto: str) -> None:
    """Agrega un párrafo numerado manualmente."""
    p = doc.add_paragraph()
    p.add_run(f"{numero}. {texto}")


def rellenar_word(plantilla_path: str, datos: dict[str, Any], output_path: str) -> None:
    if not os.path.exists(plantilla_path):
        raise FileNotFoundError(f"No se encontró la plantilla Word: {plantilla_path}")

    doc = Document(plantilla_path)

    doc.add_page_break()
    _add_section_heading(doc, "ANÁLISIS DE INTERÉS EN APERTURA DE PROGRAMAS DE POSGRADO")
    doc.add_paragraph("Encuesta a Egresados 2026")
    doc.add_paragraph(f"Fecha de generación: {datetime.now():%Y-%m-%d %H:%M}")

    _add_section_heading(doc, "RESUMEN EJECUTIVO")
    resumen = datos["resumen"]
    p = doc.add_paragraph()
    p.add_run(f"Total de egresados encuestados: {resumen['total_encuestados']}\n")
    p.add_run(f"Programas de posgrado identificados: {resumen['total_posgrados']}\n")
    p.add_run("Top 3 posgrados más solicitados:\n")
    for i, item in enumerate(resumen["top_3"], start=1):
        p.add_run(f"  {i}. {item['programa']} - {item['frecuencia']} solicitudes ({item['porcentaje']:.2f}%)\n")

    _add_section_heading(doc, "CARACTERIZACIÓN DE ENCUESTADOS")
    car = datos["caracterizacion"]
    _add_table(
        doc,
        ["Indicador", "Resultado"],
        [
            ["Total de encuestados", str(car["total_encuestados"])],
            [
                "Distribución por género",
                f"Hombres: {car['hombres']} ({car['hombres_pct']:.2f}%) / Mujeres: {car['mujeres']} ({car['mujeres_pct']:.2f}%)",
            ],
            [
                "Municipios principales de residencia",
                ", ".join(f"{k} ({v})" for k, v in car["municipios_top"].items()) or "No especificado",
            ],
            [
                "Sedes de egreso",
                ", ".join(f"{k} ({v})" for k, v in car["sedes"].items()) or "No especificado",
            ],
            [
                "Ciclo académico",
                ", ".join(f"{k}: {v}" for k, v in car["ciclos"].items()) or "No especificado",
            ],
            [
                "Programas de origen más representados",
                ", ".join(f"{k} ({v})" for k, v in car["programas_origen_top"].items()) or "No especificado",
            ],
        ],
    )

    _add_section_heading(doc, "ANÁLISIS DE INTERÉS EN POSGRADOS")
    pos_rows = [
        [
            str(row["Programa de Posgrado"]),
            str(int(row["Frecuencia"])),
            f"{float(row['Porcentaje']):.2f}%",
        ]
        for _, row in datos["tabla_posgrados"].iterrows()
    ]
    _add_table(doc, ["Programa de Posgrado", "Frecuencia", "Porcentaje"], pos_rows)

    if datos.get("grafico_path") and os.path.exists(datos["grafico_path"]):
        doc.add_paragraph()
        doc.add_picture(datos["grafico_path"], width=Inches(6.5))

    _add_section_heading(doc, "ANÁLISIS POR PROGRAMA DE ORIGEN")
    prog_rows = [
        [
            str(row["Programa de Egreso"]),
            str(row["Posgrados más solicitados"]),
            str(row["Frecuencia"]),
        ]
        for _, row in datos["tabla_programa_origen"].iterrows()
    ]
    _add_table(doc, ["Programa de Egreso", "Posgrados más solicitados", "Frecuencia"], prog_rows)

    _add_section_heading(doc, "RESPUESTAS ABIERTAS (CAMPO OTRO)")
    if datos["otros"]:
        available_styles = {s.name for s in doc.styles}
        bullet_style = "List Bullet" if "List Bullet" in available_styles else "Normal"
        for item in datos["otros"]:
            doc.add_paragraph(f"- {item}", style=bullet_style)
    else:
        doc.add_paragraph("No se registraron respuestas en el campo 'Otro'.")

    doc.add_paragraph(
        "Áreas temáticas emergentes: "
        + (
            ", ".join(f"{k} ({v})" for k, v in datos["temas_otros"].items())
            if datos["temas_otros"]
            else "No se identificaron patrones temáticos."
        )
    )

    _add_section_heading(doc, "ANÁLISIS GEOGRÁFICO")
    geo_rows = [
        [
            str(row["Municipio"]),
            str(row["Total Encuestados"]),
            str(row["Posgrados más solicitados"]),
        ]
        for _, row in datos["tabla_geografica"].iterrows()
    ]
    _add_table(doc, ["Municipio", "Total Encuestados", "Posgrados más solicitados"], geo_rows)

    _add_section_heading(doc, "CONCLUSIONES Y RECOMENDACIONES")
    doc.add_paragraph("Conclusiones:")
    for i, c in enumerate(datos["conclusiones"]["conclusiones"], start=1):
        _add_numbered_paragraph(doc, i, c)

    doc.add_paragraph("Recomendaciones:")
    for i, r in enumerate(datos["conclusiones"]["recomendaciones"], start=1):
        _add_numbered_paragraph(doc, i, r)

    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    doc.save(output_path)


def configurar_logging() -> None:
    os.makedirs("logs", exist_ok=True)
    logging.basicConfig(
        filename=f"logs/ejecucion_{datetime.now():%Y%m%d}.log",
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )


def construir_argumentos() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generar informe automatizado de posgrados")
    parser.add_argument("--excel", type=str, default=None, help="Ruta al archivo Excel de encuesta")
    parser.add_argument("--plantilla", type=str, default=None, help="Ruta al archivo Word plantilla")
    parser.add_argument("--output", type=str, default=None, help="Ruta del informe final .docx")
    parser.add_argument("--config", type=str, default="scripts/config.yaml", help="Ruta al archivo config.yaml")
    return parser.parse_args()


def main() -> None:
    args = construir_argumentos()
    configurar_logging()
    logging.info("Iniciando generación de informe...")

    cfg = cargar_configuracion(args.config)

    excel_path = args.excel or cfg.excel_file or detectar_archivo_por_patron(["*ENCUESTA*EGRESADOS*.xlsx", "*encuesta*egresados*.xlsx"])
    plantilla_path = args.plantilla or cfg.plantilla_word or detectar_archivo_por_patron(["*INTERES*APERTURA*PROGRAMAS*ACADEMICOS*.docx", "*interes*apertura*programas*academicos*.docx"])

    if not excel_path:
        raise FileNotFoundError("No se pudo resolver la ruta del Excel. Use --excel o config.yaml")
    if not plantilla_path:
        raise FileNotFoundError("No se pudo resolver la ruta de la plantilla Word. Use --plantilla o config.yaml")

    fecha = datetime.now().strftime("%Y-%m-%d")
    output_default = os.path.join(cfg.output_folder, cfg.output_name.format(fecha=fecha))
    output_path = args.output or output_default

    df = cargar_datos(excel_path)
    demografia = analizar_demografia(df)
    tabla_posgrados, otros, conteo_posgrados = analizar_posgrados(df)
    tabla_programa_origen = analizar_por_programa_origen(df)
    tabla_geografica = analizar_geografico(df)

    temp_dir = "temp"
    grafico_path = crear_graficos(tabla_posgrados, temp_dir, top_n=cfg.top_n_posgrados)

    total_solicitudes = int(tabla_posgrados["Frecuencia"].sum()) if not tabla_posgrados.empty else 0
    top_3 = []
    for _, row in tabla_posgrados.head(3).iterrows():
        top_3.append(
            {
                "programa": str(row["Programa de Posgrado"]),
                "frecuencia": int(row["Frecuencia"]),
                "porcentaje": float(row["Porcentaje"]),
            }
        )

    stats = {
        "total_solicitudes_posgrado": total_solicitudes,
        "top_posgrados": top_3,
        "geo_top": list(tabla_geografica["Municipio"].head(3)) if not tabla_geografica.empty else [],
    }

    datos = {
        "resumen": {
            "total_encuestados": demografia["total_encuestados"],
            "total_posgrados": len(conteo_posgrados),
            "top_3": top_3,
        },
        "caracterizacion": demografia,
        "tabla_posgrados": tabla_posgrados,
        "grafico_path": grafico_path,
        "tabla_programa_origen": tabla_programa_origen,
        "otros": sorted(set(otros)),
        "temas_otros": analizar_otro_temas(otros),
        "tabla_geografica": tabla_geografica,
        "conclusiones": generar_conclusiones(stats),
    }

    rellenar_word(plantilla_path, datos, output_path)
    logging.info("Informe generado correctamente en: %s", output_path)
    print(f"Informe generado correctamente: {output_path}")


if __name__ == "__main__":
    main()
