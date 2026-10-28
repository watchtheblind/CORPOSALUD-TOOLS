"""Lógica de consolidación de Excel, sin interfaz gráfica.

Regla principal: el encabezado del consolidado sale de la **fila 1** del
archivo de referencia. La fila 1 de los demás archivos se descarta; esos
archivos solo aportan datos.

Este módulo es la versión reutilizable de la maqueta del notebook
`notebooks/01_maqueta_combinador.ipynb`.
"""

from __future__ import annotations

import re
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path
from typing import TypedDict

import pandas as pd

EXTENSIONES = {".xlsx", ".xls", ".xlsm"}
RESERVADAS = {"archivo_origen", "hoja_origen", "fila_origen"}
UMBRAL_SUGERENCIA = 0.85


class PerfilArchivo(TypedDict):
    """Resultado de leer un Excel y separar encabezado de datos."""

    ruta: Path
    nombre: str
    hoja: str | int
    columnas: list[str]
    datos: pd.DataFrame


# ---------------------------------------------------------------------------
# Lectura de archivos
# ---------------------------------------------------------------------------


def listar_excels(carpeta: Path) -> list[Path]:
    """Devuelve los Excel de una carpeta, ignorando temporales de Excel."""
    carpeta = Path(carpeta)
    if not carpeta.exists():
        return []
    return sorted(
        p
        for p in carpeta.iterdir()
        if p.is_file()
        and p.suffix.lower() in EXTENSIONES
        and not p.name.startswith("~$")
    )


def hojas_del_libro(ruta: Path) -> list[str]:
    """Lista las hojas de un libro de Excel."""
    with pd.ExcelFile(ruta) as libro:
        return libro.sheet_names


def escanear_hojas(archivos: list[Path]) -> dict[str, list[str]]:
    """Devuelve ``{nombre_archivo: [hojas_disponibles]}`` para cada archivo."""
    return {Path(ruta).name: hojas_del_libro(Path(ruta)) for ruta in archivos}


def leer_crudo(ruta: Path, hoja: str | int = 0) -> pd.DataFrame:
    """Lee sin interpretar encabezados: la fila 0 es la fila 1 de Excel.

    Se lee todo como texto para preservar cédulas y códigos con ceros a la
    izquierda. Las celdas vacías quedan como cadena vacía.
    """
    return pd.read_excel(
        ruta,
        sheet_name=hoja,
        header=None,
        dtype=str,
        keep_default_na=False,
    )


# ---------------------------------------------------------------------------
# Títulos y perfiles
# ---------------------------------------------------------------------------


def limpiar_titulo(valor: object) -> str:
    """Normaliza un título para usarlo como nombre de columna."""
    if pd.isna(valor):
        return ""
    texto = str(valor).replace("\n", " ")
    return re.sub(r"\s+", " ", texto).strip()


def recortar_titulos(titulos: list[str]) -> list[str]:
    """Quita las columnas vacías del final."""
    while titulos and titulos[-1] == "":
        titulos.pop()
    return titulos


def validar_titulos(titulos: list[str], etiqueta: str) -> None:
    """Verifica que no haya títulos vacíos ni duplicados."""
    vacios = [i + 1 for i, titulo in enumerate(titulos) if titulo == ""]
    if vacios:
        raise ValueError(
            f"{etiqueta}: hay títulos vacíos en las posiciones {vacios}. "
            "Completalos en el Excel antes de combinar."
        )
    repetidos = sorted({t for t in titulos if titulos.count(t) > 1})
    if repetidos:
        raise ValueError(
            f"{etiqueta}: títulos duplicados {repetidos}. "
            "Cada columna necesita un nombre único."
        )


def buscar_archivo(archivos: list[Path], nombre: str) -> Path:
    """Busca un archivo por nombre dentro de la lista de archivos cargados."""
    coincidencias = [p for p in archivos if p.name == nombre]
    if not coincidencias:
        disponibles = ", ".join(p.name for p in archivos)
        raise FileNotFoundError(f"No encontré {nombre!r}. Disponibles: {disponibles}")
    return coincidencias[0]


def es_vacio(valor: object) -> bool:
    """Indica si una celda está vacía o solo tiene espacios."""
    return pd.isna(valor) or str(valor).strip() == ""


def cargar_perfil(ruta: Path, hoja: str | int = 0) -> PerfilArchivo:
    """Lee un Excel y separa la fila 1 (encabezado) de los datos."""
    if isinstance(hoja, int):
        hojas = hojas_del_libro(ruta)
        try:
            nombre_hoja = hojas[hoja]
        except IndexError as exc:
            raise ValueError(
                f"{ruta.name}: no existe la hoja en la posición {hoja}. "
                f"Hojas disponibles: {hojas}"
            ) from exc
    else:
        nombre_hoja = hoja

    crudo = leer_crudo(ruta, hoja)
    if crudo.empty:
        raise ValueError(f"{ruta.name}: la hoja {nombre_hoja!r} está vacía.")

    columnas = recortar_titulos(
        [limpiar_titulo(valor) for valor in crudo.iloc[0].tolist()]
    )
    validar_titulos(columnas, f"{ruta.name} / hoja {nombre_hoja!r}")

    datos = crudo.iloc[1:, : len(columnas)].copy()
    datos.columns = columnas

    # Quitamos filas totalmente vacías, conservando el índice original
    # (índice 1 = fila 2 de Excel) para poder reportar la fila real.
    fila_vacia = datos.apply(
        lambda fila: all(es_vacio(valor) for valor in fila), axis=1
    )
    datos = datos.loc[~fila_vacia]

    return {
        "ruta": Path(ruta),
        "nombre": Path(ruta).name,
        "hoja": nombre_hoja,
        "columnas": columnas,
        "datos": datos,
    }


# ---------------------------------------------------------------------------
# Mapeo de columnas
# ---------------------------------------------------------------------------


def normalizar_titulo(texto: str) -> str:
    """Mayúsculas, sin acentos y sin puntuación, para comparar títulos."""
    sin_acentos = "".join(
        c
        for c in unicodedata.normalize("NFKD", str(texto))
        if not unicodedata.combining(c)
    )
    limpio = re.sub(r"[^A-Za-z0-9]+", " ", sin_acentos.upper())
    return re.sub(r"\s+", " ", limpio).strip()


def sugerir_mapeo(
    columnas_ref: list[str],
    columnas_fuente: list[str],
    umbral: float = UMBRAL_SUGERENCIA,
) -> dict[str, str | None]:
    """Sugiere una columna fuente para cada columna de referencia.

    La sugerencia es uno-a-uno: una columna fuente no se reutiliza para dos
    campos. Si la similitud no llega al umbral, devuelve ``None``.
    """
    usadas: set[str] = set()
    sugerencias: dict[str, str | None] = {}
    for col_ref in columnas_ref:
        candidatas = [c for c in columnas_fuente if c not in usadas]
        if not candidatas:
            sugerencias[col_ref] = None
            continue
        normal_ref = normalizar_titulo(col_ref)
        puntajes = [
            (c, SequenceMatcher(None, normal_ref, normalizar_titulo(c)).ratio())
            for c in candidatas
        ]
        mejor, puntaje = max(puntajes, key=lambda par: par[1])
        if puntaje >= umbral:
            sugerencias[col_ref] = mejor
            usadas.add(mejor)
        else:
            sugerencias[col_ref] = None
    return sugerencias


def construir_mapeo(
    perfiles: dict[str, PerfilArchivo],
    columnas_ref: list[str],
    ajustes: dict[str, dict[str, str | None]] | None = None,
    usar_sugerencias: bool = True,
) -> dict[str, dict[str, str | None]]:
    """Construye el mapeo referencia -> columna fuente para cada archivo.

    ``ajustes`` pisa las sugerencias. Un valor ``None`` deja la columna vacía.
    """
    ajustes = ajustes or {}
    mapeo: dict[str, dict[str, str | None]] = {}
    for nombre, perfil in perfiles.items():
        if usar_sugerencias:
            mapa = sugerir_mapeo(columnas_ref, perfil["columnas"])
        else:
            mapa = {col: None for col in columnas_ref}

        for col_ref, col_fuente in ajustes.get(nombre, {}).items():
            if col_ref not in columnas_ref:
                raise KeyError(
                    f"{nombre}: {col_ref!r} no existe en el encabezado de referencia."
                )
            mapa[col_ref] = col_fuente
        mapeo[nombre] = mapa
    return mapeo


def detectar_extras(
    perfiles: dict[str, PerfilArchivo],
    mapeo: dict[str, dict[str, str | None]],
    columnas_ref: list[str],
    archivo_referencia: str | None = None,
) -> tuple[dict[str, list[str]], dict[str, dict[str, list[str]]]]:
    """Separa columnas usadas para mapear de las extras reales.

    Devuelve ``(extras_reales, detalle)``. Una extra real es una columna que
    no está en la referencia y tampoco se usa como fuente de un campo mapeado.
    """
    extras_reales: dict[str, list[str]] = {}
    detalle: dict[str, dict[str, list[str]]] = {}
    for nombre, perfil in perfiles.items():
        if nombre == archivo_referencia:
            continue
        mapeadas = {v for v in mapeo[nombre].values() if v}
        fuera = [c for c in perfil["columnas"] if c not in columnas_ref]
        usadas = [c for c in fuera if c in mapeadas]
        extras = [c for c in fuera if c not in mapeadas]
        detalle[nombre] = {"usadas": usadas, "extras": extras}
        if extras:
            extras_reales[nombre] = extras
    return extras_reales, detalle


def columnas_descartadas(
    perfiles: dict[str, PerfilArchivo],
    mapeo: dict[str, dict[str, str | None]],
    extras: dict[str, list[str]] | None = None,
) -> dict[str, list[str]]:
    """Columnas que no se usan ni como fuente de mapeo ni como extra."""
    extras = extras or {}
    resultado: dict[str, list[str]] = {}
    for nombre, perfil in perfiles.items():
        usadas = {v for v in mapeo[nombre].values() if v}
        incluidas = set(extras.get(nombre, []))
        descartadas = [
            c for c in perfil["columnas"] if c not in usadas and c not in incluidas
        ]
        if descartadas:
            resultado[nombre] = descartadas
    return resultado


# ---------------------------------------------------------------------------
# Validación y consolidación
# ---------------------------------------------------------------------------


def validar_configuracion(
    perfiles: dict[str, PerfilArchivo],
    mapeo: dict[str, dict[str, str | None]],
    extras: dict[str, list[str]] | None,
    columnas_ref: list[str],
    exigir_mapeo_completo: bool = False,
) -> tuple[list[str], list[str]]:
    """Valida el mapeo y las extras.

    Devuelve ``(problemas, avisos)``. Los problemas impiden combinar; los
    avisos son columnas que quedarán vacías.
    """
    extras = extras or {}
    problemas: list[str] = []
    avisos: list[str] = []

    for nombre in extras:
        if nombre not in perfiles:
            problemas.append(
                f"{nombre}: hay configuración de extras para un archivo no cargado."
            )

    for nombre, mapa in mapeo.items():
        perfil = perfiles[nombre]
        usadas: list[str] = []

        for col_ref, col_fuente in mapa.items():
            if col_ref not in columnas_ref:
                problemas.append(
                    f"{nombre}: {col_ref!r} no está en el encabezado de referencia."
                )
            if col_fuente is None:
                continue
            if col_fuente not in perfil["columnas"]:
                problemas.append(
                    f"{nombre}: {col_ref!r} apunta a {col_fuente!r}, "
                    "que no existe en ese archivo."
                )
                continue
            usadas.append(col_fuente)

        duplicadas = sorted({c for c in usadas if usadas.count(c) > 1})
        if duplicadas:
            problemas.append(
                f"{nombre}: la(s) columna(s) {duplicadas} se usan para más de un campo."
            )

        sin_mapear = [c for c, fuente in mapa.items() if fuente is None]
        if sin_mapear:
            mensaje = f"{nombre}: columnas sin mapear: {sin_mapear}"
            if exigir_mapeo_completo:
                problemas.append(mensaje)
            else:
                avisos.append(mensaje)

        for extra in extras.get(nombre, []):
            if extra not in perfil["columnas"]:
                problemas.append(f"{nombre}: la extra {extra!r} no existe.")
            if extra in columnas_ref:
                problemas.append(
                    f"{nombre}: {extra!r} ya es una columna de referencia."
                )
            if extra in usadas:
                problemas.append(
                    f"{nombre}: {extra!r} está incluida como extra y también "
                    "como fuente de un campo."
                )
            if extra in RESERVADAS:
                problemas.append(
                    f"{nombre}: {extra!r} es un nombre reservado por la trazabilidad."
                )

    return problemas, avisos


def construir_consolidado(
    perfiles: dict[str, PerfilArchivo],
    mapeo: dict[str, dict[str, str | None]],
    extras: dict[str, list[str]] | None,
    columnas_ref: list[str],
) -> pd.DataFrame:
    """Construye el DataFrame consolidado con trazabilidad."""
    extras = extras or {}
    partes: list[pd.DataFrame] = []

    for nombre, perfil in perfiles.items():
        datos = perfil["datos"]
        salida = pd.DataFrame(index=datos.index)

        for col_ref in columnas_ref:
            col_fuente = mapeo[nombre].get(col_ref)
            if col_fuente is None:
                salida[col_ref] = ""
            else:
                salida[col_ref] = datos[col_fuente]

        for extra in extras.get(nombre, []):
            salida[extra] = datos[extra]

        salida["archivo_origen"] = nombre
        salida["hoja_origen"] = str(perfil["hoja"])
        salida["fila_origen"] = datos.index + 1  # índice 1 = fila 2 de Excel
        partes.append(salida)

    return pd.concat(partes, ignore_index=True).fillna("")


# ---------------------------------------------------------------------------
# Flujo completo
# ---------------------------------------------------------------------------


def combinar_excels(
    rutas: list[str | Path],
    archivo_referencia: str,
    hoja_referencia: str | int = 0,
    hojas_por_archivo: dict[str, str | int] | None = None,
    ajustes: dict[str, dict[str, str | None]] | None = None,
    usar_sugerencias: bool = True,
    extras: dict[str, list[str]] | None = None,
    incluir_todos_los_extras: bool = False,
    exigir_mapeo_completo: bool = False,
    guardar_en: str | Path | None = None,
) -> pd.DataFrame:
    """Combina varios Excel usando la fila 1 de la referencia como encabezado.

    Parámetros
    ----------
    rutas:
        Archivos Excel a combinar. Los nombres deben ser únicos.
    archivo_referencia:
        Nombre del archivo cuyos títulos de la fila 1 serán el encabezado.
    hoja_referencia:
        Hoja de la referencia (índice o nombre). Por defecto, la primera.
    hojas_por_archivo:
        Overrides de hoja para archivos puntuales, por nombre de archivo.
    ajustes:
        Mapeo manual ``{archivo: {columna_ref: columna_fuente_o_None}}``.
    usar_sugerencias:
        Si es ``False``, todo empieza sin mapear (100 % manual).
    extras:
        Columnas extra aprobadas ``{archivo: [columna, ...]}``.
    incluir_todos_los_extras:
        Si es ``True``, incluye todas las extras detectadas.
    exigir_mapeo_completo:
        Si es ``True``, una columna sin mapear es un error.
    guardar_en:
        Ruta opcional del ``.xlsx`` de salida.

    Returns
    -------
    pandas.DataFrame
        Consolidado con columnas de la referencia, extras aprobadas y
        ``archivo_origen``, ``hoja_origen``, ``fila_origen``.
    """
    rutas = [Path(r).expanduser().resolve() for r in rutas]
    if not rutas:
        raise ValueError("No se recibió ningún archivo.")

    nombres = [r.name for r in rutas]
    duplicados = sorted({n for n in nombres if nombres.count(n) > 1})
    if duplicados:
        raise ValueError(
            f"Hay nombres de archivo repetidos: {duplicados}. "
            "Renombralos para poder distinguirlos."
        )

    hojas_por_archivo = {
        archivo_referencia: hoja_referencia,
        **(hojas_por_archivo or {}),
    }
    perfiles = {
        ruta.name: cargar_perfil(ruta, hojas_por_archivo.get(ruta.name, 0))
        for ruta in rutas
    }
    if archivo_referencia not in perfiles:
        raise FileNotFoundError(f"No se cargó {archivo_referencia!r}.")

    columnas_ref = perfiles[archivo_referencia]["columnas"]
    mapeo = construir_mapeo(perfiles, columnas_ref, ajustes, usar_sugerencias)

    extras_detectadas, _ = detectar_extras(
        perfiles, mapeo, columnas_ref, archivo_referencia
    )
    extras_finales = extras_detectadas if incluir_todos_los_extras else (extras or {})

    problemas, avisos = validar_configuracion(
        perfiles,
        mapeo,
        extras_finales,
        columnas_ref,
        exigir_mapeo_completo=exigir_mapeo_completo,
    )
    if problemas:
        raise ValueError("Revisá el mapeo:\n- " + "\n- ".join(problemas))

    consolidado = construir_consolidado(
        perfiles, mapeo, extras_finales, columnas_ref
    )
    consolidado.attrs["avisos"] = avisos
    consolidado.attrs["mapeo"] = mapeo

    if guardar_en is not None:
        destino = Path(guardar_en).expanduser().resolve()
        destino.parent.mkdir(parents=True, exist_ok=True)
        consolidado.to_excel(destino, index=False, sheet_name="CONSOLIDADO")

    return consolidado
