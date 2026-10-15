"""Adaptador web: limites, aislamiento y descarga; las reglas viven en el motor original."""

from __future__ import annotations

import hashlib
import io
import posixpath
import tempfile
import time
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from threading import BoundedSemaphore
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile
from zoneinfo import ZoneInfo

import pandas as pd

import validacion_materiales_prestaciones_AGP as motor
from catalogo_aportacion import Catalogo

MAX_ARCHIVO_MB = 20
MAX_XML_BYTES = 96 * 1024 * 1024
MAX_CELDAS = 500_000
MAX_FILAS = 50_000
MAX_COLUMNAS = 200
VERSION_APP = "1.0.0-piloto"


class ServidorOcupado(ValueError):
    pass


@contextmanager
def reservar_proceso(semaforo: BoundedSemaphore):
    if not semaforo.acquire(blocking=False):
        raise ServidorOcupado(
            "Hay otra validacion en curso. Sus archivos siguen cargados; "
            "espere un momento y presione Validar nuevamente."
        )
    try:
        yield
    finally:
        semaforo.release()


def nombre_seguro(nombre: str) -> str:
    nombre = str(nombre).replace("\\", "/").rsplit("/", 1)[-1].strip()
    if not nombre or len(nombre) > 180 or any(ord(c) < 32 for c in nombre):
        raise ValueError("El nombre del archivo no es valido. Renombrelo e intente otra vez.")
    if Path(nombre).suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Cargue un archivo Excel .xlsx o .xlsm. El formato .xls no es compatible.")
    return nombre


def _hojas_relevantes_pms(archivo: ZipFile) -> set[str] | None:
    """Resuelve las hojas de datos por nombre e ID, sin asumir su orden en el libro."""
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    ns_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    libro = ET.fromstring(archivo.read("xl/workbook.xml"))
    hojas = {
        hoja.attrib["name"].strip(): hoja.attrib[f"{{{ns_rel}}}id"]
        for hoja in libro.findall("s:sheets/s:sheet", ns)
    }
    if "RG11_budget" in hojas:
        nombres = ("RG11_budget",)
    elif {"RG10_SER", "RG36_MAT"}.issubset(hojas):
        nombres = ("RG10_SER", "RG36_MAT")
    else:
        # La deteccion del motor reportara las hojas faltantes; no omitir el prefiltro.
        return None
    relaciones = ET.fromstring(archivo.read("xl/_rels/workbook.xml.rels"))
    destinos = {}
    for relacion in relaciones:
        if relacion.attrib.get("TargetMode") == "External":
            continue
        if not relacion.attrib.get("Type", "").endswith("/worksheet"):
            continue
        destino = relacion.attrib["Target"]
        destinos[relacion.attrib["Id"]] = posixpath.normpath(
            destino.lstrip("/") if destino.startswith("/") else posixpath.join("xl", destino)
        )
    return {destinos[hojas[nombre]] for nombre in nombres}


def verificar_excel(contenido: bytes, etiqueta: str, *, es_pms: bool = False) -> None:
    if not contenido:
        raise ValueError(f"El archivo {etiqueta} esta vacio.")
    if len(contenido) > MAX_ARCHIVO_MB * 1024 * 1024:
        raise ValueError(f"El archivo {etiqueta} supera el limite de {MAX_ARCHIVO_MB} MB.")
    try:
        with ZipFile(io.BytesIO(contenido)) as archivo:
            elementos = archivo.infolist()
            if len(elementos) > 2048 or sum(e.file_size for e in elementos) > MAX_XML_BYTES:
                raise ValueError(f"El archivo {etiqueta} es demasiado grande al descomprimirse.")
            if "xl/workbook.xml" not in archivo.namelist():
                raise ValueError(f"El archivo {etiqueta} no contiene un libro Excel valido.")
            hojas = [e for e in elementos if e.filename.startswith("xl/worksheets/")
                     and e.filename.endswith(".xml") and "/_rels/" not in e.filename]
            relevantes = _hojas_relevantes_pms(archivo) if es_pms else None
            if relevantes is not None:
                disponibles = {hoja.filename for hoja in hojas}
                if not relevantes.issubset(disponibles):
                    raise KeyError("No se encontro el contenido de las hojas PMS requeridas.")
                hojas = [hoja for hoja in hojas if hoja.filename in relevantes]
            celdas = 0
            for hoja in hojas:
                with archivo.open(hoja) as entrada:
                    for _, nodo in ET.iterparse(entrada, events=("end",)):
                        tag = nodo.tag.rsplit("}", 1)[-1]
                        if tag == "dimension":
                            referencia = nodo.attrib.get("ref", "A1").replace("$", "")
                            _, _, col, fila = motor.openpyxl.utils.range_boundaries(referencia)
                            if fila > MAX_FILAS or col > MAX_COLUMNAS or fila * col > MAX_CELDAS:
                                raise ValueError(
                                    f"El archivo {etiqueta} tiene una hoja demasiado extensa. "
                                    "Exporte un resumen limpio sin filas o columnas sobrantes."
                                )
                        elif tag == "row" and int(nodo.attrib.get("r", "0")) > MAX_FILAS:
                            raise ValueError(f"El archivo {etiqueta} supera {MAX_FILAS} filas por hoja.")
                        elif tag == "c":
                            celdas += 1
                            direccion = nodo.attrib.get("r", "A1")
                            _, col = motor.openpyxl.utils.cell.coordinate_to_tuple(direccion)
                            if col > MAX_COLUMNAS or celdas > MAX_CELDAS:
                                raise ValueError(f"El archivo {etiqueta} supera el limite de celdas del piloto.")
                        nodo.clear()
    except (BadZipFile, ET.ParseError, KeyError, RuntimeError, OverflowError) as exc:
        raise ValueError(
            f"No se pudo abrir {etiqueta}. Verifique que no este protegido con contrasena "
            "y que sea un Excel valido."
        ) from exc


def firma_entradas(nombre_pms: str, pms: bytes, nombre_agp: str, agp: bytes, version: str) -> str:
    firma = hashlib.sha256()
    for parte in (nombre_pms.encode(), pms, nombre_agp.encode(), agp, version.encode()):
        firma.update(len(parte).to_bytes(8, "big"))
        firma.update(parte)
    return firma.hexdigest()


def validar_archivos(
    nombre_pms: str, pms: bytes, nombre_agp: str, agp: bytes,
    catalogo: Catalogo | None, semaforo: BoundedSemaphore,
) -> dict:
    """No conserva rutas, libros ni archivos fuente despues de cada ejecucion."""
    with reservar_proceso(semaforo):
        inicio = time.monotonic()
        nombre_pms, nombre_agp = nombre_seguro(nombre_pms), nombre_seguro(nombre_agp)
        verificar_excel(pms, "PMS", es_pms=True)
        verificar_excel(agp, "AGP")
        with tempfile.TemporaryDirectory(prefix="agp_validacion_") as temporal:
            raiz = Path(temporal)
            token = motor._DIRECTORIO_TEMPORAL.set(raiz)
            try:
                ruta_pms = raiz / "pms" / nombre_pms
                ruta_agp = raiz / "agp" / nombre_agp
                for ruta, contenido in ((ruta_pms, pms), (ruta_agp, agp)):
                    ruta.parent.mkdir()
                    ruta.write_bytes(contenido)
                try:
                    tipo_pms = motor.detectar_tipo_pms(str(ruta_pms))
                except motor.ErrorValidacion:
                    tipo_pms = "No detectado"
                codigos = catalogo.codigos if catalogo else ()
                if tipo_pms == motor.TIPO_RG11 and not codigos:
                    raise ValueError(
                        "El maestro esta vacio o no esta disponible. El administrador debe importar el maestro "
                        "antes de validar PMS RG11. E1 utiliza TIPO DE MOVIMIENTO."
                    )
                ruta_maestro = raiz / "maestro_aportacion.xlsx"
                pd.DataFrame({"MAT APORT": list(codigos)}).to_excel(ruta_maestro, index=False)
                version_maestro = catalogo.version if tipo_pms == motor.TIPO_RG11 and catalogo else (
                    "No aplica: E1" if tipo_pms == motor.TIPO_E1 else "No aplica: PMS no detectado"
                )
                resultado = motor.ejecutar_validacion(
                    str(ruta_pms), str(ruta_agp), str(ruta_maestro), str(raiz / "salida"),
                    mostrar_resultado=False,
                    metadatos_extra={
                        "FECHA_HORA_VALIDACION": datetime.now(ZoneInfo("America/Bogota")).strftime("%Y-%m-%d %H:%M:%S"),
                        "ZONA_HORARIA": "America/Bogota",
                        "VERSION_HERRAMIENTA": VERSION_APP,
                        "VERSION_MAESTRO_SHA256": version_maestro,
                        "TOTAL_MATERIALES_MAESTRO": len(codigos) if tipo_pms == motor.TIPO_RG11 else "No aplica",
                        "SHA256_PMS": hashlib.sha256(pms).hexdigest(),
                        "SHA256_AGP": hashlib.sha256(agp).hexdigest(),
                    },
                )
                salida = Path(resultado.pop("ruta_salida"))
                resultado["archivo_excel"] = salida.read_bytes()
                resultado["nombre_excel"] = salida.name
                resultado["duracion_segundos"] = round(time.monotonic() - inicio, 2)
                resultado["firma"] = firma_entradas(nombre_pms, pms, nombre_agp, agp, catalogo.version if catalogo else "No disponible")
                return resultado
            finally:
                motor._DIRECTORIO_TEMPORAL.reset(token)
