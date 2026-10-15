"""Catalogo versionado de aportacion. No necesita credenciales ni servicios externos."""

from __future__ import annotations

import csv
import hashlib
import io
import re
from dataclasses import dataclass
from pathlib import Path

from validacion_materiales_prestaciones_AGP import normalizar_codigo

RUTA_CATALOGO = Path(__file__).parent / "catalogo" / "materiales_aportacion.csv"
CODIGO_VALIDO = re.compile(r"[0-9]+\Z", re.ASCII)


@dataclass(frozen=True)
class Catalogo:
    codigos: tuple[str, ...]
    version: str
    contenido: bytes


def validar_codigo_material(valor: object) -> str:
    codigo = normalizar_codigo(valor)
    if not CODIGO_VALIDO.fullmatch(codigo):
        raise ValueError(f"Codigo de material invalido: {valor!r}. Debe contener solo digitos.")
    return codigo


def serializar_catalogo(codigos: list[str] | tuple[str, ...] | set[str]) -> bytes:
    salida = io.StringIO(newline="")
    writer = csv.writer(salida, lineterminator="\n")
    writer.writerow(["CODIGO"])
    for codigo in sorted({validar_codigo_material(c) for c in codigos}):
        writer.writerow([codigo])
    return salida.getvalue().encode("utf-8")


def cargar_catalogo(ruta: Path = RUTA_CATALOGO, *, permitir_duplicados: bool = False) -> Catalogo:
    try:
        contenido = ruta.read_bytes()
        reader = csv.DictReader(io.StringIO(contenido.decode("utf-8-sig")))
        if reader.fieldnames != ["CODIGO"]:
            raise ValueError("El catalogo debe tener una unica columna llamada CODIGO.")
        codigos = []
        for numero, fila in enumerate(reader, start=2):
            if None in fila:
                raise ValueError(f"La fila {numero} tiene columnas adicionales.")
            codigos.append(validar_codigo_material(fila.get("CODIGO", "")))
        if len(set(codigos)) != len(codigos) and not permitir_duplicados:
            raise ValueError("El catalogo contiene codigos repetidos. Ejecute el comando normalizar.")
        return Catalogo(tuple(sorted(set(codigos))), hashlib.sha256(contenido).hexdigest(), contenido)
    except (OSError, UnicodeError, ValueError, csv.Error) as exc:
        raise ValueError(f"No se pudo cargar el maestro de aportacion: {exc}") from exc
