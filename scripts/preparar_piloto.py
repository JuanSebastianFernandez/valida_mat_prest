"""Genera un ZIP con una lista explicita de archivos publicables, sin casos de trabajo."""

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

RAIZ = Path(__file__).resolve().parents[1]
ARCHIVOS = [
    "streamlit_app.py", "validacion_web.py", "catalogo_aportacion.py",
    "validacion_materiales_prestaciones_AGP.py", "requirements.txt", "README.md",
    ".gitignore", ".streamlit/config.toml", "catalogo/materiales_aportacion.csv",
    "docs/MANUAL_USUARIO.md", "docs/ADMINISTRACION.md", "docs/VALIDACION_PILOTO.md",
    "scripts/actualizar_maestro.py", "scripts/preparar_piloto.py", "tests/test_piloto.py",
]


def main() -> None:
    faltantes = [nombre for nombre in ARCHIVOS if not (RAIZ / nombre).is_file()]
    if faltantes:
        raise SystemExit(f"Faltan archivos de publicacion: {faltantes}")
    destino = RAIZ / "dist" / "validador_AGP_streamlit.zip"
    destino.parent.mkdir(exist_ok=True)
    with ZipFile(destino, "w", ZIP_DEFLATED) as paquete:
        for nombre in ARCHIVOS:
            paquete.write(RAIZ / nombre, nombre)
    print(f"Paquete listo: {destino}")
    print(f"{len(ARCHIVOS)} archivos; sin casos, resultados, tokens ni carpetas temporales.")


if __name__ == "__main__":
    main()
