"""Piloto web de validacion PMS vs AGP. Ejecutar: streamlit run streamlit_app.py."""

from __future__ import annotations

from pathlib import Path
from threading import BoundedSemaphore

import pandas as pd
import streamlit as st

from catalogo_aportacion import cargar_catalogo
from validacion_web import MAX_ARCHIVO_MB, VERSION_APP, ServidorOcupado, validar_archivos

RAIZ = Path(__file__).parent

ESTILOS = """
<style>
.block-container { max-width:1180px; padding-top:2.4rem; padding-bottom:3rem; }
[data-testid="stSidebar"] { background:#eaf0f2; }
.agp-hero { background:#142f43; color:#fff; border-radius:18px;
  padding:30px 34px; margin-bottom:24px; border-bottom:5px solid #4ec4a5; }
.agp-eyebrow { color:#8de0c8; font-size:12px; font-weight:700;
  letter-spacing:2px; margin-bottom:12px; }
.agp-hero h1 { color:#fff !important; font-size:34px; line-height:1.16;
  padding:0 0 12px; margin:0; letter-spacing:-.7px; }
.agp-hero p { color:#dbe7ee; margin:0; max-width:770px; font-size:16px; }
.agp-note { background:#fff5dc; color:#64440b; border-left:4px solid #ba840e;
  padding:15px 18px; border-radius:8px; margin:12px 0 22px; line-height:1.5; }
.agp-note strong { color:#573800; }
.agp-empty { border:1px dashed #aebfc6; border-radius:12px; padding:25px;
  color:#526976; text-align:center; background:#f0f5f5; margin:18px 0; }
.agp-footer { border-top:1px solid #d4dfe4; margin-top:32px; padding-top:16px;
  color:#576d7a; font-size:12px; }
[data-testid="stMetric"] { background:white; border:1px solid #dfe7eb;
  border-radius:12px; padding:14px 18px; }
.agp-indicadores { display:grid; grid-template-columns:repeat(4,minmax(0,1fr));
  gap:14px; margin:18px 0 10px; }
.agp-tarjeta { border:1px solid; border-top-width:4px; border-radius:12px;
  padding:16px 18px; min-width:0; }
.agp-tarjeta--ok { background:#e6f4ea; border-color:#81bc96; color:#14532d; }
.agp-tarjeta--error { background:#fde8e8; border-color:#e08a8a; color:#8f1d1d; }
.agp-tarjeta--info { background:#eaf2fb; border-color:#9db9d5; color:#234e70; }
.agp-tarjeta-label { font-size:14px; line-height:1.4; font-weight:600; }
.agp-tarjeta-valor { font-size:34px; line-height:1.2; font-weight:700; margin:7px 0; }
.agp-tarjeta-estado { font-size:12px; line-height:1.4; font-weight:600; }
@media(max-width:1000px) {
 .agp-indicadores { grid-template-columns:repeat(2,minmax(0,1fr)); }
}
@media(max-width:640px) {
 .agp-hero { padding:23px 20px; } .agp-hero h1 { font-size:26px; }
 .block-container { padding-top:4.5rem; }
}
</style>
"""


@st.cache_resource
def capacidad_servidor() -> BoundedSemaphore:
    # Compartido por las sesiones del proceso, sin compartir archivos ni resultados.
    return BoundedSemaphore(1)


def limpiar_resultado() -> None:
    st.session_state.pop("resultado", None)
    st.session_state.pop("mensaje_error", None)


def reiniciar_caso() -> None:
    limpiar_resultado()
    st.session_state["carga_id"] = st.session_state.get("carga_id", 0) + 1


def mostrar_manual() -> None:
    st.subheader("Manual de uso")
    st.caption("Guía operativa para validar el caso antes de enviarlo.")
    contenido = (RAIZ / "docs" / "MANUAL_USUARIO.md").read_text(encoding="utf-8")
    st.download_button(
        "Descargar manual", contenido, "Manual_Validador_AGP.md", "text/markdown",
        on_click="ignore",
    )
    st.markdown(contenido)


def mostrar_maestro() -> None:
    st.subheader("Maestro de materiales de aportación")
    st.write("Este catálogo se aplica automáticamente a PMS con hoja RG11_budget.")
    try:
        catalogo = cargar_catalogo()
    except ValueError as exc:
        st.error(str(exc))
        return
    a, b = st.columns([1, 3])
    a.metric("Materiales", len(catalogo.codigos))
    b.caption("Version del maestro (SHA-256)")
    b.code(catalogo.version, language=None)
    if not catalogo.codigos:
        st.warning("El administrador debe importar el maestro inicial antes de validar RG11.")
    busqueda = st.text_input("Buscar código", placeholder="Escriba parte del código")
    tabla = pd.DataFrame({"CODIGO": list(catalogo.codigos)})
    if busqueda.strip():
        tabla = tabla[tabla["CODIGO"].str.contains(busqueda.strip(), regex=False)]
    st.dataframe(tabla, hide_index=True, width="stretch", height=360)
    st.download_button(
        "Descargar maestro vigente", catalogo.contenido, "materiales_aportacion.csv",
        "text/csv", on_click="ignore",
    )
    st.info(
        "La actualización del maestro corresponde al administrador del repositorio. "
        "Los usuarios de esta página solo pueden consultarlo."
    )


def mostrar_indicadores(datos: dict) -> None:
    indicadores = [
        ("Códigos PMS", "TOTAL_CODIGOS_PMS", "info"),
        ("Códigos AGP", "TOTAL_CODIGOS_AGP", "info"),
        ("Coinciden", "TOTAL_OK", "coincidencias"),
        ("Diferencias de cantidad", "TOTAL_DIFERENCIAS_CANTIDAD", "incidencias"),
        ("Solo en PMS", "TOTAL_SOLO_PMS", "incidencias"),
        ("Solo en AGP", "TOTAL_SOLO_AGP", "incidencias"),
        ("Aportación excluida", "TOTAL_APORTACION_EXCLUIDOS", "info"),
        ("Errores", "TOTAL_ERRORES_REGISTRADOS", "incidencias"),
    ]
    tarjetas = []
    for etiqueta, clave, grupo in indicadores:
        cantidad = int(datos[clave])
        if grupo == "incidencias":
            color, estado = ("error", "Requiere revisión") if cantidad else ("ok", "Sin incidencias")
        elif grupo == "coincidencias":
            color, estado = ("ok", "Coincidencia correcta") if cantidad else ("info", "Sin coincidencias")
        else:
            color, estado = "info", "Auditoría de aportación" if clave == "TOTAL_APORTACION_EXCLUIDOS" else "Total de referencia"
        # Las etiquetas son constantes y los valores se convierten a enteros antes del HTML.
        tarjetas.append(
            f'<div class="agp-tarjeta agp-tarjeta--{color}" role="group" aria-label="{etiqueta}">'
            f'<div class="agp-tarjeta-label">{etiqueta}</div>'
            f'<div class="agp-tarjeta-valor">{cantidad}</div>'
            f'<div class="agp-tarjeta-estado">{estado}</div></div>'
        )
    st.markdown('<div class="agp-indicadores">' + "".join(tarjetas) + '</div>', unsafe_allow_html=True)
    st.caption(
        "Verde: coincidencias o indicador sin incidencias. Rojo: requiere revisión. "
        "Azul: información de referencia o auditoría. Revise siempre el resultado general."
    )


def mostrar_resultado(resultado: dict) -> None:
    datos = resultado["datos_resumen"]
    st.divider()
    st.subheader("Resultado de la validación")
    mensaje = datos["RESULTADO_GENERAL"]
    if datos["SEMAFORO"] == "OK":
        st.success(mensaje, icon=":material/check_circle:")
    elif datos["SEMAFORO"] == "ERROR":
        st.error(mensaje, icon=":material/error:")
    else:
        st.warning(mensaje, icon=":material/warning:")
    st.caption(
        f"PMS: {datos['ARCHIVO_PMS']} | AGP: {datos['ARCHIVO_AGP']} | "
        f"Procesado en {resultado['duracion_segundos']:.1f} s"
    )
    st.write(f"**AGP:** {datos['CODIGO_AGP']}")
    st.write(f"**Protocolos:** {datos['PROTOCOLO']}")
    mostrar_indicadores(datos)
    st.download_button(
        "Descargar Excel de validación", resultado["archivo_excel"],
        resultado["nombre_excel"],
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", width="stretch", on_click="ignore",
    )
    st.caption("Descargue y archive el Excel. Esta página no conserva un histórico de casos.")
    opciones = ["Comparación", "PMS normalizado", "AGP normalizado", "Aportación excluida", "Errores"]
    vista = st.selectbox("Tabla a consultar", opciones)
    claves = {
        "Comparación": "comparacion", "PMS normalizado": "pms_normalizado",
        "AGP normalizado": "agp_normalizado", "Aportación excluida": "excluidos",
        "Errores": "errores",
    }
    tabla = resultado[claves[vista]]
    if vista == "Comparación":
        filtro = st.radio("Mostrar", ["Pendientes de revisión", "Todos"], horizontal=True)
        if filtro == "Pendientes de revisión":
            tabla = tabla[tabla["ESTADO_VALIDACION"] != "OK"]
    if tabla.empty:
        st.info("No hay registros para esta selección.")
    else:
        st.dataframe(tabla.head(500), hide_index=True, width="stretch")
        if len(tabla) > 500:
            st.caption(f"Vista previa: 500 de {len(tabla)} registros. El Excel contiene todos.")


def mostrar_validacion() -> None:
    st.subheader("Prepare y cargue el caso")
    st.markdown(
        '<div class="agp-note"><strong>Antes de cargar:</strong> en E1 revise '
        '<strong>TIPO DE MOVIMIENTO</strong>: aportación se excluye y Retirado se lleva '
        'a negativo. Si hay varios AGP o protocolos, incluya todos en un solo resumen '
        'de proyecto y conserve sus niveles de agrupación.</div>', unsafe_allow_html=True,
    )
    catalogo = None
    try:
        catalogo = cargar_catalogo()
        if catalogo.codigos:
            st.caption(f"Maestro activo: {len(catalogo.codigos)} materiales | Versión {catalogo.version[:12]}")
        else:
            st.warning("Maestro inicial pendiente. Puede validar E1; RG11 requiere importar el maestro.")
    except ValueError as exc:
        st.warning(f"{exc} Puede validar E1; RG11 requiere un maestro disponible.")
    carga_id = st.session_state.get("carga_id", 0)
    izquierda, derecha = st.columns(2, gap="large")
    with izquierda:
        with st.container(border=True):
            st.markdown("#### 01 · Archivo original PMS")
            st.caption("RG11_budget o las hojas E1 RG10_SER y RG36_MAT.")
            pms = st.file_uploader(
                "Cargar PMS", type=["xlsx", "xlsm"], key=f"pms_{carga_id}",
                on_change=limpiar_resultado,
            )
    with derecha:
        with st.container(border=True):
            st.markdown("#### 02 · Resumen exportado de AGP")
            st.caption("Archivo data con todos los AGP y protocolos del caso.")
            agp = st.file_uploader(
                "Cargar AGP", type=["xlsx", "xlsm"], key=f"agp_{carga_id}",
                on_change=limpiar_resultado,
            )
    st.caption(f"Límite: {MAX_ARCHIVO_MB} MB por archivo. No necesita Google Drive ni cuentas de Google.")
    validar, nuevo = st.columns([3, 1])
    ejecutar = validar.button(
        "Validar materiales y prestaciones", type="primary", width="stretch",
        disabled=pms is None or agp is None,
    )
    nuevo.button("Nuevo caso", on_click=reiniciar_caso, width="stretch")
    if ejecutar:
        limpiar_resultado()
        try:
            with st.spinner("Leyendo archivos, comparando y preparando el Excel..."):
                st.session_state.resultado = validar_archivos(
                    pms.name, pms.getvalue(), agp.name, agp.getvalue(),
                    catalogo, capacidad_servidor(),
                )
        except (ValueError, ServidorOcupado) as exc:
            st.session_state.mensaje_error = str(exc)
        except Exception:
            # No mostrar rutas internas ni contenido de los archivos en una pagina publica.
            st.session_state.mensaje_error = (
                "No se pudo generar el informe. Verifique los archivos e intente otra vez. "
                "Si persiste, informe al administrador los nombres y tamaños de los archivos."
            )
    resultado = st.session_state.get("resultado")
    if resultado is not None and catalogo is not None:
        version_usada = resultado["datos_resumen"]["VERSION_MAESTRO_SHA256"]
        if version_usada != catalogo.version and not version_usada.startswith("No aplica:"):
            limpiar_resultado()
            resultado = None
            st.warning("El maestro cambió. Ejecute otra vez la validación para usar la versión vigente.")
    if st.session_state.get("mensaje_error"):
        st.error(st.session_state.mensaje_error)
    if resultado is not None:
        mostrar_resultado(resultado)
    else:
        st.markdown(
            '<div class="agp-empty">El resultado y la descarga aparecerán aquí '
            'después de validar sus archivos.</div>', unsafe_allow_html=True,
        )


def main() -> None:
    st.set_page_config(page_title="Validador PMS · AGP", layout="wide")
    st.markdown(ESTILOS, unsafe_allow_html=True)
    with st.sidebar:
        st.markdown("### DIRECCIÓN AGP")
        st.caption("CONTROL PREVIO DE VALORIZACIÓN")
        pagina = st.radio("Navegación", ["Validar caso", "Manual de uso", "Maestro de aportación"])
        st.divider()
        st.markdown("**Un caso. Dos archivos. Un informe.**")
        st.caption("Cargue PMS y AGP, revise las alertas y descargue su evidencia.")
        st.caption(f"Versión {VERSION_APP}")
    st.markdown(
        '<div class="agp-hero"><div class="agp-eyebrow">CONTROL DE MATERIALES Y PRESTACIONES</div>'
        '<h1>Valide antes de enviar.</h1><p>Compare el PMS original con lo registrado en AGP. '
        'Identifique diferencias de códigos y cantidades y genere el Excel de validación.</p></div>',
        unsafe_allow_html=True,
    )
    if pagina == "Manual de uso":
        mostrar_manual()
    elif pagina == "Maestro de aportación":
        mostrar_maestro()
    else:
        mostrar_validacion()
    st.markdown(
        '<div class="agp-footer">© 2026 Dirección AGP · Juan Sebastián Fernández Buitrago<br>'
        'Herramienta de apoyo a la revisión. La coincidencia de cantidades no sustituye '
        'la revisión técnica ni constituye una aprobación del caso.</div>', unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
