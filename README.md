# Validador PMS vs AGP · Piloto Streamlit

Aplicación para comparar materiales y prestaciones del PMS con el resumen exportado desde AGP. Permite cargar dos Excel, revisar diferencias y descargar el informe con el formato existente F-AGP-07. Incluye manual operativo y catálogo de aportación sin depender de Colab ni Drive.

## Ejecutar localmente

Requiere Python 3.12 (también debe funcionar con 3.11). Desde la carpeta del proyecto:

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python -m streamlit run streamlit_app.py
```

Abra la dirección local que indique Streamlit. No necesita activar el entorno en PowerShell.

## Publicar en GitHub y Streamlit

1. Use el contenido de `dist/validador_AGP_streamlit.zip` como carpeta de publicación. Ese paquete incluye únicamente los archivos necesarios para el piloto; no incluye casos de prueba, resultados ni otras herramientas del espacio de trabajo.
2. Cree su repositorio público en GitHub y cargue el contenido **descomprimido**, conservando `.streamlit`, `catalogo`, `docs`, `scripts` y `tests`.
3. En Streamlit Community Cloud cree una aplicación desde ese repositorio y la rama publicada, normalmente `main`.
4. Seleccione `streamlit_app.py` como archivo principal y Python 3.12 en la configuración avanzada.
5. Publique y pruebe la URL desde un equipo y la red de Bureau Veritas: apertura, carga de ambos archivos y descarga del Excel.

No se requieren tokens, secretos, cuentas de Google ni base de datos. GitHub conserva el código y el maestro; Streamlit ejecuta Python. No se ha incluido autenticación: el piloto publicado es accesible para cualquiera que llegue a su URL.

**Repositorio público:** cualquiera puede leer y copiar el código y el catálogo. Eso no concede permiso de escritura. Mantenga su cuenta GitHub como única cuenta con escritura y no agregue colaboradores con ese permiso. Los archivos de casos no deben cargarse al repositorio; la aplicación los procesa de forma temporal en el servidor.

El formato Excel conserva el encabezado OCA y el código documental histórico. El portal es neutral de Dirección AGP; no se ha inventado un logo ni una aprobación documental de Bureau Veritas.

## Manuales

- [Manual del dibujante](docs/MANUAL_USUARIO.md): disponible también dentro de la aplicación.
- [Administración y actualización del maestro](docs/ADMINISTRACION.md): comandos de importación, altas, bajas y publicación.
- [Verificación del piloto](docs/VALIDACION_PILOTO.md): pruebas y límites observados.

## Componentes

| Archivo | Función |
| --- | --- |
| `streamlit_app.py` | Interfaz y sesión individual. |
| `validacion_web.py` | Límites de archivos, aislamiento y adaptación del motor. |
| `validacion_materiales_prestaciones_AGP.py` | Lectura, reglas, comparación y exportación; reutilizable en Colab. |
| `catalogo_aportacion.py` | Lectura y huella de versión del catálogo. |
| `catalogo/materiales_aportacion.csv` | Maestro vigente, una fila por código. |
| `scripts/actualizar_maestro.py` | Actualizaciones locales con deduplicación y vista previa. |
| `.streamlit/config.toml` | Tema visual y límite de carga. |

## Límites del piloto

Un caso en procesamiento simultáneo por proceso de servidor; el resto recibe un aviso para reintentar. Hasta 20 MB por archivo y 96 MB de contenido descomprimido, con un máximo de 2.048 elementos internos en el Excel. No hay un límite fijo de 30 hojas: los PMS E1 pueden incluir hojas auxiliares adicionales.

Los límites de 50.000 filas por hoja, 200 columnas y 500.000 celdas se comprueban sobre las hojas que se procesan: `RG10_SER` y `RG36_MAT` en E1, o `RG11_budget` en RG11. Una hoja procesada tampoco puede superar 500.000 posiciones rectangulares. Las hojas auxiliares del PMS no se comparan ni se revisan por dimensiones. En AGP se revisan todas las hojas, porque el motor necesita identificar la hoja que contiene el resumen. Son límites preventivos del piloto, no la capacidad garantizada del alojamiento.

No hay cola persistente, histórico de resultados ni recuperación automática de sesión. No se ha certificado capacidad para 20 ejecuciones simultáneas. El servicio gratuito puede entrar en reposo o reiniciarse; mida el uso real antes de adoptarlo como dependencia obligatoria.

## Pruebas y paquete de publicación

```powershell
python -m unittest discover -s tests -v
python scripts/preparar_piloto.py
```

El segundo comando genera un ZIP de publicación mediante una lista explícita de archivos permitidos. No use `git add .` sobre una carpeta que contenga otros proyectos o archivos de casos.

Referencias técnicas: [Streamlit Community Cloud](https://docs.streamlit.io/deploy/streamlit-community-cloud), [gestión de recursos y reposo](https://docs.streamlit.io/deploy/streamlit-community-cloud/manage-your-app), [archivos cargados](https://docs.streamlit.io/knowledge-base/using-streamlit/where-file-uploader-store-when-deleted).

© 2026 Dirección AGP. Desarrollo: Juan Sebastián Fernández Buitrago.
