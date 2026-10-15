# Verificación del piloto Streamlit

Fecha: 15 de octubre de 2026.

## Alcance comprobado

Se ejecutaron **19 pruebas automatizadas satisfactorias** con Python 3.12.10, Streamlit 1.66.0, pandas 3.0.6, openpyxl 3.1.5 y Pillow 12.3.0.

Las pruebas cubren:

- E1 con materiales `Retirado`, cantidades negativas y aportación excluida por movimiento.
- E1 sin maestro externo y bloqueo de RG11 cuando el catálogo está vacío.
- E1 con 42 hojas, incluidas 40 auxiliares extensas: procesa únicamente `RG10_SER` y `RG36_MAT`.
- RG11 con 38 hojas, anexos extensos y hojas E1 adicionales: conserva la prioridad de `RG11_budget`.
- Conservación de los límites de tamaño global y de dimensiones en las hojas que sí se procesan; revisión completa de las hojas AGP.
- RG11 con cruce contra el maestro y conservación de ceros iniciales.
- Material de aportación presente en AGP: se mantiene visible como código solo en AGP.
- Consolidación de dos AGP y tres protocolos, con nombre de salida correspondiente.
- Estructura actual de cinco niveles y compatibilidad histórica de seis.
- Cantidades diferentes, códigos solo en una fuente y errores de conversión.
- PMS sin estructura reconocible: generación de Excel con hoja de errores.
- Aislamiento de dos ejecuciones con los mismos nombres de archivos y limpieza de temporales.
- Liberación del cupo del servidor incluso después de una excepción.
- Rechazo de archivos inválidos y hojas con dimensiones excesivas.
- Importaciones CSV, TXT y Excel, altas sin duplicados, bajas, vista previa y respaldos del catálogo.
- Exportación de textos que empiezan por `=` como texto literal, no como fórmulas ejecutables.
- Diez hojas del informe, logo OCA, filtros y paneles congelados.
- Navegación de Streamlit, consulta del maestro y reinicio del formulario.

Comando reproducible:

```powershell
python -m unittest discover -s tests -v
```

Además, un PMS RG11 con una columna de cantidad ausente produjo un Excel de validación incompleta con el mensaje de encabezados faltantes y liberó sus temporales.

## Comparación con tres casos reales existentes

Se utilizó el PMS `RG11PO1674I CTO PARAMO ALTO S10142.xlsx` y el maestro base `MAT APORT_VF_1.xlsx`, importado como 103 códigos únicos. Para cada resumen AGP se compararon tipo, código, cantidades, diferencia y estado contra el informe Excel anterior disponible.

| Resumen AGP | AGP detectados | Protocolos | Resultado frente al informe anterior | Tiempo local observado |
| --- | ---: | ---: | --- | ---: |
| Un solo AGP | 1 | 1 | Coincidencia de cantidades y estados | 3,44 s |
| Un AGP con varios protocolos | 1 | 3 | Coincidencia de cantidades y estados | 3,42 s |
| Dos AGP con varios protocolos | 2 | 3 | Coincidencia de cantidades y estados | 2,98 s |

Los tres casos produjeron 50 códigos PMS, 49 códigos AGP, 49 coincidencias, un código pendiente y cero errores de lectura. Estos resultados no significan que el caso de negocio esté completamente correcto: se conserva la diferencia que ya mostraba el informe anterior.

Los archivos reales y sus resultados no están incluidos en el paquete de publicación.

## Revisión en navegador

Se comprobó con Chrome en servidor local:

1. Presentación de los dos selectores de archivo.
2. Carga del PMS y del resumen de dos AGP.
3. Ejecución, visualización del semáforo y conteos.
4. Descarga del Excel `validacion_materiales_prestaciones_AGP_2000139288_2000139289.xlsx`.
5. Apertura de otra sesión sin acceso al resultado de la primera.
6. Acceso al manual y al catálogo.
7. Reinicio del caso, retirando el resultado anterior.
8. Vista en pantalla de escritorio y móvil de 390 píxeles de ancho.

No se observaron errores JavaScript en ese recorrido. Se inspeccionaron visualmente la pantalla inicial, el resultado, el manual, el maestro y la vista móvil.

## Límites de esta verificación

Los tiempos corresponden al equipo local y no permiten prometer los mismos tiempos en Community Cloud. No se ha medido un escenario de 20 usuarios cargando archivos simultáneamente ni la disponibilidad del servicio gratuito.

La aplicación limita el procesamiento a una validación a la vez por proceso del servidor y devuelve un mensaje de ocupación a las demás solicitudes. Las sesiones abiertas y los archivos cargados también consumen memoria, aunque todavía no estén en procesamiento.

El despliegue público en GitHub/Streamlit lo realizará el administrador. Sigue pendiente verificar desde la red de Bureau Veritas que la URL publicada permita abrir la aplicación, cargar archivos y descargar resultados. También queda pendiente la adopción del encabezado documental del nuevo gestor; el Excel conserva F-AGP-07 con OCA.
