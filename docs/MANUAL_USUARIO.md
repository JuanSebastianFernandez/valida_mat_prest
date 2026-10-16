# Guía de validación de materiales y prestaciones

## 1. Finalidad

Esta herramienta ayuda al dibujante a verificar que los materiales y prestaciones registrados en AGP coincidan con el PMS original antes de enviar el caso. Busca reducir errores de cantidades, omisiones, inclusiones incorrectas, reprocesos y PNC.

Compara **tipo, código y cantidad**, con cantidades sumadas por código. No verifica precios, tarifas, valores monetarios, aprobación técnica ni ejecución en campo. Un resultado OK confirma coincidencia en los datos comparados; no constituye una aprobación del caso.

## 2. Qué necesita

- El archivo PMS original en `.xlsx` o `.xlsm`.
- El resumen de proyecto exportado desde AGP, habitualmente llamado `data.xlsx`.
- Navegador y acceso a la dirección de la aplicación.

El maestro de materiales de aportación está incorporado en la aplicación. El dibujante no tiene que cargarlo. No se requiere Google Colab, Google Drive, Gmail ni ejecutar código. El piloto acepta hasta 60 MB por archivo, incluidos los de 40 y 50 MB; los libros con dimensiones excesivas requieren una exportación limpia. También existe un límite de 256 MB de contenido interno descomprimido por libro. Los archivos grandes pueden tardar más en procesarse.

## 3. Revise el PMS antes de cargarlo

### PMS RG11

Debe contener la hoja `RG11_budget`, con columnas `cod_prest/mat` y `quantity`. El sistema busca los encabezados aunque no estén en la primera fila y admite algunos nombres equivalentes.

Esta estructura se utiliza tanto para flujo nuevo como para flujo viejo fuera de E1. La herramienta no distingue esos dos procesos por el nombre de la hoja: aplica el mismo lector y cruza los materiales contra el maestro vigente de aportación.

### PMS E1: revisión obligatoria del movimiento en el archivo fuente

Debe contener estas hojas:

| Hoja | Columnas necesarias |
| --- | --- |
| `RG10_SER` | `PRESTACIÓN`, `CANTIDAD` |
| `RG36_MAT` | `CÓDIGO MATERIAL E4E`, `CANTIDAD`, `TIPO DE MOVIMIENTO` |

**Antes de cargar E1, revise `TIPO DE MOVIMIENTO` para afinar los materiales de aportación y retirados.**

- **Material de aportación:** se excluye del PMS valorizable y se reporta en auditoría.
- **Retirado:** se compara con cantidad negativa. Tanto `1` como `-1` quedan en `-1`; la regla no invierte por segunda vez un signo que ya era negativo.
- **Chatarra:** revise especialmente estos materiales. La herramienta no deduce que un material es retirado solo por su descripción. Si el movimiento fue diligenciado incorrectamente, corrija el PMS según la revisión técnica antes de validar.

En E1 se utiliza la clasificación del propio PMS; no se reemplaza por el maestro de aportación de RG11.

El archivo E1 puede traer más de 30 hojas. No tiene que borrar las hojas auxiliares: la validación utiliza `RG10_SER` y `RG36_MAT` y omite las demás hojas del PMS. En RG11 utiliza `RG11_budget`. Se mantienen los límites de tamaño del archivo y de las hojas de datos que sí se procesan.

## 4. Prepare el resumen de AGP completo

La aplicación admite un AGP con varios protocolos y varios AGP con sus respectivos protocolos.

**Si un proyecto tiene dos o más AGP, reúna todos en un único resumen de proyecto antes de cargarlo.** No cargue resúmenes separados para intentar validar por partes contra el mismo PMS completo.

Conserve la misma arquitectura del ejemplo `data_2000139288_2000139289.xlsx`:

```text
AGP 2000139288                         Nivel conceptual 1
  Centro responsable / Protocolo A    Nivel conceptual 2
    Unidad técnica                   Nivel conceptual 3
      CIT / ODM / graph              Nivel conceptual 4
        Material o prestación        Nivel conceptual 5
  Centro responsable / Protocolo B
    Unidad técnica
      CIT / ODM / graph
        Material o prestación
AGP 2000139289
  Centro responsable / Protocolo C
    Unidad técnica
      CIT / ODM / graph
        Material o prestación
```

Los niveles deben ser agrupaciones reales de Excel, no únicamente espacios, colores o sangrías visibles. El formato actual tiene detalle en el quinto nivel conceptual, normalmente `outline_level=4` porque el primer nivel usa 0. También se admite la estructura histórica con detalle en `outline_level=5`.

El AGP se reconoce por el código que empieza por `200`, por ejemplo `2000138097 - PREDUOCMT/2026_064/IX/346593`. El nombre del archivo cargado puede ser diferente; el nombre de salida se construye con los AGP detectados dentro del libro.

Utilice un solo resumen consolidado en una hoja. Si el libro contiene varias hojas con resúmenes, el lector selecciona la hoja candidata con más detalles; no suma automáticamente todas las hojas.

## 5. Ejecute la validación

1. Abra **Validar caso**.
2. En **Archivo original PMS**, cargue el PMS correspondiente.
3. En **Resumen exportado de AGP**, cargue el `data` completo del mismo caso.
4. Presione **Validar materiales y prestaciones** una sola vez y espere.
5. Revise el mensaje, los AGP y protocolos detectados y los conteos.
6. Descargue el Excel de validación, incluso cuando existan diferencias o errores de datos.
7. Archive la evidencia en la ubicación corporativa definida para el caso.

Si hay otra validación en curso, aparecerá un aviso. Espere un momento y vuelva a presionar **Validar**; no necesita volver a cargar los archivos. El piloto procesa un caso a la vez por servidor para limitar el consumo de memoria.

Use **Nuevo caso** para limpiar los archivos y el resultado de la pantalla. Al reemplazar un archivo se retira el resultado anterior, para evitar descargar un informe que corresponde a otra pareja de archivos.

## 6. Cómo se comparan los datos

Los códigos se tratan como texto, se pasan a mayúsculas y se eliminan espacios innecesarios. Los ceros iniciales se conservan si el Excel los almacena como texto. Si ya se perdieron en el archivo fuente, la herramienta no puede reconstruirlos.

Las cantidades admiten números y textos como `1`, `1.0`, `1,0` y `1,5`, además de negativos como `-1` o el formato AGP `1.000-`. En este último caso el valor es `-1`. Los errores de conversión se reportan, no se sustituyen silenciosamente por cero.

Se suman los registros por `TIPO` y `CODIGO`, y se calcula:

```text
DIFERENCIA = CANTIDAD_AGP - CANTIDAD_PMS
```

El margen de coincidencia es de `0,000001`. Las descripciones se muestran como contexto, pero no se comparan semánticamente. Si un código aparece repetido, sus cantidades se suman; la repetición por sí sola no es un error.

En proyectos con varios AGP o protocolos, la comparación es **consolidada para todo el caso**. Un faltante en un protocolo y un exceso del mismo código en otro pueden compensarse en el total. Revise la distribución técnica por separado; `AGP_Normalizado` conserva los AGP y protocolos asociados a cada código para facilitar esa revisión.

## 7. Materiales de aportación

**Los materiales de aportación no deben valorizarse.**

En RG11, el sistema excluye del PMS los materiales cuyos códigos cruzan con el maestro. En E1, excluye los clasificados como `Material de aportación` en el propio PMS. Todos quedan en `Materiales_Aportacion_Excluidos` con el motivo y la ubicación del registro fuente.

La exclusión se aplica al PMS valorizable. No se eliminan silenciosamente esos códigos de AGP: si un material de aportación fue valorizado allí, puede aparecer como `SOLO_EN_AGP` y requiere revisión.

En **Maestro de aportación** puede consultar y descargar la lista vigente. Solo el administrador autorizado de GitHub puede publicar cambios. El Excel de validación registra la huella SHA-256 del catálogo empleado, que permite identificar su versión exacta.

## 8. Interprete las alertas

### Colores de las tarjetas del resumen

Las tarjetas muestran el color y una etiqueta de estado para facilitar la revisión:

| Tarjeta | Color y significado |
| --- | --- |
| **Coinciden** | Verde cuando hay códigos con cantidades coincidentes. Si el conteo es cero, aparece azul con la etiqueta `Sin coincidencias`. |
| **Diferencias de cantidad** | Rojo si el conteo es mayor que cero: hay cantidades por corregir o justificar. Verde si es cero. |
| **Solo en PMS** | Rojo si el conteo es mayor que cero: hay códigos del PMS valorizable que faltan en AGP. Verde si es cero. |
| **Solo en AGP** | Rojo si el conteo es mayor que cero: hay códigos valorizados que no aparecen en el PMS valorizable. Verde si es cero. |
| **Errores** | Rojo si el conteo es mayor que cero: la validación está incompleta y debe revisar la hoja `Errores`. Verde si es cero. |
| **Códigos PMS / Códigos AGP** | Azul: totales de referencia; el color no indica por sí solo si el caso es correcto. |
| **Aportación excluida** | Azul: cantidad de códigos excluidos para auditoría. Una exclusión prevista no representa un error. |

**Verde significa coincidencia o ausencia de incidencias en ese indicador; rojo significa que requiere revisión; azul significa información de referencia o auditoría.** Las etiquetas `Coincidencia correcta`, `Sin incidencias` y `Requiere revisión` permiten interpretar el resultado sin depender únicamente del color.

Una tarjeta verde no aprueba el caso completo. Siempre prevalece el mensaje general de validación: si hay errores de lectura, los conteos pueden ser parciales aunque otras tarjetas aparezcan verdes.

### Estados de los registros

| Estado | Qué significa | Acción |
| --- | --- | --- |
| `OK` | El código existe en ambos archivos y coincide la cantidad. | Continúe con las demás revisiones técnicas del procedimiento. |
| `DIFERENCIA_CANTIDAD` | El código existe en ambos, pero cambia la cantidad. | Revise cantidades, signos y alcance antes de enviar. |
| `SOLO_EN_PMS` | Existe en el PMS valorizable y no aparece en AGP. | Revise si se omitió valorizar o si cambió el alcance. |
| `SOLO_EN_AGP` | Existe en AGP y no en el PMS valorizable. | Revise inclusiones, aportación y correspondencia del caso. |
| `ERROR_DATO` | Un registro no permite obtener una cantidad confiable. | Corrija la fuente y ejecute nuevamente. |

La diferencia positiva indica mayor cantidad en AGP; la negativa indica menor cantidad en AGP. Para códigos presentes solo en una fuente, la cantidad de la fuente ausente se deja vacía; la alerta de presencia es la referencia principal.

- **Validación OK:** no se encontraron diferencias ni errores de lectura en los datos comparados.
- **Validación con diferencias:** revise los registros pendientes y corrija antes de enviar.
- **Validación incompleta:** existen errores de lectura o estructura. No use los conteos parciales como confirmación de coincidencia; revise `Errores`.

## 9. Qué contiene el Excel

| Hoja | Contenido |
| --- | --- |
| `Resumen` | Fuentes, fecha/hora de Bogotá, AGP, protocolos, cantidades de códigos, resultado y versión del maestro. |
| `Comparacion` | Todos los códigos comparados, cantidades, diferencia, estado y observación. |
| `PMS_Normalizado` | PMS agregado después de excluir aportación. |
| `AGP_Normalizado` | AGP agregado, con referencias de proyectos y protocolos. |
| `Materiales_Aportacion_Excluidos` | Registros excluidos y motivo. |
| `Solo_PMS` | Códigos presentes únicamente en PMS valorizable. |
| `Solo_AGP` | Códigos presentes únicamente en AGP. |
| `Diferencias_Cantidad` | Códigos con cantidades diferentes. |
| `Errores` | Problemas detectados con fuente, hoja, fila y mensaje cuando están disponibles. |
| `Control Cambios` | Información del formato documental heredado; no representa firmas ni aprobación electrónica. |

El archivo conserva el formato Excel existente `F-AGP-07`, incluidos su encabezado y logo OCA. La versión web usa una presentación neutral de Dirección AGP. La adaptación documental a Bureau Veritas debe realizarse con el formato y los activos institucionales que se definan para el nuevo gestor.

Ejemplo de nombre para dos AGP:

`validacion_materiales_prestaciones_AGP_2000139288_2000139289.xlsx`

## 10. Problemas frecuentes

| Mensaje o situación | Qué revisar |
| --- | --- |
| No se detecta PMS. | Hojas `RG11_budget` o ambas `RG10_SER` y `RG36_MAT`; no basta con renombrar el archivo. |
| Falta una columna. | Encabezados reales y estructura del PMS; no cambie solo el título de una tabla diferente. |
| No se detecta AGP o la jerarquía. | Exportación original con agrupaciones Excel, código `200...` y todos los niveles. |
| Cantidad inválida. | Texto, errores de fórmula o separadores incorrectos en la fila indicada. |
| Maestro vacío o inválido. | El administrador debe corregirlo; RG11 no puede validarse sin un maestro válido. |
| Archivo demasiado extenso. | Revise el tamaño total y las filas/columnas sobrantes en las hojas de datos que se procesan. Las dimensiones de las hojas auxiliares del PMS no bloquean la validación. |
| Excel con contraseña. | Utilice una copia autorizada sin cifrado para la validación. |
| Página inactiva después de varias horas. | El alojamiento gratuito puede entrar en reposo. Espere su reactivación. |
| No abre desde la red corporativa. | Revise con TI el acceso a la dirección de la aplicación, carga y descarga. |

Los archivos `.xls` no son compatibles. Los libros `.xlsm` se leen como datos: no se ejecutan macros. Las fórmulas se leen con sus valores guardados; si estos no están actualizados, recalcule y guarde el Excel antes de cargarlo.

## 11. Manejo de archivos

Los Excel se envían al servidor que ejecuta la aplicación. Se utilizan temporales separados por validación y se eliminan al finalizar el procesamiento. El resultado permanece en la sesión para permitir la descarga; la aplicación no ofrece un histórico ni una copia automática en Drive.

No cierre o recargue la página antes de descargar: una desconexión, reinicio del servidor o pérdida de la sesión puede requerir repetir la validación. El archivado de la evidencia corresponde al usuario.

El repositorio público contiene código y códigos del maestro, no los archivos PMS/AGP cargados por la página. Las sesiones no comparten sus tablas ni resultados.

© 2026 Dirección AGP. Desarrollo: Juan Sebastián Fernández Buitrago.
