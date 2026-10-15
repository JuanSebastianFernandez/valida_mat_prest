# Administración del piloto y del maestro

## Modelo de administración

El maestro es `catalogo/materiales_aportacion.csv`. Se distribuye con la aplicación y se aplica automáticamente a RG11. E1 utiliza su columna `TIPO DE MOVIMIENTO`.

En este piloto no hay una pestaña para editar el maestro desde la web. El administrador actualiza su copia local y publica el CSV en GitHub. Esto evita introducir contraseñas o credenciales de escritura en una aplicación pública.

El repositorio público permite leer el maestro, pero solo las cuentas autorizadas con escritura pueden cambiar su versión oficial. Mantenga esos permisos bajo su administración; no depende de que los demás usuarios desconozcan GitHub.

## 1. Preparación inicial

Instale Python 3.12 y descargue o clone la carpeta publicada. Abra PowerShell dentro de ella y ejecute:

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
```

Los siguientes ejemplos usan el Python de ese entorno, por lo que no es necesario activarlo.

## 2. Consultar el maestro

```powershell
.venv\Scripts\python scripts/actualizar_maestro.py listar
```

Se imprimen el total, la huella SHA-256 y los códigos. La misma lista puede consultarse en **Maestro de aportación** en la aplicación.

## 3. Importar o reemplazar el maestro completo

Use `importar` solo cuando quiera que el archivo suministrado sea la lista completa vigente. Puede retirar materiales que estaban antes y ya no figuran en la nueva lista.

Primero vea el efecto:

```powershell
.venv\Scripts\python scripts/actualizar_maestro.py importar --archivo "C:\ruta\MAT APORT_VF_1.xlsx"
```

Después aplíquelo:

```powershell
.venv\Scripts\python scripts/actualizar_maestro.py importar --archivo "C:\ruta\MAT APORT_VF_1.xlsx" --guardar
```

Los encabezados admitidos para encontrar la columna son `MAT APORT`, `CODIGO MATERIAL`, `CODIGO`, `MATERIAL` o `UBICACION`. Se busca en las primeras 30 filas de las hojas. Se prefiere `MAT APORT` si hay varias candidatas.

El catálogo inicial de este paquete se obtiene del archivo base `MAT APORT_VF_1.xlsx` disponible en el equipo. Verifique su vigencia funcional antes de publicarlo.

## 4. Agregar materiales masivamente sin duplicados

Para añadir una lista manteniendo los materiales que ya existen, use `agregar`:

```powershell
.venv\Scripts\python scripts/actualizar_maestro.py agregar --archivo "C:\ruta\nuevos_materiales.xlsx"
.venv\Scripts\python scripts/actualizar_maestro.py agregar --archivo "C:\ruta\nuevos_materiales.xlsx" --guardar
```

También puede utilizar `.csv` con encabezado `CODIGO` o `.txt` con un código por línea. Si usa TXT, también se admiten espacios, comas o punto y coma como separadores.

Para pocos códigos:

```powershell
.venv\Scripts\python scripts/actualizar_maestro.py agregar --codigos "001234 120570 970698"
.venv\Scripts\python scripts/actualizar_maestro.py agregar --codigos "001234 120570 970698" --guardar
```

El comando informa recibidos, repetidos dentro de la carga, ya existentes, agregados, retirados y total final. Cada código queda una sola vez. Un código `001234` se conserva distinto de `1234` si la fuente mantiene esos ceros como texto.

Los códigos de material deben contener solo dígitos después de normalizar. Si encuentra un dato inválido se rechaza la actualización completa y no se aplica una carga parcial.

## 5. Retirar materiales

```powershell
.venv\Scripts\python scripts/actualizar_maestro.py retirar --codigos "120570 970698"
.venv\Scripts\python scripts/actualizar_maestro.py retirar --codigos "120570 970698" --guardar
```

`retirar` también admite `--archivo` para listas completas de bajas. No produce error si alguno ya no existe.

**Los códigos de los ejemplos no son una instrucción de clasificación real.** Agregue o retire únicamente los aprobados para el maestro de aportación.

## 6. Publicar el cambio en GitHub

Guardar localmente no actualiza por sí solo la página desplegada.

### Opción sencilla: desde la web de GitHub

1. Abra su repositorio y entre a `catalogo`.
2. Use **Add file > Upload files**.
3. Cargue su `materiales_aportacion.csv` actualizado para reemplazar el del mismo nombre.
4. Escriba un mensaje descriptivo, por ejemplo: `Actualiza aportacion: altas y bajas revisadas`.
5. Confirme el cambio en la rama desplegada.
6. Espere la actualización de Streamlit. Si no se refleja, reinicie la aplicación desde su panel.
7. Abra **Maestro de aportación** y compruebe el total y la huella de versión.

No suba el Excel original, la carpeta `.respaldos` ni los archivos de casos.

### Opción por consola: si clonó el repositorio con Git

```powershell
git pull --ff-only
```

Haga la actualización local del maestro después de sincronizar. Luego publique solo el CSV:

```powershell
git add catalogo/materiales_aportacion.csv
git commit -m "Actualiza maestro de materiales de aportacion"
git push
```

Git puede solicitar autenticación a su cuenta. La aplicación no necesita esa credencial.

## 7. Versiones y recuperación

Antes de reemplazar un CSV local, el comando conserva una copia en `catalogo/.respaldos`, identificada por el SHA-256 de su contenido. Esa carpeta se excluye de la publicación.

GitHub conserva el historial de los cambios publicados. Para recuperar una lista anterior, obtenga el CSV de la versión correcta, revíselo y publíquelo como un nuevo cambio. No necesita modificar el código Python.

Cada Excel generado registra la huella SHA-256 completa del maestro usado en RG11. En E1 se registra que no aplica. También se registran las huellas de los archivos de entrada para ayudar a identificar la evidencia.

Si cambia el maestro mientras existe un resultado abierto, la aplicación solicita volver a validar al detectar la nueva versión. Un resultado descargado anteriormente sigue siendo evidencia de la versión con la que fue generado.

## 8. Operación y límites

El piloto procesa un caso a la vez por proceso de servidor. Los demás usuarios pueden navegar y cargar, pero reciben un aviso si intentan validar mientras está ocupado. No hay una cola automática ni un segundo plano persistente.

Los temporales de cada validación tienen carpetas independientes y se eliminan al terminar. Los informes se mantienen solo en la sesión del usuario para la descarga. No hay copia automática de casos en GitHub ni Drive.

La capacidad final depende del tamaño de los Excel y de los recursos del alojamiento. Los 20 usuarios pueden compartir la herramienta, pero este paquete no garantiza 20 procesos simultáneos ni un nivel de disponibilidad. Evalúe la operación con casos reales y considere alojamiento con recursos asignados si los reintentos se vuelven frecuentes.

Para soporte recopile el mensaje visible, fecha/hora, tamaño de los archivos, tipo PMS y versión del maestro. No pegue contenido completo de casos en registros públicos.

## 9. Formato corporativo

El Excel mantiene F-AGP-07, Rev.01, el encabezado OCA y su control documental histórico. La interfaz del piloto es neutral de Dirección AGP. Para adoptar el formato bajo Bureau Veritas se necesitan el encabezado, logo y datos documentales definidos por el nuevo gestor; no se han sustituido por una identificación inventada.

El notebook original permanece disponible como referencia. El módulo `.py` incorpora ajustes compatibles para el servidor web, pero la copia de código embebida en el `.ipynb` no se modifica con este piloto.
