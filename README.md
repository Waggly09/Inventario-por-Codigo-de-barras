# Lector de Codigo de barras

Aplicación de escritorio para Windows que permite identificar productos y cajas con un lector USB en modo teclado. El lector envía Enter al terminar cada lectura. La aplicación conserva ceros iniciales y muestra todas las coincidencias cuando un código se comparte entre varios registros.

## Uso

1. Abre `Corrugados.exe`.
2. Ve a **Datos y respaldos → Importar Excel** y elige el archivo con la columna **NUMERO BASE + NUMERO MOVIL**. Revisa la vista previa y confirma la importación.
3. Presiona **F2** y escanea o escribe el código completo. Revisa la ficha antes de usar el producto.
4. Presiona **F3** para consultar, añadir, editar o eliminar registros.

Una instalación nueva comienza con el catálogo vacío. Los cambios se guardan en `%LOCALAPPDATA%\Corrugados`. Usa **Exportar catálogo** y **Recibir catálogo** para transferir una copia completa a otra computadora. Recibir una copia reemplaza el catálogo local y guarda un respaldo previo.

## Compilar en Windows

Requiere Python 3.12 de 64 bits con Tcl/Tk. Desde PowerShell, en la carpeta del proyecto:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m unittest discover -p 'test_*.py'
.\.venv\Scripts\python.exe build.py
```

El ejecutable queda en `dist\Corrugados.exe` y contiene su propio entorno de ejecución. Para instalarlo por usuario, abre el programa y usa **Datos y respaldos → Instalar en este equipo**.

Para ejecutar el código en un catálogo de desarrollo separado:

```powershell
.\.venv\Scripts\python.exe app.py --data-dir .\datos-desarrollo
```
