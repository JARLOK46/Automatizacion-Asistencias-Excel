# Automatización de Asistencias con Excel

Aplicación local para registrar asistencias desde un navegador, guardarlas en SQLite y proyectarlas en `Listado ejemplo.xlsx` sin reescribir innecesariamente el paquete XLSX.

## Características

- Formulario web local, accesible únicamente desde `127.0.0.1`.
- Validación de los datos antes de guardarlos.
- Control de documentos duplicados por campaña.
- Persistencia local en SQLite.
- Exportación automática al libro `Listado ejemplo.xlsx`.
- Copias de seguridad del libro antes de cada exportación.
- Registro operativo en `attendance.log`.
- Reintento de exportaciones pendientes o fallidas.
- Validación del candidato XLSX antes de reemplazar el archivo original.
- Preservación de hojas, dibujos y partes no relacionadas con la hoja `Formato`.

## Requisitos

- Python 3.10 o superior.
- Windows, Linux o macOS.
- El archivo `Listado ejemplo.xlsx` ubicado en la raíz del proyecto.

Las dependencias están definidas en `requirements.txt`:

```text
openpyxl==3.1.5
```

Aunque la exportación principal utiliza ZIP/XML para proteger la estructura del libro, `openpyxl` se utiliza para las comprobaciones de compatibilidad y lectura.

## Instalación

Cloná el repositorio y entrá en la carpeta del proyecto:

```bash
git clone https://github.com/JARLOK46/Automatizacion-Asistencias-Excel.git
cd Automatizacion-Asistencias-Excel
```

Creá y activá un entorno virtual:

### Windows PowerShell

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### Windows CMD

```bat
python -m venv .venv
.venv\Scripts\activate
```

### Linux/macOS

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Instalá las dependencias:

```bash
python -m pip install -r requirements.txt
```

## Ejecución del servidor

Desde la raíz del proyecto ejecutá:

```bash
python run_server.py
```

El servidor se inicia en:

```text
http://127.0.0.1:8000
```

Abrí esa dirección en el navegador. Para detenerlo, presioná `Ctrl+C` en la terminal.

El servidor utiliza por defecto:

- Base de datos: `attendance.db`
- Libro de trabajo: `Listado ejemplo.xlsx`
- Logs: `attendance.log`
- Backups: `backups/`

Estos archivos se crean localmente y no deben subirse al repositorio.

## Flujo de una inscripción

1. El usuario abre el formulario local.
2. Completa los campos de asistencia.
3. El servidor valida y normaliza los datos.
4. Se verifica que el documento no exista en la campaña.
5. El registro se guarda en SQLite con estado `pending`.
6. Se buscan todos los registros pendientes o fallidos, en orden de llegada.
7. Se crea un backup del libro original.
8. Se prepara un candidato temporal del XLSX.
9. El exporter modifica la hoja `Formato` desde la fila 17 y escribe las columnas A:N.
10. Se valida el XML y el paquete ZIP completo.
11. Solo si la validación termina correctamente se reemplaza el archivo original.
12. El registro se marca como exportado y el navegador muestra el número de llegada.

Si la exportación falla, el registro permanece en SQLite como `failed` y puede reintentarse en una solicitud posterior. El navegador informa que la asistencia fue guardada, pero que el Excel necesita revisión.

## Campos del formulario

El formulario registra:

- Nombre completo.
- Tipo de documento: `CC`, `TI` o `CE`.
- Número de documento.
- Género: `F`, `M` u `Otro`.
- Número de ficha de formación.
- Modalidad.
- Programa.
- Jornada/horario.
- Nivel educativo.
- Centro de formación.
- Correo electrónico.
- Teléfono.

El sistema también genera automáticamente datos técnicos como el identificador de envío, la fecha UTC y la fecha local.

## Exportación XLSX segura

El archivo `editar_excel.py` también puede utilizarse como script independiente para importar registros desde JSON.

### Entrada JSON

Creá un archivo, por ejemplo `datos.json`:

```json
{
  "arrival_seq": 1,
  "full_name": "Ada Lovelace",
  "document_type": "CC",
  "document_number": "123456",
  "gender": "F",
  "training_sheet_number": "456789",
  "modality": "Presencial",
  "program": "Python",
  "schedule": "Mañana",
  "education_level": "Técnico",
  "training_center": "Centro de formación",
  "email": "ada@example.com",
  "phone": "5551234",
  "registered_at_local": "01/01/2026 10:00"
}
```

También se admite un arreglo JSON con varios registros.

### Ejecución

```bash
python editar_excel.py --input datos.json
```

El script:

- lee el JSON;
- valida que estén todos los campos;
- verifica que no haya caracteres XML inválidos;
- crea un backup en `backups/`;
- modifica el libro raíz;
- informa la ruta del backup generado.

### Dry run

Para validar la entrada y preparar la transformación sin modificar el libro:

```bash
python editar_excel.py --input datos.json --dry-run
```

## Protección contra corrupción del Excel

El exporter trabaja con un archivo temporal y no toca el original hasta completar las comprobaciones. Antes del reemplazo verifica:

- que el archivo sea un ZIP XLSX válido;
- que no existan entradas ZIP duplicadas;
- que exista la hoja `Formato`;
- que el XML tenga un elemento raíz `worksheet` válido;
- que los namespaces de Excel requeridos estén presentes;
- que exista `sheetData`;
- que las partes no relacionadas con la hoja editada no hayan cambiado;
- que el candidato pueda volver a abrirse como XLSX.

Si alguna comprobación falla, el original permanece sin cambios.

La primera fila de datos de asistencia es la fila **17**. La exportación escribe las columnas **A:N** y conserva las partes no editadas del paquete, incluyendo dibujos y otras hojas.

## Logs y diagnóstico

El archivo `attendance.log` registra:

- inicio de solicitudes;
- validaciones rechazadas;
- documentos duplicados;
- guardados en la base de datos;
- intentos de exportación;
- filas exportadas;
- errores de exportación;
- excepciones inesperadas con traceback.

No se registran valores completos sensibles del formulario.

Cuando una exportación falla, el sistema puede generar un informe diagnóstico dentro de `backups/`, con extensión `.diagnostic.txt`.

## Recuperación ante un fallo

1. No borres el backup más reciente.
2. Cerrá Excel si el archivo está abierto.
3. Revisá `attendance.log`.
4. Revisá el archivo `.diagnostic.txt` dentro de `backups/`.
5. Compará el libro principal con el backup más reciente.
6. Corregí el problema y realizá una nueva solicitud desde el formulario.

No uses un backup como archivo de trabajo mientras el servidor esté ejecutándose, porque el programa siempre utiliza el archivo raíz `Listado ejemplo.xlsx`.

## Estructura principal

```text
.
├── Listado ejemplo.xlsx       # Plantilla de Excel utilizada por la aplicación
├── run_server.py              # Punto de entrada del servidor web
├── editar_excel.py            # Exportador XLSX independiente
├── requirements.txt            # Dependencias Python
└── attendance_app/
    ├── config.py              # Rutas y configuración
    ├── db.py                  # SQLite y repositorio de asistencias
    ├── logging_setup.py        # Configuración de attendance.log
    ├── service.py              # Exportación y reintentos
    ├── validation.py           # Validación de entradas
    └── web.py                 # Formulario y servidor HTTP local
```

## Limitaciones actuales

- El servidor está diseñado para uso local y no incluye autenticación.
- No debe exponerse directamente a Internet o a una red LAN sin una capa adicional de seguridad.
- El libro debe conservar una hoja llamada `Formato`.
- El nombre y la estructura de las columnas A:N deben seguir la plantilla incluida.

## Licencia

No se especificó una licencia para este proyecto. Consultá con el propietario antes de reutilizarlo o redistribuirlo.
