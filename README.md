# Descarga Automática de CFDI

Herramientas en Python para un despacho contable:

| Script | Qué hace |
|---|---|
| `cfdi_sat.py` | Descarga los CFDI (emitidos y recibidos) del portal del SAT para varios clientes y los guarda en carpetas ordenadas. |
| `guardar_credencial.py` | Guarda la CIEC de un cliente en el Administrador de credenciales de Windows. |
| `auto_clasificador.py` | Clasifica gastos ambiguos de un estado de cuenta en el catálogo contable usando IA (Azure OpenAI o Google Vertex AI). Se puede usar solo o desde `conciliador.py`. |
| `conciliador.py` | Lee los XML descargados, los concilia contra el estado de cuenta bancario y (opcional) clasifica con IA. Genera un Excel. |

## Requisitos

- Windows, Python 3.10 o superior
- Para la descarga: Chromium de Playwright
- Para el clasificador: acceso a Azure OpenAI o a Vertex AI

## Instalación

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
playwright install chromium
```

## 1. Descarga de CFDI

### Configuración inicial

1. **Lista de clientes.** Edita `clientes.csv` con las columnas `rfc,nombre`:

   ```csv
   rfc,nombre
   XAXX010101000,Cliente Ejemplo SA
   ```

   El RFC se valida (12 caracteres para personas morales, 13 para personas físicas); los inválidos se omiten con un error en el log.

2. **Credenciales.** Guarda la CIEC de cada cliente. Se almacena en el Administrador de credenciales de Windows, nunca en archivos:

   ```powershell
   python guardar_credencial.py XAXX010101000
   ```

3. **Destino.** Por defecto es `G:\Mi unidad\CFDI` (carpeta de Drive for Desktop, que sincroniza solo). Cámbialo con `--destino` o con la variable de entorno `CFDI_DESTINO`.

### Uso

```powershell
python cfdi_sat.py --anio 2026 --mes 9                       # todos los clientes, emitidos y recibidos
python cfdi_sat.py --anio 2026 --mes 9 --rfc XAXX010101000   # un solo cliente
python cfdi_sat.py --anio 2026 --mes 9 --tipo recibidos      # solo recibidos
python cfdi_sat.py --anio 2026 --mes 9 --destino D:\CFDI     # otro destino
```

| Opción | Descripción |
|---|---|
| `--anio` | Año (obligatorio) |
| `--mes` | Mes 1-12 (obligatorio) |
| `--rfc` | Procesar solo este RFC |
| `--tipo` | `emitidos`, `recibidos` o `ambos` (default) |
| `--destino` | Carpeta raíz de destino |
| `--clientes` | Ruta a otro CSV de clientes |

### Cómo funciona

- Abre un Chromium **visible**: por cada cliente guarda el captcha en `captcha.png` y te pide escribirlo en la consola (3 intentos).
- Estructura de salida:

  ```
  <DESTINO>/<RFC - Nombre>/<AAAA>/<MM>/<Emitidos|Recibidos>/<UUID>.xml
  ```

- **Se puede repetir sin duplicar:** si el `<UUID>.xml` ya existe, lo salta.
- **Descargas seguras:** cada XML se baja a un `.part`, se valida y solo entonces se renombra; hasta 3 intentos por archivo. Un fallo no detiene el resto.
- **Límite del SAT:** el portal trunca a 500 resultados por búsqueda. Si se alcanza, el script repite la búsqueda día por día.
- **Diagnóstico:** `cfdi.log` (con rotación) y, si falla un cliente, una captura en `errores/<RFC>.png`.
- **Código de salida:** `0` si todo salió bien, `1` si hubo clientes o descargas con error (útil para programarlo).

### Si el SAT cambia el portal

Todos los selectores están en el diccionario `SEL` al inicio de `cfdi_sat.py`; solo hay que ajustarlos ahí.

## 2. Clasificador de gastos con IA

Dado un concepto y un monto, devuelve un JSON con `cuenta_sugerida`, `codigo_cuenta`, `nivel_confianza` (Alto/Medio/Bajo) y `justificacion_fiscal`. El prompt asume contabilidad mexicana y el código agrupador del SAT.

### Variables de entorno

| Proveedor | Variables |
|---|---|
| Azure OpenAI | `AZURE_OPENAI_API_KEY`, `AZURE_OPENAI_ENDPOINT`, `AZURE_OPENAI_DEPLOYMENT`, opcional `AZURE_OPENAI_API_VERSION` |
| Google Vertex AI | `GOOGLE_APPLICATION_CREDENTIALS`, `GOOGLE_CLOUD_PROJECT`, opcional `GOOGLE_CLOUD_LOCATION` (default `us-central1`) y `GCP_MODEL` (default `gemini-2.5-flash`) |
| Cifrado local (opcional) | `CLASIFICADOR_FERNET_KEY` |

Las claves nunca van en el código.

### Uso

```powershell
python auto_clasificador.py --provider azure --concepto "AMAZON MKTPLACE" --monto 1250.50
python auto_clasificador.py --provider gcp --input-file movimientos.xlsx
python auto_clasificador.py          # modo interactivo
```

Un archivo de entrada (CSV o Excel) debe tener las columnas `concepto` y `monto`. El resultado se guarda junto al original como `<archivo>_clasificado.csv`, de forma incremental, así que si se interrumpe no se pierde el avance. Opciones: `--pausa N` (espera entre llamadas) y `-v` (log detallado).

Salida de ejemplo:

```json
{
  "cuenta_sugerida": "Gastos generales",
  "codigo_cuenta": "601.84",
  "nivel_confianza": "Medio",
  "justificacion_fiscal": "Compra en marketplace sin detalle; requiere CFDI para deducir."
}
```

Códigos de salida: `0` ok, `2` terminó con filas fallidas, `1` error, `130` cancelado.

> **Nota:** `cryptography.fernet` usa AES-128-CBC con HMAC, no AES-256.

## 3. Conciliación bancaria (`conciliador.py`)

Flujo mensual completo: `cfdi_sat.py` descarga → `conciliador.py` lee, concilia y clasifica → Excel.

```powershell
python conciliador.py --rfc XAXX010101000 --anio 2026 --mes 9 --estado-cuenta banco.xlsx
python conciliador.py --rfc XAXX010101000 --anio 2026 --mes 9 --estado-cuenta banco.xlsx --provider azure
python conciliador.py --rfc XAXX010101000 --anio 2026 --mes 9      # solo leer/clasificar los CFDI
```

| Opción | Descripción |
|---|---|
| `--rfc`, `--anio`, `--mes` | Cliente y periodo (obligatorios) |
| `--estado-cuenta` | CSV/Excel del banco. Sin él solo se leen y clasifican los CFDI |
| `--destino` | Carpeta donde `cfdi_sat.py` guardó los XML |
| `--provider` | `azure` o `gcp`: activa la clasificación con IA |
| `--tolerancia` | Diferencia máxima de monto en MXN (default 0.01) |
| `--ventana-dias` | Días máximos entre la fecha del CFDI y el pago (default 45) |
| `--meses-atras` | Meses previos de CFDI a considerar, porque un pago puede cubrir facturas anteriores (default 1) |
| `--cargos-positivos` | Si en la columna `monto` los cargos vienen en positivo |
| `--salida` | Ruta del Excel (default `conciliacion_<RFC>_<AAAA>-<MM>.xlsx`) |

**Estado de cuenta:** columnas `fecha`, `concepto` (o `descripcion`) y `monto` con los cargos en negativo, o bien `cargo` y `abono` por separado. Fechas `AAAA-MM-DD` o `DD/MM/AAAA`. Los movimientos fuera del mes pedido se ignoran.

**Cómo concilia**
- Lee CFDI 3.3 y 4.0 (facturas, notas de crédito, nómina y complementos de pago). Convierte a MXN las facturas en otra moneda con su tipo de cambio.
- Empareja cada movimiento con un CFDI por monto, signo (cargo/abono según sea emitido o recibido) y fecha. El puntaje sube si el RFC, el nombre o el folio aparecen en el concepto bancario.
- También detecta un pago que cubre 2 o 3 facturas del mismo proveedor o cliente.
- **Facturas PPD:** no cruzan directo con el banco; se concilian con su complemento de pago (tipo P) y aparecen en "CFDI sin mov" con una nota.
- Cada par lleva confianza Alto/Medio/Bajo. Si hay varios candidatos parecidos, baja un nivel. **Revisa siempre los Bajo.**

**Clasificación (con `--provider`)**: se clasifican los CFDI *recibidos* de tipo ingreso/egreso (usando proveedor, conceptos, uso de CFDI y clave de producto SAT) y los cargos del banco sin CFDI. Los emitidos y los abonos no se clasifican. Los CFDI repetidos de un mismo proveedor se clasifican una sola vez.

**Hojas del Excel:** `Resumen`, `Conciliados`, `Mov sin CFDI` (el trabajo pendiente: cargos sin comprobante), `CFDI sin mov` y `CFDI` (detalle completo con clasificación).

**Límites conocidos**
- No verifica si un CFDI fue cancelado ante el SAT: el XML no lo indica.
- No concilia pagos parciales de una misma factura ni un movimiento que cubra más de 3 facturas.
- Solo mira CFDI de los meses ya descargados; un CFDI de un mes posterior al pedido no se considera.

## Seguridad

- Las CIEC viven solo en el Administrador de credenciales de Windows.
- Las claves de IA se leen de variables de entorno.
- `clientes.csv` contiene RFC reales y está versionado en git: no lo subas a un repositorio público.
- Los conceptos bancarios se envían a Azure o Google; revisa que tu política de datos lo permita.
