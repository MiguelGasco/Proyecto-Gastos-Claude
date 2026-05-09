# Control de gastos e ingresos mensuales en Excel — Diseño

**Fecha:** 2026-05-09
**Estado:** propuesta validada, pendiente de plan de implementación
**Usuario:** Miguel (miguel.gasco@appyweb.es)

## 1. Objetivo

Llevar un registro persistente de los ingresos y gastos personales de Miguel en un único archivo Excel (`gastos.xlsx`) con KPIs y gráficas vivas, alimentado conversacionalmente: Miguel describe los movimientos en lenguaje natural ("hoy he gastado 30 € comiendo con amigos") y Claude los registra automáticamente en el archivo y en la memoria persistente Engram.

El objetivo final es que, al abrir el Excel, Miguel vea de un vistazo en qué se le va el dinero cada mes y cómo evoluciona a lo largo del año.

## 2. Modo de uso (flujo del usuario)

1. Miguel le dice a Claude un movimiento por chat, en lenguaje natural.
2. Claude interpreta `fecha`, `tipo` (Ingreso/Gasto), `categoría`, `importe` y `descripción`.
   - Si hay ambigüedad (categoría dudosa, importe o fecha poco clara), Claude pregunta antes de registrar.
   - Si está claro, Claude registra directamente sin pedir confirmación.
3. Claude ejecuta el script `add_movement.py` que añade una fila a la tabla del `.xlsx`.
4. Claude además guarda el movimiento en Engram (`mem_save` tipo `transaction`) para tener histórico consultable entre sesiones.
5. Cuando Miguel abre `gastos.xlsx`, Excel recalcula y las gráficas/KPIs reflejan los datos actuales.

Reglas adicionales:

- **Lote**: si Miguel manda varios movimientos seguidos, Claude los procesa uno a uno y al terminar muestra un resumen breve.
- **Rectificación**: errores se corrigen con `add_movement.py --edit-last` (modificar la última fila) o `--delete-last` (eliminarla).
- **Fechas relativas**: "ayer", "el lunes", "3/5", etc. se traducen a `YYYY-MM-DD`.
- **Sincronización Excel ↔ Engram**: el script no se considera exitoso hasta que ambos quedan grabados.

## 3. Estructura de archivos

```
Proyecto Gastos Claude/
├── gastos.xlsx              ← fuente única de datos (creado una sola vez)
├── scripts/
│   ├── init_workbook.py     ← crea gastos.xlsx desde cero
│   └── add_movement.py      ← añade / edita / borra filas
├── requirements.txt         ← openpyxl
└── README.md                ← instrucciones de uso para Miguel
```

## 4. Estructura del Excel (`gastos.xlsx`)

### 4.1 Hoja `Movimientos`

Tabla nativa de Excel llamada `tblMov`. Es la única hoja que crece con el tiempo y la única fuente de datos.

| Columna | Tipo | Notas |
|---|---|---|
| Fecha | Fecha (`dd/mm/yyyy`) | Obligatoria |
| Tipo | Texto (`Ingreso` o `Gasto`) | Validación de lista |
| Categoría | Texto | Validación contra hoja `Categorías` (dependiente de `Tipo`) |
| Importe | Número (€, formato `#.##0,00 €`) | Siempre positivo |
| Descripción | Texto libre | Opcional |

### 4.2 Hoja `Categorías`

Lista de referencia, dos columnas (`Tipo`, `Categoría`):

**Gastos:** Alimentación, Restauración, Vivienda, Suministros, Transporte, Salud, Ocio, Compras, Suscripciones, Otros.
**Ingresos:** Nómina, Extras, Otros.

Editable a mano si en el futuro se quieren añadir categorías; los desplegables de `Movimientos` la leen.

### 4.3 Hoja `Dashboard`

Lo que Miguel ve al abrir el archivo.

- **Selector de mes** (celda con desplegable): por defecto `=TEXTO(HOY();"mm/yyyy")`.
- **4 KPIs** en cajas grandes arriba:
  - Ingresos del mes seleccionado
  - Gastos del mes seleccionado
  - Ahorro neto = Ingresos − Gastos
  - % ahorro = Ahorro / Ingresos
- **Gráfico de tarta**: Gastos por categoría del mes seleccionado.
- **Gráfico de líneas**: Evolución últimos 12 meses (3 series: Ingresos, Gastos, Ahorro).
- **Tabla "Top 5 gastos del mes"**: las 5 filas con mayor importe del tipo `Gasto` en el mes seleccionado, mostrando Fecha, Descripción e Importe.

### 4.4 Hoja `_aux` (oculta)

Cálculos intermedios que alimentan las gráficas:

- **Resumen mensual**: 12 filas con `Año-Mes`, `Ingresos`, `Gastos`, `Ahorro` calculados con `SUMIFS` sobre `tblMov`.
- **Resumen por categoría del mes seleccionado**: 10 filas (una por categoría de gasto) con `SUMIFS` filtrando por mes seleccionado.
- **Top 5 del mes**: cálculo con `LARGE` + `INDEX/MATCH` o `FILTER` sobre `tblMov`.

Las gráficas del Dashboard apuntan a estas tablas auxiliares; los KPIs leen directamente con `SUMIFS`.

## 5. Componentes técnicos

### 5.1 `init_workbook.py`

Ejecutado una sola vez al arrancar el proyecto. Usa `openpyxl` para crear `gastos.xlsx` con:

- Las 4 hojas (`Movimientos`, `Categorías`, `Dashboard`, `_aux`).
- La tabla nativa `tblMov` vacía (con cabeceras y formato).
- Las validaciones de datos en `Tipo` y `Categoría`.
- La lista de categorías inicial.
- Las fórmulas del Dashboard y `_aux`.
- Los gráficos (tarta y líneas) apuntando a `_aux`.
- La hoja `_aux` oculta.
- Formatos de moneda (€) y fecha (`dd/mm/yyyy`) configurados con locale español.

Es idempotente: si `gastos.xlsx` ya existe, falla con mensaje claro en vez de sobrescribir.

### 5.2 `add_movement.py`

Script reutilizable invocado por Claude por cada movimiento. Argumentos:

```
add_movement.py --fecha YYYY-MM-DD --tipo {Ingreso|Gasto} \
                --categoria <Categoría> --importe <número> \
                [--descripcion "<texto>"]

add_movement.py --edit-last [mismos flags para los campos a cambiar]
add_movement.py --delete-last
```

Comportamiento:

1. Abre `gastos.xlsx` con `openpyxl`.
2. Valida que la categoría existe en la hoja `Categorías` para el `tipo` indicado; si no, falla.
3. Añade/edita/borra fila en `tblMov` (la tabla se extiende sola al añadir).
4. Guarda y cierra.
5. Sale con código 0 si todo OK; código ≠ 0 con mensaje claro en errores.

El guardado en Engram lo hace Claude tras una salida exitosa del script, no el script en sí.

## 6. Categorización automática (responsabilidad de Claude)

Claude es responsable de mapear lenguaje natural → categoría cerrada. Reglas básicas:

- "súper", "compra del mercado", "Mercadona/Lidl/etc" → **Alimentación**
- "comer/cenar fuera", "café", "bar", restaurante con nombre → **Restauración**
- "alquiler", "hipoteca", "comunidad", "seguro hogar" → **Vivienda**
- "luz", "agua", "gas", "internet", "factura del móvil" → **Suministros**
- "gasolina", "metro", "tren", "taxi", "Uber", "ITV", "taller" → **Transporte**
- "médico", "farmacia", "dentista", "gimnasio" → **Salud**
- "cine", "viaje", "hobby", "regalo" → **Ocio**
- "ropa", "móvil nuevo", "electrónica", "hogar" (no esencial) → **Compras**
- "Netflix", "Spotify", suscripciones SaaS → **Suscripciones**
- Solo si no encaja en nada → **Otros**

Si una descripción podría caer en dos categorías razonables, Claude pregunta antes de registrar.

## 7. Persistencia en Engram

Cada movimiento exitoso se guarda con `mem_save`:

- `type`: `transaction`
- `topic_key`: `movimiento-YYYY-MM-DD-<n>`
- `content`: resumen estructurado del movimiento (fecha, tipo, categoría, importe, descripción)

Esto permite recuperar el histórico aunque el `.xlsx` no esté disponible y mantener contexto entre sesiones.

## 8. Fuera de alcance (v1)

- Saldo bancario / balance acumulado.
- Presupuestos por categoría con alertas.
- Múltiples cuentas / métodos de pago.
- Importación automática de extractos bancarios.
- Movimientos recurrentes autogenerados.
- Interfaz móvil / web propia (se puede usar Excel móvil sobre OneDrive si hace falta).

Cualquiera de estas se puede añadir como v2 sin romper la estructura propuesta.

## 9. Criterios de éxito

1. Miguel dice "gasté 30€ comiendo" → aparece la fila en `Movimientos` y al abrir el Excel los KPIs reflejan el cambio.
2. El Dashboard muestra los 4 KPIs, la tarta de categorías, la línea de evolución mensual y el top 5, todos del mes seleccionado.
3. Cambiar el selector de mes recalcula KPIs, tarta y top 5 sin tocar nada más.
4. Engram contiene cada movimiento como observación y se puede buscar por mes/categoría.
5. Miguel puede editar/borrar el último movimiento sin abrir el Excel a mano.
