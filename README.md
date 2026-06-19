# Script de Limpieza de Dataset de Ventas

## Tabla de contenidos

- [Descripción general](#descripción-general)
- [Requisitos](#requisitos)
- [Estructura del script](#estructura-del-script)
  - [1. Carga del dataset](#1-carga-del-dataset)
  - [2. Eliminación de filas duplicadas](#2-eliminación-de-filas-duplicadas)
  - [3. Estandarización de medios de pago](#3-estandarización-de-medios-de-pago)
  - [4. Estandarización de sucursales](#4-estandarización-de-sucursales)
  - [5. Saneamiento de fechas inválidas](#5-saneamiento-de-fechas-inválidas)
  - [6. Saneamiento de columnas numéricas](#6-saneamiento-de-columnas-numéricas)
  - [7. Estandarización de columnas categóricas restantes](#7-estandarización-de-columnas-categóricas-restantes)
  - [8. Resumen final y exportación](#8-resumen-final-y-exportación)
- [Salida del script](#salida-del-script)
- [Notas y posibles mejoras](#notas-y-posibles-mejoras)

---

## Descripción general

El script toma un dataset de ventas en Excel con errores típicos de calidad de datos (duplicados, texto mal escrito, tildes inconsistentes, mayúsculas/minúsculas mezcladas, valores fuera de rango, fechas inválidas, etc.) y aplica una serie de correcciones column por columna, imprimiendo en consola un reporte detallado de cada paso. Al final, genera un nuevo archivo Excel con los datos ya normalizados.

## Requisitos

```bash
pip install pandas numpy openpyxl
```

- `pandas` y `numpy`: manipulación de datos.
- `openpyxl`: motor necesario para que pandas pueda leer/escribir archivos `.xlsx`.
- El archivo de entrada `dataset_ventas_tp_5000_lineas_calidad_datos.xlsx` debe estar en el mismo directorio que el script.

## Estructura del script

### 1. Carga del dataset

```python
df = pd.read_excel("dataset_ventas_tp_5000_lineas_calidad_datos.xlsx")
```

- Se carga el archivo Excel original en un `DataFrame` (`df`).
- Se imprime la cantidad de filas y columnas detectadas.
- Se crean **columnas auxiliares numéricas/fecha** (`precio_num`, `cant_num`, `total_num`, `desc_num`, `fecha_dt`) usando `pd.to_numeric` y `pd.to_datetime` con `errors='coerce'`. Esto convierte cualquier valor no numérico/no fecha en `NaN`/`NaT`, lo que permite hacer un análisis exploratorio rápido de la calidad de los datos antes de limpiar.
- Inmediatamente después, se vuelve a leer el archivo en una **segunda copia independiente** (`df_clean`), que es la que efectivamente se irá modificando a lo largo del script. Esto separa el dataset "de diagnóstico" (`df`) del dataset "de trabajo" (`df_clean`).

> **Nota:** las columnas auxiliares creadas sobre `df` (`precio_num`, `cant_num`, etc.) no se usan en el resto del flujo de limpieza, ya que las correcciones reales se aplican sobre `df_clean`.

### 2. Eliminación de filas duplicadas

```python
df_clean = df_clean.drop_duplicates(keep='first')
```

- **Qué hace:** elimina filas que son **idénticas en todas sus columnas** (incluyendo `id_venta`), dejando solo la primera ocurrencia.
- **Por qué:** los registros 100% duplicados no aportan información nueva y pueden distorsionar métricas (sumas, conteos, promedios).
- Se calcula y muestra cuántas filas había antes, cuántas se eliminaron y cuántas quedaron, y luego se reinicia el índice del `DataFrame` con `reset_index(drop=True)` para mantener un índice consecutivo y limpio.

### 3. Estandarización de medios de pago

Esta sección unifica todas las variantes de texto de la columna `medio_pago` en un conjunto cerrado de valores válidos.

```python
MEDIOS_VALIDOS = {
    'efectivo', 'transferencia', 'tarjeta credito',
    'tarjeta debito', 'mercado pago', 'cheque'
}
```

- **`quitar_tildes(texto)`:** función auxiliar que usa `unicodedata.normalize('NFD', texto)` para descomponer caracteres acentuados en su letra base + el carácter de acento (categoría Unicode `Mn`, *Mark, nonspacing*), y luego filtra esos acentos. Así, `"crédito"` se convierte en `"credito"`. Esto evita tener que escribir variantes con y sin tilde para cada valor.
- **`MAPA_MEDIOS`:** diccionario que traduce variantes conocidas y mal escritas (por ejemplo `"transf"`, `"mercadopago"`) a su forma canónica (por ejemplo `"transferencia"`, `"mercado pago"`).
- **`normalizar_medio_pago(valor)`:** función que se aplica fila por fila (`.apply`) y sigue esta lógica:
  1. Si el valor es nulo (`NaN`) o una cadena vacía tras quitar espacios → devuelve `''` (cadena vacía, que se trata como "nulo semántico").
  2. Convierte el texto a minúsculas, sin espacios al inicio/fin y sin tildes.
  3. Busca esa forma normalizada en `MAPA_MEDIOS`; si no está, usa el valor normalizado tal cual.
  4. Si el resultado final **no** pertenece al conjunto `MEDIOS_VALIDOS`, se descarta y se devuelve `''` (se considera un valor inválido/desconocido).
- Antes y después de aplicar la función se imprime un conteo de valores (`value_counts`) para que se pueda comparar visualmente el "antes" (con todas las variantes) y el "después" (ya estandarizado), además del total de registros que quedaron en blanco.

### 4. Estandarización de sucursales

La lógica es análoga a la de medios de pago, pero aplicada a la columna `sucursal`:

```python
SUCURSALES_VALIDAS = {'centro', 'norte', 'local 1', 'online'}
```

- **`MAPA_SUCURSALES`:** traduce variantes como `"Sucursal Centro"` o `"Sucursal Norte"` a las formas canónicas `"centro"` y `"norte"`.
- **`normalizar_sucursal(valor)`:** sigue el mismo patrón de 4 pasos que `normalizar_medio_pago`: nulo → `''`, minúsculas sin tildes, búsqueda en el mapa, y descarte (`''`) si el resultado no está en `SUCURSALES_VALIDAS`.
- Se imprime el conteo de valores antes y después, y la cantidad de registros que quedaron en blanco por no encajar en ninguna sucursal válida.

### 5. Saneamiento de fechas inválidas

```python
df_clean['fecha_ok'] = pd.to_datetime(df_clean['fecha'], errors='coerce')
```

- Se crea una columna temporal `fecha_ok` convirtiendo `fecha` a tipo `datetime`. Con `errors='coerce'`, cualquier valor que no se pueda interpretar como fecha válida (por ejemplo `"2024-13-40"`, `"31/02/2025"` o el texto `"sin fecha"`) se convierte en `NaT` (Not a Time, el equivalente a un valor nulo para fechas).
- Se calculan dos métricas distintas:
  - `n_nulas_orig`: fechas que ya eran nulas en el dataset original.
  - `n_invalidas`: fechas que **tenían algún valor** pero no pudieron convertirse a una fecha real (es decir, estaban mal escritas o eran imposibles).
- Finalmente, la columna `fecha` original se reemplaza por `fecha_ok`, y se elimina la columna temporal (`drop(columns=['fecha_ok'])`).
- **Criterio aplicado:** no se intenta "adivinar" o corregir automáticamente una fecha inválida (por ejemplo, no se asume que `31/02/2025` "debería ser" el 28 de febrero), porque eso introduciría información inventada. En cambio, se preserva el hecho de que el dato es incorrecto marcándolo como nulo (`NaT`).

### 6. Saneamiento de columnas numéricas

Esta sección corrige las columnas `cantidad`, `precio_unitario`, `total` y `descuento`.

```python
def sanear_numerico(serie, nombre_col, min_val=None):
    serie_num = pd.to_numeric(serie, errors='coerce')
    ...
    serie_num[serie_num < 0] = np.nan
    if min_val is not None:
        serie_num[serie_num == 0] = np.nan
    return serie_num
```

- **`sanear_numerico(serie, nombre_col, min_val=None)`** es una función reutilizable que:
  1. Convierte la columna a tipo numérico con `pd.to_numeric(..., errors='coerce')`. Cualquier texto no numérico (por ejemplo `"gratis"`, `"dos"`, `"mal calculado"`) se transforma en `NaN`.
  2. Cuenta cuántos valores de texto no numérico había (`n_texto`), usando una expresión que verifica si el valor original era un string que no representa un número (permitiendo un único punto decimal y un signo negativo).
  3. Cuenta y elimina (convierte a `NaN`) los valores **negativos**, ya que no tiene sentido tener una cantidad, precio o total negativo.
  4. Si se pasa el parámetro `min_val` (usado para `cantidad` y `precio_unitario`), también cuenta y elimina los valores **iguales a 0**, porque una cantidad o un precio en 0 no es un dato válido para una venta.
  5. Imprime un resumen de cuántos valores se corrigieron en cada categoría.
- Esta función se aplica a:
  - `cantidad` (con `min_val=0`, por lo que también descarta ceros).
  - `precio_unitario` (con `min_val=0`, ídem).
  - `total` (sin `min_val`, por lo que los ceros se conservan, ya que un total en 0 podría ser válido en algunos contextos, p. ej. si hubo un descuento del 100%).
- **Descuento:** se trata por separado, fuera de la función `sanear_numerico`, porque su regla de validez es distinta: debe estar en el rango **[0, 100]** (un porcentaje). Cualquier valor fuera de ese rango (negativo o mayor a 100) se convierte en `NaN`.

### 7. Estandarización de columnas categóricas restantes

Se normalizan tres columnas más: `categoria`, `estado_envio` y `vendedor`. El patrón es similar al usado para medios de pago y sucursales, pero en este caso el resultado final **conserva mayúscula inicial** en lugar de quedar todo en minúsculas.

- **Categoría:**
  ```python
  MAPA_CATEGORIA = {
      'tecnologia': 'Tecnologia', 'tecnología': 'Tecnologia',
      'perifericos': 'Perifericos', ...
      'sin categoria': '',  # semánticamente nulo
  }
  ```
  - `normalizar_categoria(valor)` quita tildes y pasa a minúsculas solo para **buscar** en el mapa; si encuentra coincidencia, devuelve la forma canónica (con mayúscula inicial). Si no encuentra coincidencia, conserva el valor original (sin modificar su capitalización), asumiendo que es una categoría no contemplada en el mapeo.
  - El valor especial `"sin categoria"` se traduce explícitamente a `''`, tratándolo como un nulo semántico (el dato indica expresamente la ausencia de categoría).

- **Estado de envío:**
  ```python
  MAPA_ESTADO = {
      'pendiente': 'Pendiente', 'entregado': 'Entregado',
      'cancelado': 'Cancelado', 'en camino': 'En camino',
      'no aplica': 'No aplica',
  }
  ```
  - `normalizar_estado_envio(valor)` sigue la misma lógica: unifica mayúsculas/minúsculas y tildes, dejando como resultado final una forma con mayúscula inicial (`Pendiente`, `Entregado`, etc.).

- **Vendedor:**
  ```python
  MAPA_VENDEDOR = {
      'admin': 'Admin',
      'sin vendedor': '',  # semánticamente nulo
  }
  ```
  - `normalizar_vendedor(valor)` unifica la variante `"admin"`/`"Admin"` en `"Admin"`, y convierte `"sin vendedor"` en `''` (nulo semántico). Cualquier otro nombre de vendedor se conserva tal como vino en el dataset original.

En las tres columnas se imprime el conteo de valores antes y después de la normalización, para poder verificar visualmente el efecto de la limpieza.

### 8. Resumen final y exportación

```python
cols_aux = ['precio_num','cant_num','total_num','desc_num','fecha_dt',
            'total_esp','diff_total']
df_final = df_clean.drop(columns=[c for c in cols_aux if c in df_clean.columns], errors='ignore')
```

- Se eliminan las posibles columnas auxiliares que pudieran haber quedado en `df_clean` (por ejemplo, si se hubieran generado columnas de apoyo en el camino), usando `errors='ignore'` para que no falle si alguna de ellas no existe.
- Se imprime un resumen con la cantidad de filas originales, la cantidad de filas eliminadas por ser duplicadas exactas, y la cantidad final de filas en el dataset limpio.
- Finalmente, se exporta el `DataFrame` resultante a un nuevo archivo Excel:
  ```python
  df_final.to_excel("dataset_ventas_LIMPIO.xlsx", index=False)
  ```
  El parámetro `index=False` evita que pandas agregue una columna adicional con el índice numérico del `DataFrame`.

## Salida del script

Al ejecutarse, el script imprime en consola, en orden:

1. Dimensiones del dataset original.
2. Detalle de la eliminación de duplicados exactos.
3. Conteo de valores de `medio_pago` antes/después de estandarizar, y criterio aplicado.
4. Conteo de valores de `sucursal` antes/después de estandarizar, y criterio aplicado.
5. Estadísticas de fechas nulas/inválidas y criterio de saneamiento.
6. Estadísticas de valores de texto, negativos y ceros corregidos en `cantidad`, `precio_unitario`, `total` y `descuento`.
7. Conteo de valores antes/después para `categoria`, `estado_envio` y `vendedor`.
8. Un resumen final con filas originales, filas eliminadas y filas en el dataset limpio.

Y genera el archivo **`dataset_ventas_LIMPIO.xlsx`** en el mismo directorio donde se ejecuta el script.
