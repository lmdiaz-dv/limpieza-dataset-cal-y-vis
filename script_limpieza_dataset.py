import pandas as pd
import numpy as np
import re
import unicodedata
import warnings
warnings.filterwarnings("ignore")


# CARGA DEL DATASET

df = pd.read_excel("dataset_ventas_tp_5000_lineas_calidad_datos.xlsx")
print(f"Dataset cargado: {df.shape[0]} filas × {df.shape[1]} columnas\n")

# Conversión anticipada de columnas numéricas para análisis
df['precio_num']  = pd.to_numeric(df['precio_unitario'], errors='coerce')
df['cant_num']    = pd.to_numeric(df['cantidad'],         errors='coerce')
df['total_num']   = pd.to_numeric(df['total'],            errors='coerce')
df['desc_num']    = pd.to_numeric(df['descuento'],        errors='coerce')
df['fecha_dt']    = pd.to_datetime(df['fecha'],           errors='coerce')

sep = "=" * 70

df_clean = pd.read_excel("dataset_ventas_tp_5000_lineas_calidad_datos.xlsx")
print(f"\nDataset original: {len(df_clean)} filas\n")

# ─────────────────────────────────────────────────────────────────────────────
# DICCIONARIO DE MÉTRICAS (para el resumen TXT final)
# ─────────────────────────────────────────────────────────────────────────────
# Aquí se va guardando, paso a paso, la cantidad de nulos/valores inválidos
# "antes" y "después" de cada corrección, para poder armar al final un
# archivo de texto con el resumen comparativo de todo el proceso.
metricas = {}

def contar_nulos_semanticos(serie):
    """
    Cuenta valores 'nulos' considerando tanto NaN/None reales como
    strings vacíos o en blanco (que en este script se usan como
    'nulo semántico' luego de normalizar texto).
    """
    return (serie.isna() | (serie.astype(str).str.strip() == '')).sum()

# --- Snapshot "ANTES" (estado original, antes de cualquier corrección) ---
filas_originales = len(df_clean)

metricas['filas'] = {
    'antes': filas_originales,
    'despues': None,  # se completa al final
}

for col in ['medio_pago', 'sucursal', 'categoria', 'estado_envio', 'vendedor']:
    metricas[col] = {
        'nulos_antes': contar_nulos_semanticos(df_clean[col]),
        'nulos_despues': None,
    }

metricas['fecha'] = {
    'nulos_antes': df_clean['fecha'].isna().sum(),
    'invalidas_antes': None,   # se completa en la sección de fechas
    'nulos_despues': None,
}

for col in ['cantidad', 'precio_unitario', 'total', 'descuento']:
    metricas[col] = {
        'nulos_antes': pd.to_numeric(df_clean[col], errors='coerce').isna().sum(),
        'texto_invalido': 0,
        'negativos': 0,
        'ceros': 0,
        'fuera_de_rango': 0,
        'nulos_despues': None,
    }

print("──────────────────────────────────────────────────────────────────────")
print("Corrección 1: Eliminar filas exactamente duplicadas")
print("──────────────────────────────────────────────────────────────────────")

antes = len(df_clean)
df_clean = df_clean.drop_duplicates(keep='first')
despues = len(df_clean)
eliminadas = antes - despues
df_clean = df_clean.reset_index(drop=True)

metricas['filas']['duplicadas_eliminadas'] = eliminadas
metricas['filas']['despues_dedup'] = despues

print(f"""
  Criterio: se eliminan filas donde TODAS las columnas son idénticas
  (incluyendo id_venta y el resto de columnas).
  Se conserva la primera ocurrencia con keep='first'.

  Filas antes:      {antes:>6}
  Filas eliminadas: {eliminadas:>6}
  Filas después:    {despues:>6}
""")

print("──────────────────────────────────────────────────────────────────────")
print("Corrección 2: Estandarizar medios de pago")
print("──────────────────────────────────────────────────────────────────────")

# Lista de valores válidos (en minúscula y sin tilde, forma canónica)
MEDIOS_VALIDOS = {
    'efectivo',
    'transferencia',
    'tarjeta credito',
    'tarjeta debito',
    'mercado pago',
    'cheque'
}

def quitar_tildes(texto: str) -> str:
    """Elimina acentos/tildes de una cadena."""
    return ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )

# Mapa de normalización: variantes conocidas → forma correcta
MAPA_MEDIOS = {
    # transferencia
    'transf'          : 'transferencia',
    'transferencia'   : 'transferencia',
    # efectivo
    'efectivo'        : 'efectivo',
    # tarjeta crédito
    'tarjeta credito' : 'tarjeta credito',
    'tarjeta crédito' : 'tarjeta credito',
    # tarjeta débito
    'tarjeta debito'  : 'tarjeta debito',
    'tarjeta débito'  : 'tarjeta debito',
    # mercado pago
    'mercado pago'    : 'mercado pago',
    'mercadopago'     : 'mercado pago',
    # cheque
    'cheque'          : 'cheque',
}

def normalizar_medio_pago(valor) -> str:
    """
    1. Si es nulo → '' (en blanco)
    2. Normalizar a minúsculas sin tilde
    3. Buscar en mapa de equivalencias
    4. Si no está en la lista válida → '' (en blanco / nulo semántico)
    """
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    normalizado = quitar_tildes(str(valor).strip().lower())
    canonico = MAPA_MEDIOS.get(normalizado, normalizado)
    return canonico if canonico in MEDIOS_VALIDOS else ''

print("  Variantes detectadas antes de la corrección:")
print(df_clean['medio_pago'].value_counts(dropna=False).to_string())

df_clean['medio_pago'] = df_clean['medio_pago'].apply(normalizar_medio_pago)

print("\n  Valores después de la estandarización:")
print(df_clean['medio_pago'].value_counts(dropna=True).to_string())
nulos_mp = (df_clean['medio_pago'] == '').sum()
metricas['medio_pago']['nulos_despues'] = nulos_mp
print(f"\n  Registros dejados en blanco (nulo o fuera de lista): {nulos_mp}")
print(f"""
  Criterio aplicado con el mapeo:
    - "transf"     → "transferencia"
    - "MercadoPago"→ "mercado pago"  |  "Mercado Pago" → "mercado pago"
    - "tarjeta credito" / "Tarjeta Crédito" → "tarjeta credito"
    - "efectivo" / "Efectivo"         → "efectivo"
    - Todo en minúsculas y sin tildes.
    - Valores nulos → ''  |  Fuera de lista → '' (tratados como nulos)
""")

print("──────────────────────────────────────────────────────────────────────")
print("Corrección 3: Estandarizar sucursales")
print("──────────────────────────────────────────────────────────────────────")

SUCURSALES_VALIDAS = {
    'centro',
    'norte',
    'local 1',
    'online'
}

MAPA_SUCURSALES = {
    # centro
    'centro'           : 'centro',
    'sucursal centro'  : 'centro',
    # norte
    'norte'            : 'norte',
    'sucursal norte'   : 'norte',
    # local 1
    'local 1'          : 'local 1',
    # online
    'online'           : 'online',
}

def normalizar_sucursal(valor) -> str:
    """
    1. Si es nulo → ''
    2. Normalizar a minúsculas sin tilde
    3. Buscar en mapa de equivalencias
    4. Si no está en lista válida → ''
    """
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    normalizado = quitar_tildes(str(valor).strip().lower())
    canonico = MAPA_SUCURSALES.get(normalizado, normalizado)
    return canonico if canonico in SUCURSALES_VALIDAS else ''

print("  Variantes detectadas antes de la corrección:")
print(df_clean['sucursal'].value_counts(dropna=False).to_string())

df_clean['sucursal'] = df_clean['sucursal'].apply(normalizar_sucursal)

print("\n  Valores después de la estandarización:")
print(df_clean['sucursal'].value_counts(dropna=True).to_string())
nulos_suc = (df_clean['sucursal'] == '').sum()
metricas['sucursal']['nulos_despues'] = nulos_suc
print(f"\n  Registros dejados en blanco: {nulos_suc}")
print(f"""
  Criterio aplicado:
    - "Sucursal Centro" → "centro"   |  "Centro"       → "centro"
    - "Sucursal Norte"  → "norte"    |  "Norte"         → "norte"
    - "local 1"  / "Local 1"         → "local 1"
    - "online"   / "Online"          → "online"
    - Todo en minúsculas y sin tildes.
    - Nulos o valores fuera de lista  → ''
""")

print("──────────────────────────────────────────────────────────────────────")
print("Corrección 5: Sanear fechas inválidas")
print("──────────────────────────────────────────────────────────────────────")

df_clean['fecha_ok'] = pd.to_datetime(df_clean['fecha'], errors='coerce')

invalidas = df_clean['fecha_ok'].isna() & df_clean['fecha'].notna()
n_invalidas = invalidas.sum()
n_nulas_orig = df_clean['fecha'].isna().sum()

metricas['fecha']['invalidas_antes'] = n_invalidas
# n_nulas_orig ya coincide con metricas['fecha']['nulos_antes'] capturado al inicio,
# se reafirma aquí por claridad ya que se calcula sobre el mismo estado de 'fecha'.
metricas['fecha']['nulos_antes'] = n_nulas_orig

print(f"""
  Fechas nulas en el original:      {n_nulas_orig:>5}
  Fechas con formato inválido:      {n_invalidas:>5}
  (Ejemplos: "2024-13-40", "sin fecha", "31/02/2025", "2025/99/01")

  Criterio: se convierte 'fecha' a datetime con errors='coerce'.
  Las fechas inválidas quedan como NaT (equivalente a nulo en datetime).
  Esto preserva la información sabiendo que no puede ser corregida
  automáticamente sin la fuente original.

  Columna resultante: 'fecha_ok' (tipo datetime64)
""")

df_clean['fecha'] = df_clean['fecha_ok']
df_clean.drop(columns=['fecha_ok'], inplace=True)

fecha_validas = df_clean['fecha'].notna().sum()
metricas['fecha']['nulos_despues'] = df_clean['fecha'].isna().sum()
print(f"  Fechas válidas conservadas: {fecha_validas}")
print(f"  Fechas como NaT (nulo):     {df_clean['fecha'].isna().sum()}")

print("\n──────────────────────────────────────────────────────────────────────")
print("Corrección 5: Sanear columnas numéricas con valores inválidos")
print("──────────────────────────────────────────────────────────────────────")

def sanear_numerico(serie, nombre_col, min_val=None):
    """
    - Convierte a numérico (strings como 'gratis', 'dos' → NaN)
    - Valores negativos → NaN (no tienen sentido en precio/cantidad/total)
    - Valores = 0 en cantidad/precio → NaN
    Devuelve la tupla (serie_limpia, dict_metricas) para poder registrar
    el detalle en el resumen final.
    """
    serie_num = pd.to_numeric(serie, errors='coerce')
    n_texto   = serie.apply(lambda x: isinstance(x,str) and not str(x).replace('.','',1).replace('-','',1).isdigit() if pd.notna(x) else False).sum()
    n_neg     = (serie_num < 0).sum()
    serie_num[serie_num < 0] = np.nan
    n_cero = 0
    if min_val is not None:
        n_cero = (serie_num == 0).sum()
        serie_num[serie_num == 0] = np.nan
    print(f"  [{nombre_col}]")
    print(f"    Valores de texto convertidos a NaN: {n_texto}")
    print(f"    Valores negativos convertidos a NaN: {n_neg}")
    if min_val is not None:
        print(f"    Valores = 0 convertidos a NaN:       {n_cero}")
    metricas_col = {
        'texto_invalido': n_texto,
        'negativos': n_neg,
        'ceros': n_cero,
        'nulos_despues': serie_num.isna().sum(),
    }
    return serie_num, metricas_col

print()
df_clean['cantidad'], metricas_cant = sanear_numerico(df_clean['cantidad'],        'cantidad',        min_val=0)
df_clean['precio_unitario'], metricas_precio = sanear_numerico(df_clean['precio_unitario'], 'precio_unitario', min_val=0)
df_clean['total'], metricas_total = sanear_numerico(df_clean['total'],           'total',           min_val=0)

metricas['cantidad'].update(metricas_cant)
metricas['precio_unitario'].update(metricas_precio)
metricas['total'].update(metricas_total)

# Descuento: debe estar en [0, 100]
desc_num = pd.to_numeric(df_clean['descuento'], errors='coerce')
n_desc_rango = ((desc_num < 0) | (desc_num > 100)).sum()
desc_num[(desc_num < 0) | (desc_num > 100)] = np.nan
df_clean['descuento'] = desc_num
metricas['descuento']['fuera_de_rango'] = n_desc_rango
metricas['descuento']['nulos_despues'] = desc_num.isna().sum()
print(f"  [descuento]")
print(f"    Valores fuera de [0, 100] → NaN: {n_desc_rango}")

print(f"""
  Criterio:
    - Textos no numéricos ("gratis", "mal calculado", "dos") → NaN
    - Precios, cantidades y totales negativos                 → NaN
    - Cantidades y precios igual a 0                          → NaN
    - Descuentos fuera del rango [0, 100]                     → NaN
""")

print("──────────────────────────────────────────────────────────────────────")
print("Corrección 6: Estandarizar columnas categóricas restantes")
print("             (categoria, estado_envio, vendedor)")
print("──────────────────────────────────────────────────────────────────────")

# --- Categoría ---
MAPA_CATEGORIA = {
    'tecnologia'    : 'Tecnologia',
    'tecnología'    : 'Tecnologia',
    'perifericos'   : 'Perifericos',
    'periféricos'   : 'Perifericos',
    'accesorios'    : 'Accesorios',
    'audio'         : 'Audio',
    'hogar'         : 'Hogar',
    'computacion'   : 'Computacion',
    'computación'   : 'Computacion',
    'sin categoria' : '',   # semánticamente nulo
}

def normalizar_categoria(valor):
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    norm = quitar_tildes(str(valor).strip().lower())
    return MAPA_CATEGORIA.get(norm, valor.strip())

print("\n  [categoria] – antes:")
print(df_clean['categoria'].value_counts(dropna=False).to_string())
df_clean['categoria'] = df_clean['categoria'].apply(normalizar_categoria)
metricas['categoria']['nulos_despues'] = contar_nulos_semanticos(df_clean['categoria'])
print("\n  [categoria] – después:")
print(df_clean['categoria'].value_counts(dropna=True).to_string())

# --- Estado de envío ---
MAPA_ESTADO = {
    'pendiente'  : 'Pendiente',
    'entregado'  : 'Entregado',
    'cancelado'  : 'Cancelado',
    'en camino'  : 'En camino',
    'no aplica'  : 'No aplica',
}

def normalizar_estado_envio(valor):
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    norm = quitar_tildes(str(valor).strip().lower())
    return MAPA_ESTADO.get(norm, valor.strip())

print("\n  [estado_envio] – antes:")
print(df_clean['estado_envio'].value_counts(dropna=False).to_string())
df_clean['estado_envio'] = df_clean['estado_envio'].apply(normalizar_estado_envio)
metricas['estado_envio']['nulos_despues'] = contar_nulos_semanticos(df_clean['estado_envio'])
print("\n  [estado_envio] – después:")
print(df_clean['estado_envio'].value_counts(dropna=True).to_string())

# --- Vendedor ---
MAPA_VENDEDOR = {
    'admin'          : 'Admin',
    'sin vendedor'   : '',   # semánticamente nulo
}

def normalizar_vendedor(valor):
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    norm = quitar_tildes(str(valor).strip().lower())
    return MAPA_VENDEDOR.get(norm, valor.strip())

print("\n  [vendedor] – antes:")
print(df_clean['vendedor'].value_counts(dropna=False).to_string())
df_clean['vendedor'] = df_clean['vendedor'].apply(normalizar_vendedor)
metricas['vendedor']['nulos_despues'] = contar_nulos_semanticos(df_clean['vendedor'])
print("\n  [vendedor] – después:")
print(df_clean['vendedor'].value_counts(dropna=True).to_string())

print(f"""
  Criterio aplicado:
    - Unificación de case (Pendiente / pendiente → Pendiente)
    - Eliminación de tildes para comparación, resultado con mayúscula inicial
    - "Sin categoria" / "Sin vendedor" → '' (vacío = nulo semántico)
    - "admin" / "Admin" → "Admin" (forma canónica)
""")



# RESUMEN FINAL
print(sep)
print("RESUMEN FINAL")
print(sep)

# Limpiar columnas auxiliares antes de guardar
cols_aux = ['precio_num','cant_num','total_num','desc_num','fecha_dt',
            'total_esp','diff_total']
df_final = df_clean.drop(columns=[c for c in cols_aux if c in df_clean.columns], errors='ignore')

metricas['filas']['despues'] = len(df_final)

print(f"""
  Filas originales:                  {filas_originales:>6}
  Filas eliminadas (exactas):        {filas_originales - len(df_final):>6}
  Filas en dataset limpio:           {len(df_final):>6}
""")

# Exportar dataset limpio
output_path = "dataset_ventas_LIMPIO.xlsx"
df_final.to_excel(output_path, index=False)
print(f"  Dataset limpio exportado a:\n  {output_path}\n")

# ─────────────────────────────────────────────────────────────────────────────
# GENERACIÓN DEL ARCHIVO TXT CON EL RESUMEN DE MÉTRICAS (ANTES vs DESPUÉS)
# ─────────────────────────────────────────────────────────────────────────────
# Se construye un reporte de texto plano que consolida, columna por columna,
# la cantidad de nulos/valores inválidos detectados antes de limpiar y la
# cantidad que queda después de aplicar todas las correcciones. Sirve como
# evidencia rápida del impacto del proceso de limpieza, sin tener que revisar
# todo el log impreso en consola.

def linea(titulo, ancho=70, relleno="─"):
    return relleno * ancho if titulo is None else f" {titulo} ".center(ancho, relleno)

reporte = []
reporte.append("=" * 70)
reporte.append("RESUMEN DE MÉTRICAS DE CALIDAD DE DATOS — ANTES vs DESPUÉS")
reporte.append("=" * 70)
reporte.append("")
reporte.append(f"Archivo original: dataset_ventas_tp_5000_lineas_calidad_datos.xlsx")
reporte.append(f"Archivo limpio:   {output_path}")
reporte.append("")

# --- Filas ---
reporte.append(linea("FILAS"))
reporte.append(f"  Filas originales:                 {metricas['filas']['antes']:>6}")
reporte.append(f"  Filas duplicadas eliminadas:       {metricas['filas']['duplicadas_eliminadas']:>6}")
reporte.append(f"  Filas en dataset limpio:           {metricas['filas']['despues']:>6}")
reporte.append("")

# --- Columnas categóricas (medio_pago, sucursal, categoria, estado_envio, vendedor) ---
reporte.append(linea("COLUMNAS CATEGÓRICAS — valores nulos / sin estandarizar"))
reporte.append(f"  {'Columna':<15}{'Nulos antes':>15}{'Nulos después':>16}{'Variación':>12}")
reporte.append(f"  {'-'*15:<15}{'-'*15:>15}{'-'*16:>16}{'-'*12:>12}")
for col in ['medio_pago', 'sucursal', 'categoria', 'estado_envio', 'vendedor']:
    a = metricas[col]['nulos_antes']
    d = metricas[col]['nulos_despues']
    var = d - a
    reporte.append(f"  {col:<15}{a:>15}{d:>16}{var:>+12}")
reporte.append("")
reporte.append("  Nota: en 'antes' se cuentan nulos reales (NaN) + strings vacíos.")
reporte.append("  En 'después' se cuentan los valores que quedaron en blanco ('')")
reporte.append("  por ser nulos originales o por no encajar en ninguna categoría válida.")
reporte.append("")

# --- Fechas ---
reporte.append(linea("FECHA"))
reporte.append(f"  Nulas en el original:              {metricas['fecha']['nulos_antes']:>6}")
reporte.append(f"  Con formato inválido (no nulas):   {metricas['fecha']['invalidas_antes']:>6}")
reporte.append(f"  Total nulas/NaT después:           {metricas['fecha']['nulos_despues']:>6}")
reporte.append("")

# --- Columnas numéricas ---
reporte.append(linea("COLUMNAS NUMÉRICAS — valores inválidos detectados y corregidos"))
reporte.append(f"  {'Columna':<17}{'Nulos antes':>13}{'Texto inv.':>12}{'Negativos':>11}{'Ceros':>8}{'Nulos después':>15}")
reporte.append(f"  {'-'*17:<17}{'-'*13:>13}{'-'*12:>12}{'-'*11:>11}{'-'*8:>8}{'-'*15:>15}")
for col in ['cantidad', 'precio_unitario', 'total']:
    m = metricas[col]
    reporte.append(
        f"  {col:<17}{m['nulos_antes']:>13}{m['texto_invalido']:>12}"
        f"{m['negativos']:>11}{m['ceros']:>8}{m['nulos_despues']:>15}"
    )
reporte.append("")
reporte.append(f"  [descuento] — debe estar en el rango [0, 100]")
reporte.append(f"    Nulos antes:                     {metricas['descuento']['nulos_antes']:>6}")
reporte.append(f"    Valores fuera de [0, 100]:        {metricas['descuento']['fuera_de_rango']:>6}")
reporte.append(f"    Nulos después:                   {metricas['descuento']['nulos_despues']:>6}")
reporte.append("")

# --- Totales generales ---
total_nulos_antes = (
    sum(metricas[c]['nulos_antes'] for c in
        ['medio_pago', 'sucursal', 'categoria', 'estado_envio', 'vendedor',
         'cantidad', 'precio_unitario', 'total', 'descuento'])
    + metricas['fecha']['nulos_antes']
)
total_nulos_despues = (
    sum(metricas[c]['nulos_despues'] for c in
        ['medio_pago', 'sucursal', 'categoria', 'estado_envio', 'vendedor',
         'cantidad', 'precio_unitario', 'total', 'descuento'])
    + metricas['fecha']['nulos_despues']
)

reporte.append(linea("TOTAL GENERAL DE NULOS/INVÁLIDOS (todas las columnas relevadas)"))
reporte.append(f"  Total antes:                       {total_nulos_antes:>6}")
reporte.append(f"  Total después:                     {total_nulos_despues:>6}")
reporte.append(f"  Variación:                         {total_nulos_despues - total_nulos_antes:>+6}")
reporte.append("")
reporte.append("  Nota: el total 'después' suele ser mayor que el total 'antes'")
reporte.append("  porque varios valores que originalmente parecían válidos")
reporte.append("  (texto fuera de lista, números negativos/cero, fechas con")
reporte.append("  formato imposible, etc.) se detectan como inválidos durante")
reporte.append("  la limpieza y se convierten explícitamente en nulos, en lugar")
reporte.append("  de dejarlos como datos incorrectos silenciosos.")
reporte.append("")
reporte.append("=" * 70)
reporte.append("Fin del resumen.")
reporte.append("=" * 70)

reporte_txt_path = "resumen_metricas_limpieza.txt"
with open(reporte_txt_path, "w", encoding="utf-8") as f:
    f.write("\n".join(reporte))

print(f"  Resumen de métricas (antes/después) exportado a:\n  {reporte_txt_path}\n")
print("  ¡Proceso completado exitosamente!")
