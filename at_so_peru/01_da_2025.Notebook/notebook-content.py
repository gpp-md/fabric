# Fabric notebook source

# METADATA ********************

# META {
# META   "kernel_info": {
# META     "name": "jupyter",
# META     "jupyter_kernel_name": "python3.11"
# META   },
# META   "dependencies": {
# META     "lakehouse": {
# META       "default_lakehouse": "39f98f73-8776-48bc-905b-31d3f11de9f7",
# META       "default_lakehouse_name": "lh_2025_peru",
# META       "default_lakehouse_workspace_id": "9680dc28-bc8c-4d47-900d-ccf7674e4b49",
# META       "known_lakehouses": [
# META         {
# META           "id": "39f98f73-8776-48bc-905b-31d3f11de9f7"
# META         }
# META       ]
# META     }
# META   }
# META }

# CELL ********************

# MAGIC %%configure -f
# MAGIC {
# MAGIC     "vcores": 8
# MAGIC }

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

# Welcome to your new notebook
# Type here in the cell editor to add code!
%pip install wget xlsxwriter fastexcel pyarrow polars deltalake

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

#pip install xlrd
import os
import sys
import pandas as pd
import openpyxl as olx
import time
from datetime import datetime, date, timedelta
import glob
import warnings
import pyarrow as pa
import pyarrow.csv as csv
import wget
from os import remove
import urllib.request
import polars as pl
import polars.selectors as cs
from xlsxwriter import Workbook
import zipfile
from deltalake.writer import write_deltalake
# !apt install chromium-chromedriver
import ssl
ssl._create_default_https_context = ssl._create_unverified_context
warnings.filterwarnings("ignore")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************


fecha = datetime.today().strftime('%d/%m/%Y')
start_time = time.time()
# ruta = os.getcwd()
# cwd = os.getcwd()
cwd = '/lakehouse/default/Files'
ruta ='/lakehouse/default/Files/archivos_csv'
# ruta_pq ='/lakehouse/default/Tables/'
dtype_polars = {'SECTOR':pl.String,
           'PLIEGO':pl.String,
           'SEC_EJEC':pl.String,
           'TIPO_RECURSO':pl.String,
           'EJECUTORA':pl.String,
           'PROGRAMA_PPTO':pl.String,
           'FUNCION':pl.String,
           'DIVISION_FUNCIONAL':pl.String,
           'GRUPO_FUNCIONAL':pl.String,
           'FINALIDAD':pl.String,
           'RUBRO':pl.String,
           'TIPO_RECURSO':pl.String,
           'PRODUCTO_PROYECTO':pl.String,
           'ACTIVIDAD_ACCION_OBRA':pl.String,
           'DEPARTAMENTO_META':pl.String,
           'DEPARTAMENTO_EJECUTORA':pl.String,
           'PROVINCIA_EJECUTORA':pl.String,
           'DISTRITO_EJECUTORA':pl.String,
           'META':pl.String,
           'FUENTE_FINANCIAMIENTO':pl.String,
           'RUBRO':pl.String,
           'TIPO_TRANSACCION':pl.String,
           'GENERICA':pl.String,
           'SUBGENERICA':pl.String,
           'SUBGENERICA_DET':pl.String,
           'ESPECIFICA':pl.String,
           'ESPECIFICA_DET':pl.String,
           'CATEGORIA_GASTO':pl.String,

           'id_clasificador':pl.String,
           'Especifica det':pl.String,
           'id_clasificador_2':pl.String,
           'origen_clasificador':pl.String,
           'estado':pl.String,
           'es_presupuestal':pl.String,
           'restring':pl.String,
           'key_total':pl.String,
           'key_office':pl.String,
           'key_iri':pl.String,
           'dnpp':pl.String,
           'PROYECCION_DEVENGADO_OCTUBRE':pl.Float64,
           'PROYECCION_DEVENGADO_NOVIEMBRE':pl.Float64,
           'PROYECCION_DEVENGADO_DICIEMBRE':pl.Float64,
           'MONTO_DEVENGADO_ENERO':pl.Float64,
           'MONTO_DEVENGADO_FEBRERO':pl.Float64,
           'MONTO_DEVENGADO_MARZO':pl.Float64,
           'MONTO_DEVENGADO_ABRIL':pl.Float64,
           'MONTO_DEVENGADO_MAYO':pl.Float64,
           'MONTO_DEVENGADO_JUNIO':pl.Float64,
           'MONTO_DEVENGADO_JULIO':pl.Float64,
           'MONTO_DEVENGADO_AGOSTO':pl.Float64,
           'MONTO_DEVENGADO_SEPTIEMBRE':pl.Float64,
           'MONTO_DEVENGADO_OCTUBRE':pl.Float64,
           'MONTO_DEVENGADO_NOVIEMBRE':pl.Float64,
           'MONTO_DEVENGADO_DICIEMBRE':pl.Float64,
           'TOTAL PROGG':pl.Float64,
           'prog_mef':pl.Float64,
           'MONTO_PIA':pl.Int64,
           'MONTO_PIM':pl.Int64,
           'mto_girado_01': pl.Float64,
           'mto_girado_02': pl.Float64,
           'mto_girado_03': pl.Float64,
           'mto_girado_04': pl.Float64,
           'mto_girado_05': pl.Float64,
           'mto_girado_06': pl.Float64,
           'mto_girado_07': pl.Float64,
           'mto_girado_08': pl.Float64,
           'mto_girado_09': pl.Float64,
           'mto_girado_10': pl.Float64,
           'mto_girado_11': pl.Float64,
           'mto_girado_12': pl.Float64,
           'Girado.': pl.Float64,
           'Saldo Girado.': pl.Float64
           }

df1 = pl.DataFrame()
print("--- %s ANTES DE LEER  ---" % (time.time() - start_time))
dfc = pl.read_excel(f'{cwd}/programacion/clasif.xlsx', sheet_name="CLASIF",)# schema_overrides=dtype_polars,)# engine='openpyxl')# read_options={"infer_schema_length": None})
dfp = pl.read_excel(f"{cwd}/programacion/programacion_1669.xlsx",)# schema_overrides=dtype_polars,)# engine='openpyxl')
dfpi = pl.read_excel(f"{cwd}/programacion/Consolidado_formato_19.xlsx", )#schema_overrides=dtype_polars, )#engine='openpyxl')#, read_options={"infer_schema_length": 10000})#, dtype={'key_total':str})
dfo = pl.read_excel(f"{cwd}/programacion/oficinas_data.xlsx", schema_overrides={"coordinador": pl.String, "coordinador_email": pl.String})#engine='openpyxl')#read_options={"infer_schema_length": None})#dtype={'key_office':str})
df_i = pl.read_excel(f"{cwd}/programacion/iri.xlsx",)# schema_overrides=dtype_polars, )#engine='openpyxl')#read_options={"infer_schema_length": None})#dtype={'key_iri':str})
df_im = pl.read_excel(f"{cwd}/programacion/inversiones_mod.xlsx",)# schema_overrides=dtype_polars, )#engine='openpyxl')#read_options={"infer_schema_length": None})#dtype={'dnpp':str})
dfadd= pl.read_excel(f"{cwd}/programacion/agrega_inversiones_2023.xlsx", sheet_name="Sheet1",)# schema_overrides=dtype_polars, )#engine='openpyxl')#read_options={"infer_schema_length": 1000000})#ignore_errors=True)#dtype=dtype_dic)
# print('dfadd',dfadd)
df_estrategica = pl.read_excel(f"{cwd}/programacion/Anexo_I_LeyPpto2024.xlsx",)# schema_overrides=dtype_polars, )#engine='openpyxl')#read_options={"infer_schema_length": 1000000})#dtype={'dnpp':str})
df_a4 = pl.read_excel(f"{cwd}/programacion/Anexo_IV_LeyPpto2024.xlsx",)# schema_overrides=dtype_polars, )#engine='openpyxl')#read_options={"infer_schema_length": 1000000})#dtype={'dnpp':str})
df_a6 = pl.read_excel(f"{cwd}/programacion/Anexo_VI_LeyPpto2024.xlsx",)# schema_overrides=dtype_polars, )#engine='openpyxl')#read_options={"infer_schema_length": 1000000})#dtype={'dnpp':str})
print("--- %s DESPUES DE LEER VARIABLES ---" % (time.time() - start_time))

############
# anioo = int(input('Ingrese año :'))
# sector = input('Ingrese código de Sector :')
anioo = 2025
sector = '1'
mainpath = ruta
filename_web = f'{anioo}-Gasto-Devengado-Diario.csv'
# filename_web = f"{anioo}-Gasto-Devengado.csv"
# filename_web = f'{anioo}-Gasto-Devengado-Mensual.csv'
# filename_web = '2013-Gasto-Devengado.csv'
#filename_web = f"mtc_gasto_2_36_al-27062023.csv"
# filename_web = f"Data_2023.xlsx"
# filename_web = f"ReporteGasto*.xl*"

#########       DESCARGA ZIP          #########
filename_base = f'{anioo}-Gasto-Devengado-Diario'
# filename_base = f'{anioo}-Gasto-Devengado'
filename_zip = f'{filename_base}.zip'
filename_csv_inside_zip = f'{filename_base}.csv' # Asumimos que el CSV dentro del ZIP tiene el mismo nombre base
url_zip = f'https://fs.datosabiertos.mef.gob.pe/datastorefiles/{filename_zip}'
archivo_zip_local = os.path.join(mainpath, filename_zip)
print('filename_zip',archivo_zip_local)
try:
    os.remove(f'{archivo_zip_local}')
    print(f"Archivo '{archivo_zip_local}' eliminado correctamente.")
except FileNotFoundError:
    print(f"Error: Archivo '{archivo_zip_local}' no encontrado.")
except PermissionError:
    print(f"Error: No tienes permisos para eliminar '{archivo_zip_local}'.")
except Exception as e:
    print(f"Error inesperado: {e}")
#####
if os.path.exists(archivo_zip_local):
    print('Archivo presente :\n',archivo_zip_local)
else:
    print('Arquivo ZIP no prisenti, download now')
    try:
        wget.download(url=url_zip, out=mainpath)
        print(f'\nArchivo ZIP descargado exitosamente en: {archivo_zip_local}')#, bar=bar_thermometer)
    except Exception as e:
        print(f'Error al descargar el archivo ZIP: {e}')
        exit()
# Ahora vamos a leer el CSV dentro del archivo ZIP directamente con Polars
try:
    with zipfile.ZipFile(archivo_zip_local, 'r') as zf:
        with zf.open(filename_csv_inside_zip) as csvfile:
            df1 = pl.read_csv(csvfile, dtypes=dtype_polars)
    print('Archivo CSV dentro del ZIP leído exitosamente con Polars.')
    print("--- %s seconds --- LECTURA DE CSV " % (time.time() - start_time))
    # print(df1.head()) # Para ver las primeras filas
except zipfile.BadZipFile:
    print(f'Error: El archivo {archivo_zip_local} parece ser un archivo ZIP inválido.')
except KeyError:
    print(f'Error: No se encontró el archivo {filename_csv_inside_zip} dentro del ZIP.')
except Exception as e:
    print(f'Error al leer el CSV desde el ZIP con Polars: {e}')
    print("--- %s seconds --- LECTURA DE CSV " % (time.time() - start_time))
print("--- %s seconds --- UNICO DATAFRAME" % (time.time() - start_time))
#########       DESCARGA ZIP          #########

# #############################################################################################
# archivo = os.path.join(ruta, filename_web)
# try:
#     os.remove(f'{archivo}')
#     print(f"Archivo '{archivo}' eliminado correctamente.")
# except FileNotFoundError:
#     print(f"Error: Archivo '{archivo}' no encontrado.")
# except PermissionError:
#     print(f"Error: No tienes permisos para eliminar '{archivo}'.")
# except Exception as e:
#     print(f"Error inesperado: {e}")
# if os.path.exists(archivo) == True:
#     print('Archivo presente :\n',archivo)
# else:
#     print('Arquivo no prisenti, download now')
#     url = 'https://fs.datosabiertos.mef.gob.pe/datastorefiles/' + filename_web
#     filename_web = wget.download(url=url, out=ruta)#, bar=bar_thermometer)
#     print('\n', filename_web)
# #############################################################################################
# fullpath_web = os.path.join(mainpath, filename_web)
# files_web = glob.glob(fullpath_web)
# for excel in files_web:
#     print(excel)
#     df1 = pl.read_csv(excel, ignore_errors=True, dtypes=dtype_polars, use_pyarrow=True)
#     print("--- %s seconds --- LECTURA DE CSV " % (time.time() - start_time))
# print("--- %s seconds --- UNICO DATAFRAME" % (time.time() - start_time))

###################################                PYARROW               ################################
print('Lectura inicial :', df1.shape)
prn_nivel = (
    df1
    .group_by('NIVEL_GOBIERNO')
    .agg(pl.sum('MONTO_PIA', 'MONTO_PIM', 'MONTO_DEVENGADO_ANUAL').cast(pl.Int64))
    .sort(['NIVEL_GOBIERNO'])
    )
print('Montos Iniciales', prn_nivel)

# df.with_column(
#     pl.col("x").round(2)
# )
    # df1 = pl.concat([df1, df2],axis=0)
print("--- %s seconds --- CONCATENADO" % (time.time() - start_time))

# if os.path.exists(f'{sector}') == True:
#     os.chdir(f'{sector}')
# else:
#     os.mkdir(f'{sector}')
#     os.chdir(f'{sector}')

# df1.drop_duplicates(inplace=True)
print("--- %s seconds --- DIRECTORIO" % (time.time() - start_time))

#### ARREGLOS DE DATA DE DATOS ABIERTOS #####


df1 = df1.rename({col: col.lower() for col in df1.columns})
# print('columnas \n',df1)
print("--- %s seconds --- COLUMNAS A MAYSUC" % (time.time() - start_time))

######################     SELECT: SECTOR O EJECUTORA     ################################

# if len(str(sector)) == 2:
#     # print(len(str(sector)))
#     print(f'{len(str(sector))} Entidad : {sector}')
#     df1 = df1.filter(pl.col('sector') == sector)
#     # print('SELECT: SECTOR O EJECUTORA',df1)

# elif len(str(sector)) == 6:
#     # print(len(str(sector)))
#     print(f'{len(str(sector))} Entidad : {sector}')
#     df1 = df1.filter(pl.col('sec_ejec')== sector)
# elif len(str(sector)) == 3:
#     # print(len(str(sector)))
#     print(f'{len(str(sector))} Entidad : {sector}')
#     df1 = df1.filter(pl.col('pliego')== sector)
# elif sector == '1':
#     print('All entity')

print("--- %s seconds --- FILTRA ENTIDAD" % (time.time() - start_time))
######################   FILTRA AÑO     ################################
anio = df1['ano_eje'].unique()

######################   AGREGAR INVERSIONES 2023     ################################
# print('subido :', df1.shape )

# columnas = dfadd.with_columns([col_name.upper() for col_name in dfadd.columns])
# print('columnas \n',columnas)

dfadd = dfadd.rename({col: col.lower() for col in dfadd.columns})
# print('columnas \n',dfadd)


# # Cambiar todas las columnas a minúsculas
# dfadd = dfadd.select(pl.col("*").to_lowercase())
# Imprimir el resultado
# print(dfadd)
# columnas = dfadd.columns.str.lower().tolist()
# dfadd.columns= columnas

df1 = pl.concat([df1, dfadd],how='vertical_relaxed')#,axis=0)
# df1.extend(dfadd)
# print(df1)
print("--- %s seconds --- CON DATA DE INVERSIONES" % (time.time() - start_time))
######################   DATA PURA DE DATOS ABIERTOS     ################################
# anio = df1['ano_eje'].unique()
# df1.to_excel(f'{sector}_ejecucion_gastos_{anio[0]}_noviembre_datos_abiertos_data.xlsx', index=False)
print("--- %s seconds --- DATA DE DATOS ABIERTOS" % (time.time() - start_time))
# print(df1.groupby(['nivel_gobierno'])['monto_pim'].sum())

######################               ################################
# df1.drop(df1.columns[[1,41,54]], axis='columns', inplace=True)
# df1 = df_gn.copy()
# df1['ejecutora_nombre'] = df1.with_columns(pl.col('ejecutora_nombre').str.strip_chars())
# df1 = df1.with_columns(subgenerica_det=pl.col("subgenerica_det").str.strip_chars())
df1 = df1.with_columns(ejecutora_nombre=pl.col("ejecutora_nombre").str.strip_chars())
df1 = df1.with_columns(tipo_prod_proy = pl.lit(''))
# df1 = df1.with_columns(pl.when(pl.col('producto_proyecto').str.slice(0,1) == '0').then(pl.lit('0.OTROS')).otherwise(df1['tipo_prod_proy']).alias('tipo_prod_proy'))

df1 = df1.with_columns(pl
                       .when(pl.col('producto_proyecto').str.slice(0,1) == '0').then(pl.lit('0.OTROS'))
                       .when(pl.col('producto_proyecto').str.slice(0,1) == '2').then(pl.lit('2.PROYECTO'))
                       .when(pl.col('producto_proyecto').str.slice(0,1) == '3').then(pl.lit('3.PRODUCTO'))
                       .otherwise(df1['tipo_prod_proy']).alias('tipo_prod_proy')
                       )



################# UNIR COLUMNAS  ######################
# df1 = df1.with_columns(
#     (pl.col('nivel_gobierno') + ". " + pl.col("nivel_gobierno_nombre")).alias('nivel.'),
#     (pl.col('sector') + ". " + pl.col("sector_nombre")).alias('sector.'),
#     )
# print('columnas nuevas',df1)

df1 = df1.with_columns(
    (pl.col('nivel_gobierno') + ". " + pl.col("nivel_gobierno_nombre")).alias('nivel.'),
    (pl.col('sector') + ": " + pl.col("sector_nombre")).alias('sector.'),
    (pl.col('pliego') + ": " + pl.col("pliego_nombre")).alias('pliego.'),
    (pl.col('departamento_ejecutora').cast(pl.String) + ". " + pl.col("departamento_ejecutora_nombre")).alias('departamento_ejecutora.'),
    (pl.col('provincia_ejecutora') + ". " + pl.col("provincia_ejecutora_nombre")).alias('provincia_ejecutora.'),
    (pl.col('distrito_ejecutora') + ". " + pl.col("distrito_ejecutora_nombre")).alias('distrito_ejecutora.'),
    (pl.col('programa_ppto') + ". " + pl.col("programa_ppto_nombre")).alias('programa_pptal'),
    (pl.col('producto_proyecto') + ". " + pl.col("producto_proyecto_nombre")).alias('producto_proyecto.'),
    (pl.col('actividad_accion_obra') + ". " + pl.col("actividad_accion_obra_nombre")).alias('activ_obra_accinv'),
    (pl.col('ejecutora') + ". " + pl.col("ejecutora_nombre")).alias('unidad_ejecutora'),
    (pl.col('funcion') + ". " + pl.col("funcion_nombre")).alias('funcion.'),
    (pl.col('division_funcional') + ". " + pl.col("division_funcional_nombre")).alias('division_fn'),
    (pl.col('grupo_funcional') + ". " + pl.col("grupo_funcional_nombre")).alias('grupo_fn'),
    (pl.col('finalidad') + ". " + pl.col("meta_nombre")).alias('finalidad.'),
    (pl.col('departamento_meta').cast(pl.String) + ". " + pl.col("departamento_meta_nombre")).alias('departamento_meta.'),
    (pl.col('fuente_financiamiento') + ". " + pl.col("fuente_financiamiento_nombre")).alias('fuente'),
    (pl.col('rubro') + ". " + pl.col("rubro_nombre")).alias('rubro.'),
    (pl.col('tipo_recurso') + ". " + pl.col("tipo_recurso_nombre")).alias('tipo_recurso.'),
    (pl.col('categoria_gasto') + ". " + pl.col("categoria_gasto_nombre")).alias('categoria_gasto.'),
    (pl.col('tipo_transaccion') + ". " + pl.col("tipo_transaccion_nombre")).alias('tipo_transaccion.'),
    )
# print('columnas nuevas',df1['departamento_meta.'])

# df1['nivel.'] = df1['nivel_gobierno']+". "+df1['nivel_gobierno_nombre']##############################   ARREGLO
# df1['departamento_ejecutora.'] = df1['departamento_ejecutora']+". "+df1['departamento_ejecutora_nombre']##############################   ARREGLO
# df1['provincia_ejecutora.'] = df1['provincia_ejecutora']+". "+df1['provincia_ejecutora_nombre']##############################   ARREGLO
# df1['distrito_ejecutora.'] = df1['distrito_ejecutora']+". "+df1['distrito_ejecutora_nombre']##############################   ARREGLO

# df1.with_columns(pl.col('sector')=='36')

# df1 = df1.with_columns(
#     # (pl.when(df1['sector'] == '99').then(pl.lit('Poder Ejecutivo')).otherwise(pl.lit('')).alias('Rango')),)
#     (pl.when((pl.col('sector') == '01')
#              | (pl.col('sector')=='03')
#              )
#              .then(pl.lit('Poder Ejecutivo'))
#              .otherwise(pl.lit(''))
#              .alias('Rango'))
#     )
# print('rango',df1)

# df1 = df1.with_columns(
#     (pl.when((pl.col('sector') == '19')
#              | (pl.col('sector')=='20')
#              )
#              .then(pl.lit('Organismo Autónomo'))
#              .otherwise(pl.lit(False))
#              .alias('Rango'))    
# )

df1 = df1.with_columns(
    pl.col('sector').str.replace(r'01|03|05|06|07|08|09|10|11|12|13|16|26|35|36|37|38|39|40','Poder Ejecutivo').alias('Rango'),
)
# print('rango pe',df1)
df1 = df1.with_columns(
    pl.col('Rango').str.replace(r'19|20|21|22|24|31|32|33','Organismo Autónomo')#.alias('Rango'),
)
# print('rango oa',df1)
df1 = df1.with_columns(pl.col('Rango').str.replace(r'04|27','Poder Judicial'))#.alias('Rango'),
df1 = df1.with_columns(pl.col('Rango').str.replace(r'28','Poder Legislativo'))#.alias('Rango'),
df1 = df1.with_columns(pl.col('Rango').str.replace(r'98','Mancomunidad Regional'))#.alias('Rango'),
df1 = df1.with_columns(pl.col('Rango').str.replace(r'99','Gobierno Regional'))#.alias('Rango'),
# print('rango otros',df1)

df1 = df1.with_columns(
    pl.when(pl.col('nivel_gobierno') == 'M')
            .then(pl.lit('Gobierno Local'))
            .otherwise(df1['Rango'])
            .alias('Rango')
)
df1 = df1.with_columns(
    pl.when(pl.col('ejecutora').str.slice(0,2) == '97')
    .then(pl.lit('Mancomunidad Municipal'))
    .otherwise(df1['Rango'])
    .alias('Rango')
)

df1 = df1.with_columns(
    pl.when(pl.col('nivel_gobierno') == 'M')
            .then(pl.lit('M. Municipalidad'))
            .otherwise(df1['sector.'])
            .alias('sector.')
)
df1 = df1.with_columns(
    pl.when(pl.col('ejecutora').str.slice(0,2) == '97')
    .then(pl.lit('N. Mancomunidad'))
    .otherwise(df1['sector.'])
    .alias('sector.')
)

# print('Rango',df1['sector.'].unique())
# print(df111)
# df2 = df1.filter(pl.col('nivel_gobierno')== 'M')
# print('M2',df2['sector.'].unique())


# df2['Rango']

#################          PARA EVALUAR                  ###################
# df1 = df1.select(
#     pl.col('sector'),
#     # pl.when((pl.col('sector') == '99' | '01'))
#     pl.when((pl.col('sector') == '99') | (pl.col('sector')=='01'))
#     # pl.when((df1['sector'] == '99') | (df1['sector'] == '01'))
#              .then(pl.lit('Poder Ejecutivo'))
#              .otherwise(pl.lit(False))
#              .alias('Rango'),
#              )
# print('rango',df1)


# df1.loc[(df1['nivel_gobierno']== 'M'),'pliego.'] = df1['unidad_ejecutora']
print("--- %s seconds --- SECTORESS" % (time.time() - start_time))
# print('subido :', df1.shape)

df1 = df1.drop(['sector','sector_nombre','pliego','pliego_nombre','programa_ppto','programa_ppto_nombre',
          'producto_proyecto','producto_proyecto_nombre','actividad_accion_obra','actividad_accion_obra_nombre',
          'departamento_ejecutora','departamento_ejecutora_nombre','provincia_ejecutora','provincia_ejecutora_nombre', 'distrito_ejecutora', 'distrito_ejecutora_nombre',
          'ejecutora', 'ejecutora_nombre', 'nivel_gobierno', 'nivel_gobierno_nombre',
          'funcion','funcion_nombre', 'division_funcional','division_funcional_nombre','grupo_funcional','grupo_funcional_nombre',
          'finalidad','meta_nombre','rubro','rubro_nombre','tipo_recurso','tipo_recurso_nombre','tipo_act_proy',
          'departamento_meta','departamento_meta_nombre', 'categoria_gasto','categoria_gasto_nombre',
          'fuente_financiamiento','fuente_financiamiento_nombre','monto_devengado_anual',
          'monto_girado_anual',
          ])
print("--- %s seconds --- ELIMINAR COLUMNAS" % (time.time() - start_time))
# print('subido :', df1.shape)
# print(df1.groupby(['nivel.'])['monto_pim'].sum())
# print(df1)

# df1 = df1.rename({col: col.lower() for col in df1.columns})

df1 = df1.rename({
    'nivel.':'nivel',
    # 'pliego_nombre.':'pliego_nombre',
    # 'monto_girado_anual':'Girado_',
    'sector.':'sector',
    'pliego.':'pliego',
    'departamento_ejecutora.':'departamento_ejecutora',
    'provincia_ejecutora.':'provincia_ejecutora',
    'distrito_ejecutora.':'distrito_ejecutora',    
    # 'programa_pptal.':'programa_pptal',
    'producto_proyecto.':'producto_proyecto',
    'funcion.':'funcion',
    # 'division_fn.':'division_fn',
    # 'grupo_fn.':'grupo_fn',
    'finalidad.':'finalidad',
    'rubro.':'rubro',
    'tipo_recurso.':'tipo_recurso',
    'departamento_meta.':'departamento_meta',
    'fuente':'fuente_financ',
    'categoria_gasto.':'categoria_gasto',
    'tipo_transaccion':'tipo_transaccion_code',
    'tipo_transaccion.':'tipo_transaccion',
    'generica':'generica_code',
    'generica_nombre':'generica_name',
    'subgenerica':'subgenerica_code',
    'subgenerica_nombre':'subgenerica_name',
    'subgenerica_det':'subgenerica_det_code',
    'subgenerica_det_nombre':'subgenerica_det_name',
    'especifica':'especifica_code',
    'especifica_nombre':'especifica_name',
    'especifica_det':'especifica_det_code',
    'especifica_det_nombre':'especifica_det_name',
    'monto_pia':'mto_pia',
    'monto_pim':'mto_pim',
    'monto_certificado_anual':'mto_certificado',
    'monto_comprometido_anual':'mto_compro_anual',
    'monto_devengado_enero':'mto_devenga_01',
    'monto_devengado_febrero':'mto_devenga_02',
    'monto_devengado_marzo':'mto_devenga_03',
    'monto_devengado_abril':'mto_devenga_04',
    'monto_devengado_mayo':'mto_devenga_05',
    'monto_devengado_junio':'mto_devenga_06',
    'monto_devengado_julio':'mto_devenga_07',
    'monto_devengado_agosto':'mto_devenga_08',
    'monto_devengado_septiembre':'mto_devenga_09',
    'monto_devengado_octubre':'mto_devenga_10',
    'monto_devengado_noviembre':'mto_devenga_11',
    'monto_devengado_diciembre':'mto_devenga_12',
    })
print("--- %s seconds --- RENAME COLUMNS" % (time.time() - start_time))
# print(df1)
# print(df1.groupby(['nivel'])['mto_pim'].sum())

# df1 = df1.with_columns(
#     pl.col('tipo_act_obra_ac'), pl.col('activ_obra_accinv').apply(lambda x: x[:1])
# )

df1 = df1.with_columns((pl.col('activ_obra_accinv').str.slice(0,1)).alias('tipo_act_obra_ac').cast(pl.Utf8, strict=False))
df1 = df1.with_columns(
    pl.when(pl.col('tipo_act_obra_ac') == '0').then(pl.lit('0.SIN NOMBRE'))
    .when(pl.col('tipo_act_obra_ac') == '4').then(pl.lit('4.OBRA'))
    .when(pl.col('tipo_act_obra_ac') == '5').then(pl.lit('5.ACTIVIDAD'))
    .when(pl.col('tipo_act_obra_ac') == '6').then(pl.lit('6.ACCION DE INVERSION'))
    .otherwise(pl.col('tipo_act_obra_ac')).alias('tipo_act_obra_ac')
)

# df1 = df1.with_columns(pl.col('tipo_act_obra_ac').str.replace(r'0','0.SIN NOMBRE'))#.alias('Rango'),


####      ARREGLOS DE DATA ABIERTOS     #####

# df1 = df1.with_columns(
#     pl.when(pl.col('generica_code') == '0')
#     .then(pl.lit('0 Reserva de contingencia'))
#     .otherwise(pl.lit(''))
#     .alias('Concepto1')
# )
df1 = df1.with_columns((pl.col('generica_code').alias('Concepto1')),
                       (pl.col('generica_code').alias('Concepto2'))
                       )

# df1 = df1.with_columns(pl.col('generica_code').alias('Concepto2'))


df1 = df1.with_columns(pl.col('Concepto1').str.replace(r'0','0 Reserva de contingencia'))#.alias('Rango'),
df1 = df1.with_columns(pl.col('Concepto1').str.replace(r'1','1 Personal activo y pensionista'))#.alias('Rango'),
df1 = df1.with_columns(pl.col('Concepto1').str.replace(r'2','1 Personal activo y pensionista'))#.alias('Rango'),
df1 = df1.with_columns(pl.col('Concepto1').str.replace(r'3','2 Bienes y servicios'))#.alias('Rango'),
df1 = df1.with_columns(pl.col('Concepto1').str.replace(r'4','3 Transferencias a otras entidades'))#.alias('Rango'),
df1 = df1.with_columns(pl.col('Concepto1').str.replace(r'5','4 Otros gastos'))#.alias('Rango'),
df1 = df1.with_columns(pl.col('Concepto1').str.replace(r'6','5 Inversiones'))#.alias('Rango'),
df1 = df1.with_columns(pl.col('Concepto1').str.replace(r'7','6 Adquisición de activos financieros'))#.alias('Rango'),
df1 = df1.with_columns(pl.col('Concepto1').str.replace(r'8','7 Servicio de la Deuda'))#.alias('Rango'),
# print(df1)

df1 = df1.with_columns(
    (pl.col('tipo_transaccion_code') + pl.col("generica_code") + pl.col("subgenerica_code") + pl.col("subgenerica_det_code")).alias('clasif_4'),
    (pl.col('tipo_transaccion_code') + pl.col("generica_code") + pl.col("subgenerica_code") + pl.col("subgenerica_det_code") + pl.col("especifica_code")).alias('clasif_5'),
    (pl.col('tipo_transaccion_code') + pl.col("generica_code") + pl.col("subgenerica_code") + pl.col("subgenerica_det_code") + pl.col("especifica_code") + pl.col("especifica_det_code")).alias('clasif_6'),
    )

# print(df1['clasif_4'])
df1 = df1.with_columns(pl.col('Concepto2').str.replace(r'0','Reserva de contingencia'))#.alias('Rango'),
df1 = df1.with_columns(pl.col('Concepto2').str.replace(r'1','Personal activo'))#.alias('Rango'),
df1 = df1.with_columns(pl.col('Concepto2').str.replace(r'2','Personal pensionista'))#.alias('Rango'),
df1 = df1.with_columns(pl.col('Concepto2').str.replace(r'7','Adquisición de activos financieros'))#.alias('Rango'),
df1 = df1.with_columns(pl.col('Concepto2').str.replace(r'8','Servicio de la Deuda'))#.alias('Rango'),

valores_gasto_o = ['2311', '2312', '2313', '2315', '2316', '2322', '2326']
df1 = df1.with_columns(
    pl.when(pl.col('clasif_4').is_in(valores_gasto_o))
    .then(pl.lit('Gastos operativos fijos (Alim. Comb. Útil. Rep. S.Bas. S.Adm/Financ/Segur.)'))
    .otherwise(df1['Concepto2'])
    .alias('Concepto2')
)
df1 = df1.with_columns(pl.when(pl.col('clasif_4') == '2325').then(pl.lit('Alquileres')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when(pl.col('clasif_4') == '2328').then(pl.lit('Contratos Administrativos de Servicios')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when(pl.col('clasif_4') == '2314').then(pl.lit('Material explosivos')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when(pl.col('clasif_4') == '2321').then(pl.lit('Viajes')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when(pl.col('clasif_4') == '2323').then(pl.lit('Limpieza, seguridad y vigilancia')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when(pl.col('clasif_4') == '2329').then(pl.lit('Locadores de servicios')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when(pl.col('clasif_4') == '23210').then(pl.lit('Propinas a voluntarios')).otherwise(df1['Concepto2']).alias('Concepto2'))

valores_gasto_ob = ['2317', '2318', '2319', '23110', '23111', '23199']
df1 = df1.with_columns(
    pl.when(pl.col('clasif_4').is_in(valores_gasto_ob))
    .then(pl.lit('Otros bienes'))
    .otherwise(df1['Concepto2'])
    .alias('Concepto2')
)

df1 = df1.with_columns(pl.when(pl.col('clasif_4') == '2327').then(pl.lit('Otros servicios')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when(pl.col('clasif_4') == '2413').then(pl.lit('Transferencias diversas')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when(pl.col('clasif_4') == '2423').then(pl.lit('Transferencias por Control Concurrente')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when(pl.col('clasif_4') == '2521').then(pl.lit('Mantenimiento de Telecomunicaciones')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when(pl.col('clasif_4') == '2511').then(pl.lit('Subsidio a empresas públicas NF')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when(pl.col('clasif_4') == '2531').then(pl.lit('Subvenciones Financieras')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when(pl.col('clasif_4') == '2551').then(pl.lit('Sentencias judiciales')).otherwise(df1['Concepto2']).alias('Concepto2'))

valores_gasto_og = ['2512', '2522', '2541', '2542', '2543', '2552', '2561']
df1 = df1.with_columns(
    pl.when(pl.col('clasif_4').is_in(valores_gasto_og))
    .then(pl.lit('Otros gastos'))
    .otherwise(df1['Concepto2'])
    .alias('Concepto2')
)

########### clasif_5
valores_gasto_5om = ['23241', '23242', '23245', '23246', '23247', '23248', '232499']
df1 = df1.with_columns(
    pl.when(pl.col('clasif_5').is_in(valores_gasto_5om))
    .then(pl.lit('Otros mantenimientos'))
    .otherwise(df1['Concepto2'])
    .alias('Concepto2')
)

df1 = df1.with_columns(pl.when(pl.col('clasif_5') == '23243').then(pl.lit('Mantenimiento de Carreteras no concesionados')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when(pl.col('clasif_5') == '23244').then(pl.lit('Mantenimiento de Carreteras concesionados')).otherwise(df1['Concepto2']).alias('Concepto2'))

########### GASTOS NO CRITICOS LEY 32185 PTO 2025
valores_gasto_no_criticos = ['23211', '23212', '23271', '23272', '232221', '231211', '231212', '231213', '232241', '232243','231311','2327101', '2327102', '23271099']
df1 = df1.with_columns(
    pl.when(pl.col('clasif_5').is_in(valores_gasto_no_criticos))
    .then(pl.lit('Si'))
    .otherwise(pl.lit('No'))
    .alias('Gasto critico')
)
# valores_gasto_no_criticos = ['232221', '231211', '231212', '231213', '232241', '232243','231311','2327101','23271099']
df1 = df1.with_columns(
    pl.when(pl.col('clasif_6').is_in(valores_gasto_no_criticos))
    .then(pl.lit('Si'))
    .otherwise(df1['Gasto critico'])
    .alias('Gasto critico')
)

######### literales del gasto no critico
clasificadorgnc = {
    '232221': 'a)',
    '231211': 'c)',
    '231212': 'c)',
    '231213': 'c)',
    '232241': 'd)',
    '232243': 'd)',
    '231311': 'e)',
    '2327101': 'f)',
    '2327102': 'f)',
    '23271099': 'f)'
}

#############################################################
# Crear una lista de condiciones y resultados
conditions = [
    (pl.col('clasif_6') == key, pl.lit(value))
    for key, value in clasificadorgnc.items()
]
# Aplicar las condiciones en un bucle
expr = pl.when(conditions[0][0]).then(conditions[0][1])
for condition, result in conditions[1:]:
    expr = expr.when(condition).then(result)
# Finalizar con otherwise para mantener el valor original si no se cumple ninguna condición
df1 = df1.with_columns(
    expr.otherwise(pl.lit('No')).alias('GC numeral')
)
###############
clasificadorgnc = {
    '23211': 'b)',
    '23212': 'b)',
    '23271': 'g)',
    '23272': 'g)',
}
# Crear una lista de condiciones y resultados
conditions = [
    (pl.col('clasif_5') == key, pl.lit(value))
    for key, value in clasificadorgnc.items()
]
# Aplicar las condiciones en un bucle
expr = pl.when(conditions[0][0]).then(conditions[0][1])
for condition, result in conditions[1:]:
    expr = expr.when(condition).then(result)
# Finalizar con otherwise para mantener el valor original si no se cumple ninguna condición
df1 = df1.with_columns(
    expr.otherwise(df1['GC numeral']).alias('GC numeral')
)
# print(df1.filter(pl.col('Gasto critico')=='No'))

# sccc

######################################
print("--- %s seconds --- ARREGLO DE CONCEPTOS" % (time.time() - start_time))
# print('subido :', df1.shape)

df1 = df1.with_columns(pl.col('division_fn').alias('lamina 1'))
df1 = df1.with_columns(pl.col('division_fn').str.slice(0, 3).alias('division_fn3'))

valores_divfcom = ['037', '038']
df1 = df1.with_columns(
    pl.when(pl.col('division_fn3').is_in(valores_divfcom))
    .then(pl.lit('Comunicaciones'))
    .otherwise(df1['lamina 1'])
    .alias('lamina 1')
)
# #######                   es igual que el codigo anterior
# df1 = df1.with_columns(pl.when(pl.col('division_fn').str.slice(0, 3) == '038').then(pl.lit('Comunicaciones')).otherwise(df1['lamina 1']).alias('lamina 1'))
# df1 = df1.with_columns(pl.when(pl.col('division_fn').str.slice(0, 3) == '037').then(pl.lit('Comunicaciones')).otherwise(df1['lamina 1']).alias('lamina 1'))
# df1 = df1.with_columns(pl.when(pl.col('division_fn').str.slice(0, 3) == '045').then(pl.lit('Deporte')).otherwise(df1['lamina 1']).alias('lamina 1'))
# df1 = df1.with_columns(pl.when(pl.col('division_fn').str.slice(0, 3) == '046').then(pl.lit('Deporte')).otherwise(df1['lamina 1']).alias('lamina 1'))
df1 = df1.with_columns(pl.when(pl.col('division_fn').str.slice(0, 3) == '034').then(pl.lit('Ferroviario')).otherwise(df1['lamina 1']).alias('lamina 1'))
df1 = df1.with_columns(pl.when(pl.col('division_fn').str.slice(0, 3) == '035').then(pl.lit('Hidroviario')).otherwise(df1['lamina 1']).alias('lamina 1'))

valores_divfdep = ['045', '046']
df1 = df1.with_columns(
    pl.when(pl.col('division_fn3').is_in(valores_divfdep))
    .then(pl.lit('Deporte'))
    .otherwise(df1['lamina 1'])
    .alias('lamina 1')
)

valores_divfrv = ['033']
df1 = df1.with_columns(
    pl.when(pl.col('division_fn3').is_in(valores_divfrv))
    .then(pl.lit('Red vial'))
    .otherwise(df1['lamina 1'])
    .alias('lamina 1')
)

############################################################     DOBLE CONSULTA     #####################################################################
df1 = df1.with_columns(
    pl.when((pl.col('division_fn').str.slice(0, 3) == '033') & (pl.col('grupo_fn').str.slice(0, 4) == '0069'))
    .then(pl.lit('Gestión y otros'))
    .otherwise(df1['lamina 1'])
    .alias('lamina 1'))


# df1 = df1.with_columns([
#     pl.when(
#         pl.col('division_fn').str.slice(0, 3) == '033')
#         .then(pl.lit('Comunicaciones'))
#         .otherwise(df1['lamina 1'])
#         .alias('lamina 1'),
#     pl.when(
#         pl.col('division_fn').str.slice(0, 3) == '037')
#         .then(pl.lit('Comunicaciones'))
#         .otherwise(df1['lamina 1'])
#         .alias('lamina 1')
# ])


df1 = df1.with_columns(pl.when(pl.col('division_fn3') == '036').then(pl.lit('Urbano')).otherwise(df1['lamina 1']).alias('lamina 1'))
df1 = df1.with_columns(pl.when(pl.col('division_fn3') == '032').then(pl.lit('Aéreo')).otherwise(df1['lamina 1']).alias('lamina 1'))

valores_divfges = ['004', '006', '018', '052', '016', '055', '015']
df1 = df1.with_columns(
    pl.when(pl.col('division_fn3').is_in(valores_divfges))
    .then(pl.lit('Gestión y otros'))
    .otherwise(df1['lamina 1'])
    .alias('lamina 1')
)
df1 = df1.drop(['division_fn3'])

# df1 = df1.with_columns(pl.when(pl.col('division_fn').str.slice(0, 3) == '004').then(pl.lit('Gestión y otros')).otherwise(df1['lamina 1']).alias('lamina 1'))
# df1 = df1.with_columns(pl.when(pl.col('division_fn').str.slice(0, 3) == '006').then(pl.lit('Gestión y otros')).otherwise(df1['lamina 1']).alias('lamina 1'))
# df1 = df1.with_columns(pl.when(pl.col('division_fn').str.slice(0, 3) == '018').then(pl.lit('Gestión y otros')).otherwise(df1['lamina 1']).alias('lamina 1'))
# df1 = df1.with_columns(pl.when(pl.col('division_fn').str.slice(0, 3) == '052').then(pl.lit('Gestión y otros')).otherwise(df1['lamina 1']).alias('lamina 1'))
# df1 = df1.with_columns(pl.when(pl.col('division_fn').str.slice(0, 3) == '016').then(pl.lit('Gestión y otros')).otherwise(df1['lamina 1']).alias('lamina 1'))
# df1 = df1.with_columns(pl.when(pl.col('division_fn').str.slice(0, 3) == '055').then(pl.lit('Gestión y otros')).otherwise(df1['lamina 1']).alias('lamina 1'))
# df1 = df1.with_columns(pl.when(pl.col('division_fn').str.slice(0, 3) == '015').then(pl.lit('Gestión y otros')).otherwise(df1['lamina 1']).alias('lamina 1'))
print("--- %s seconds --- ARREGLO DE DIVISION FUNCIONAL" % (time.time() - start_time))
###########################  link images   ##########################

df1 = df1.with_columns(images = pl.lit(''))  ######  CREA COLUMNA
df1 = df1.with_columns(pl.when(pl.col('lamina 1').str.contains('Urbano')).then(pl.lit('https://cdn.pixabay.com/photo/2013/03/13/17/49/bus-93219_1280.png')).otherwise(df1['images']).alias('images'))
df1 = df1.with_columns(pl.when(pl.col('lamina 1').str.contains('Deporte')).then(pl.lit('https://cdn.pixabay.com/photo/2013/07/12/15/36/handball-150163_1280.png')).otherwise(df1['images']).alias('images'))
df1 = df1.with_columns(pl.when(pl.col('lamina 1').str.contains('Gestión y otros')).then(pl.lit('https://cdn.pixabay.com/photo/2013/07/12/19/04/building-154304_1280.png')).otherwise(df1['images']).alias('images'))
df1 = df1.with_columns(pl.when(pl.col('lamina 1').str.contains('Comunicaciones')).then(pl.lit('https://cdn.pixabay.com/photo/2013/07/12/19/32/antenna-154946_1280.png')).otherwise(df1['images']).alias('images'))
df1 = df1.with_columns(pl.when(pl.col('lamina 1').str.contains('Aéreo')).then(pl.lit('https://cdn.pixabay.com/photo/2014/03/24/17/19/airplane-295396_1280.png')).otherwise(df1['images']).alias('images'))
df1 = df1.with_columns(pl.when(pl.col('lamina 1').str.contains('Red vial')).then(pl.lit('https://cdn.pixabay.com/photo/2014/04/02/10/19/truck-303460_1280.png')).otherwise(df1['images']).alias('images'))
df1 = df1.with_columns(pl.when(pl.col('lamina 1').str.contains('Ferroviario')).then(pl.lit('https://cdn.pixabay.com/photo/2014/04/03/10/11/high-speed-train-310079_1280.png')).otherwise(df1['images']).alias('images'))
df1 = df1.with_columns(pl.when(pl.col('lamina 1').str.contains('Hidroviario')).then(pl.lit('https://cdn.pixabay.com/photo/2015/04/25/05/42/boat-738679_1280.png')).otherwise(df1['images']).alias('images'))
print("--- %s seconds --- LINK" % (time.time() - start_time))

###########################  MANTENIMIENTO   ##########################
##################           ESTE METODO PARA CREA COLUMNAS NOMBRES DEBE SER SIN ESPACIO          ##################
df1 = df1.with_columns(lamina6 = pl.lit('').cast(pl.Utf8, strict=False) ) 
df1 = df1.rename({'lamina6':'lamina 6'})
df1 = df1.with_columns(
    pl.when((pl.col('producto_proyecto').str.slice(0,7) == '3000131') & (pl.col('clasif_5') == '23243'))
    .then(pl.lit('2 Mantenimiento de Vías No Concesionadas'))
    .otherwise(df1['lamina 6'])
    .alias('lamina 6'))
df1 = df1.with_columns(pl.when((pl.col('producto_proyecto').str.slice(0,7) == '3000131') & (pl.col('clasif_5') == '23244')).then(pl.lit('1 Mantenimiento de Vías Concesionadas')).otherwise(df1['lamina 6']).alias('lamina 6'))
df1 = df1.with_columns(pl.when((pl.col('producto_proyecto').str.slice(0,7) == '3000132') & (pl.col('clasif_5') == '23243')).then(pl.lit('3 Mantenimiento PRO REGION')).otherwise(df1['lamina 6']).alias('lamina 6'))
df1 = df1.with_columns(pl.when((pl.col('activ_obra_accinv').str.slice(0,7) == '5004885')).then(pl.lit('4 Mantenimiento Red Dorsal')).otherwise(df1['lamina 6']).alias('lamina 6'))
df1 = df1.with_columns(pl.when((pl.col('activ_obra_accinv').str.slice(0,7) == '5006283')).then(pl.lit('5 Mantenimiento de proyectos regionales')).otherwise(df1['lamina 6']).alias('lamina 6'))
# df1 = df1.with_columns(pl.col('lamina 6').fill_null(pl.lit('Por definir')))
df1 = df1.with_columns(pl.when((pl.col('lamina 6') == '')).then(pl.lit('Por definir')).otherwise(df1['lamina 6']).alias('lamina 6'))

# print(df1['lamina 6'].unique())
print("--- %s seconds --- MANTENIMIENTO" % (time.time() - start_time))

################              CLASIFICADORES                 ######################

df1 = df1.with_columns(pl.when((pl.col('subgenerica_code')== '0')).then(pl.lit(' 0')).otherwise(df1['subgenerica_code']).alias('subgenerica_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_code')== '1')).then(pl.lit(' 1')).otherwise(df1['subgenerica_code']).alias('subgenerica_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_code')== '2')).then(pl.lit(' 2')).otherwise(df1['subgenerica_code']).alias('subgenerica_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_code')== '3')).then(pl.lit(' 3')).otherwise(df1['subgenerica_code']).alias('subgenerica_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_code')== '4')).then(pl.lit(' 4')).otherwise(df1['subgenerica_code']).alias('subgenerica_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_code')== '5')).then(pl.lit(' 5')).otherwise(df1['subgenerica_code']).alias('subgenerica_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_code')== '6')).then(pl.lit(' 6')).otherwise(df1['subgenerica_code']).alias('subgenerica_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_code')== '7')).then(pl.lit(' 7')).otherwise(df1['subgenerica_code']).alias('subgenerica_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_code')== '8')).then(pl.lit(' 8')).otherwise(df1['subgenerica_code']).alias('subgenerica_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_code')== '9')).then(pl.lit(' 9')).otherwise(df1['subgenerica_code']).alias('subgenerica_code'))


df1 = df1.with_columns(pl.when((pl.col('subgenerica_det_code')== '0')).then(pl.lit(' 0')).otherwise(df1['subgenerica_det_code']).alias('subgenerica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_det_code')== '1')).then(pl.lit(' 1')).otherwise(df1['subgenerica_det_code']).alias('subgenerica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_det_code')== '2')).then(pl.lit(' 2')).otherwise(df1['subgenerica_det_code']).alias('subgenerica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_det_code')== '3')).then(pl.lit(' 3')).otherwise(df1['subgenerica_det_code']).alias('subgenerica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_det_code')== '4')).then(pl.lit(' 4')).otherwise(df1['subgenerica_det_code']).alias('subgenerica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_det_code')== '5')).then(pl.lit(' 5')).otherwise(df1['subgenerica_det_code']).alias('subgenerica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_det_code')== '6')).then(pl.lit(' 6')).otherwise(df1['subgenerica_det_code']).alias('subgenerica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_det_code')== '7')).then(pl.lit(' 7')).otherwise(df1['subgenerica_det_code']).alias('subgenerica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_det_code')== '8')).then(pl.lit(' 8')).otherwise(df1['subgenerica_det_code']).alias('subgenerica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('subgenerica_det_code')== '9')).then(pl.lit(' 9')).otherwise(df1['subgenerica_det_code']).alias('subgenerica_det_code'))
# df1 = df1.with_columns(pl.when((pl.col('subgenerica_det_code')== '10')).then(pl.lit('10')).otherwise(df1['subgenerica_det_code']).alias('subgenerica_det_code'))
# df1 = df1.with_columns(pl.when((pl.col('subgenerica_det_code')== '11')).then(pl.lit('11')).otherwise(df1['subgenerica_det_code']).alias('subgenerica_det_code'))
# df1 = df1.with_columns(pl.when((pl.col('subgenerica_det_code')== '12')).then(pl.lit('12')).otherwise(df1['subgenerica_det_code']).alias('subgenerica_det_code'))
# df1 = df1.with_columns(pl.when((pl.col('subgenerica_det_code')== '13')).then(pl.lit('13')).otherwise(df1['subgenerica_det_code']).alias('subgenerica_det_code'))
# df1 = df1.with_columns(pl.when((pl.col('subgenerica_det_code')== '99')).then(pl.lit('99')).otherwise(df1['subgenerica_det_code']).alias('subgenerica_det_code'))

df1 = df1.with_columns(pl.when((pl.col('especifica_code')== '0')).then(pl.lit(' 0')).otherwise(df1['especifica_code']).alias('especifica_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_code')== '1')).then(pl.lit(' 1')).otherwise(df1['especifica_code']).alias('especifica_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_code')== '2')).then(pl.lit(' 2')).otherwise(df1['especifica_code']).alias('especifica_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_code')== '3')).then(pl.lit(' 3')).otherwise(df1['especifica_code']).alias('especifica_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_code')== '4')).then(pl.lit(' 4')).otherwise(df1['especifica_code']).alias('especifica_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_code')== '5')).then(pl.lit(' 5')).otherwise(df1['especifica_code']).alias('especifica_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_code')== '6')).then(pl.lit(' 6')).otherwise(df1['especifica_code']).alias('especifica_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_code')== '7')).then(pl.lit(' 7')).otherwise(df1['especifica_code']).alias('especifica_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_code')== '8')).then(pl.lit(' 8')).otherwise(df1['especifica_code']).alias('especifica_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_code')== '9')).then(pl.lit(' 9')).otherwise(df1['especifica_code']).alias('especifica_code'))

df1 = df1.with_columns(pl.when((pl.col('especifica_det_code')== '0')).then(pl.lit(' 0')).otherwise(df1['especifica_det_code']).alias('especifica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_det_code')== '1')).then(pl.lit(' 1')).otherwise(df1['especifica_det_code']).alias('especifica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_det_code')== '2')).then(pl.lit(' 2')).otherwise(df1['especifica_det_code']).alias('especifica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_det_code')== '3')).then(pl.lit(' 3')).otherwise(df1['especifica_det_code']).alias('especifica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_det_code')== '4')).then(pl.lit(' 4')).otherwise(df1['especifica_det_code']).alias('especifica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_det_code')== '5')).then(pl.lit(' 5')).otherwise(df1['especifica_det_code']).alias('especifica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_det_code')== '6')).then(pl.lit(' 6')).otherwise(df1['especifica_det_code']).alias('especifica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_det_code')== '7')).then(pl.lit(' 7')).otherwise(df1['especifica_det_code']).alias('especifica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_det_code')== '8')).then(pl.lit(' 8')).otherwise(df1['especifica_det_code']).alias('especifica_det_code'))
df1 = df1.with_columns(pl.when((pl.col('especifica_det_code')== '9')).then(pl.lit(' 9')).otherwise(df1['especifica_det_code']).alias('especifica_det_code'))
# print('subgenerica_det_code :\n',df1['subgenerica_det_code'].unique())

print("--- %s seconds --- CLASIFICADORES" % (time.time() - start_time))

df1 = df1.with_columns(
    pl.concat_str([
        pl.col("tipo_transaccion_code").str.strip_chars(),
        pl.lit(' '),
        pl.col("generica_code"),
        pl.lit(' '),
        pl.col("subgenerica_code").str.strip_chars(),
        pl.lit(' '),
        pl.col("subgenerica_det_code").str.strip_chars(),
        pl.lit(' '),
        pl.col("especifica_code").str.strip_chars(),
        pl.lit(' '),
        pl.col("especifica_det_code").str.strip_chars(),
        # pl.col("edad").cast(pl.Utf8),
    ]).alias("clasif_m")
)

df1 = df1.with_columns(
    # (pl.col('tipo_transaccion_code') + ". " + pl.col("generica_code")+ ". " + pl.col("subgenerica_code")+ ". " + pl.col("subgenerica_det_code")+ ". " + pl.col("especifica_code")+ ". " + pl.col("especifica_det_code")).alias('clasif_m'),
    # (pl.col('clasif_m') + ". " + pl.col("especifica_det_name")).alias('clasificador_m'),
    (pl.col('tipo_transaccion_code') + pl.col("generica_code")+ " " + pl.col("generica_name")).alias('Genérica'),
    (pl.col('tipo_transaccion_code') + pl.col("generica_code")+ pl.col("subgenerica_code")+ " " + pl.col("subgenerica_name")).alias('Sub generica'),
    (pl.col('tipo_transaccion_code') + pl.col("generica_code")+ pl.col("subgenerica_code")+ pl.col("subgenerica_det_code")+ " " + pl.col("subgenerica_det_name")).alias('Subgenerica det'),
    (pl.col('tipo_transaccion_code') + pl.col("generica_code")+ pl.col("subgenerica_code")+ pl.col("subgenerica_det_code")+ pl.col("especifica_code")+ " " + pl.col("especifica_name")).alias('Especifica.'),
    (pl.col('tipo_transaccion_code') + pl.col("generica_code")+ pl.col("subgenerica_code")+ pl.col("subgenerica_det_code")+ pl.col("especifica_code")+ pl.col("especifica_det_code")).alias('Especifica det'),
    (pl.col('tipo_transaccion_code') + pl.col("generica_code")+ pl.col("subgenerica_code")+ pl.col("subgenerica_det_code")+ pl.col("especifica_code")+ pl.col("especifica_det_code")+ " "+pl.col("especifica_det_name")).alias('Clasificador'),
    )
df1 = df1.with_columns((pl.col('clasif_m') + ". " + pl.col("especifica_det_name")).alias('clasificador_m'))
# print(df1['clasif_m'].unique())

# print('subido :', df1.shape)
# print(df1.groupby(['nivel'])['mto_pim'].sum())


# df1.drop_duplicates(inplace=True)
# df1 = df1.merge(dfc, left_on='Especifica det', right_on='clasif')
# df1.drop(['id_clasificador','id_clasificador_2','origen_clasificador','estado','es_presupuestal','tipo_transaccion_code','tipo_transaccion_nombre','generica_code','generica_name','subgenerica_code','subgenerica_name','subgenerica_det_code','subgenerica_det_name','especifica_code','especifica_name','especifica_det_code','especifica_det_name','clasif'],axis=1, inplace=True)

# df1 = df1.merge(dfc, how='left', on='Especifica det')
# print(df1)
df1 = df1.join(dfc, on='Especifica det', how=('left'))



# print('data open',df1['clasif_m'].unique())
# print(df1)

df1 = df1.drop(['id_clasificador','id_clasificador_2','origen_clasificador','estado','es_presupuestal','tipo_transaccion_code','tipo_transaccion_nombre','generica_code','generica_name','subgenerica_code','subgenerica_name','subgenerica_det_code','subgenerica_det_name','especifica_code','especifica_name','especifica_det_code','especifica_det_name',])
# print(df1)
# df1 = df1.unique()
# print(df1)
print("--- %s seconds --- CLASIFICADORES 2" % (time.time() - start_time))
#########################         CALCULOS     #########################
# print(df1.groupby(['nivel'])['mto_pim'].sum())

df1 = df1.with_columns(
    # Compromiso_=pl.sum_horizontal('mto_at_comp_01', 'mto_at_comp_02', 'mto_at_comp_03', 'mto_at_comp_04', 'mto_at_comp_05', 'mto_at_comp_06', 'mto_at_comp_07', 'mto_at_comp_08', 'mto_at_comp_09', 'mto_at_comp_10', 'mto_at_comp_11', 'mto_at_comp_12'),
    Devengadoo=pl.sum_horizontal('mto_devenga_01', 'mto_devenga_02', 'mto_devenga_03', 'mto_devenga_04', 'mto_devenga_05', 'mto_devenga_06', 'mto_devenga_07', 'mto_devenga_08', 'mto_devenga_09', 'mto_devenga_10', 'mto_devenga_11', 'mto_devenga_12'),
    # Girado_=pl.sum_horizontal('mto_girado_01', 'mto_girado_02', 'mto_girado_03', 'mto_girado_04', 'mto_girado_05', 'mto_girado_06', 'mto_girado_07', 'mto_girado_08', 'mto_girado_09', 'mto_girado_10', 'mto_girado_11', 'mto_girado_12'),
)

# df1 = df1.rename({'Compromiso_':'Compromiso.'})
df1 = df1.rename({'Devengadoo':'Devengado.'})
# df1 = df1.rename({'Girado_':'Girado.'})

# prn_ejec = (
#     df1.group_by(['nivel',])
#     .agg(pl.col('mto_pia').sum().cast(pl.Int64), pl.col('mto_pim').sum().cast(pl.Int64), (pl.col('Devengado.').sum().cast(pl.Int64)))
#     # .agg(pl.col('mto_pim').sum().cast(pl.Int64), (pl.col('Devengado.').sum().cast(pl.Float64).round(2)))
#     .sort(['nivel']))
# print(prn_ejec)


df1 = df1.with_columns((pl.col('mto_pim') - pl.col('mto_certificado')).alias('Saldo Cert.'))
# df1 = df1.with_columns((pl.col('mto_certificado') - pl.col('mto_compro_anual')).alias('Saldo Compromiso anual.'))
# df1 = df1.with_columns((pl.col('mto_compro_anual') - pl.col('Compromiso.')).alias('Saldo Compromiso.')),
# df1 = df1.with_columns((pl.col('Compromiso.') - pl.col('Devengado.')).alias('Saldo Devengado.')),
# df1 = df1.with_columns((pl.col('Devengado.') - pl.col('Girado.')).alias('Saldo Girado.')),
df1 = df1.with_columns((pl.col('mto_pim') - pl.col('Devengado.')).alias('Saldo Dev.'))

# df1 = df1.with_columns(
#     # (pl.col('mto_at_comp_01') + pl.col('mto_at_comp_02')+pl.col('mto_at_comp_03')+pl.col('mto_at_comp_04')+pl.col('mto_at_comp_05')+pl.col('mto_at_comp_06')+pl.col('mto_at_comp_07')+pl.col('mto_at_comp_08')+pl.col('mto_at_comp_09')+pl.col('mto_at_comp_10')+pl.col('mto_at_comp_11')+pl.col('mto_at_comp_12').alias('Compromiso.')),
#     (pl.col('mto_devenga_01') + pl.col('mto_devenga_02')+pl.col('mto_devenga_03')+pl.col('mto_devenga_04')+pl.col('mto_devenga_05')+pl.col('mto_devenga_06')+pl.col('mto_devenga_07')+pl.col('mto_devenga_08')+pl.col('mto_devenga_09')+pl.col('mto_devenga_10')+pl.col('mto_devenga_11')+pl.col('mto_devenga_12').alias('Devengado.')),
#     # (pl.col('mto_girado_01') + pl.col('mto_girado_02')+pl.col('mto_girado_03')+pl.col('mto_girado_04')+pl.col('mto_girado_05')+pl.col('mto_girado_06')+pl.col('mto_girado_07')+pl.col('mto_girado_08')+pl.col('mto_girado_09')+pl.col('mto_girado_10')+pl.col('mto_girado_11')+pl.col('mto_girado_12').alias('Girado.')),
#     # (pl.col('mto_pim') - pl.col('mto_certificado').alias('Saldo Cert.')),
#     # (pl.col('mto_certificado') - pl.col('mto_compro_anual').alias('Saldo Compromiso anual.')),
#     # (pl.col('mto_compro_anual') - pl.col('Compromiso.').alias('Saldo Compromiso.')),
#     # (pl.col('Compromiso.') - pl.col('Devengado.').alias('Saldo Devengado.')),
#     # (pl.col('Devengado.') - pl.col('Girado.').alias('Saldo Girado.')),
#     (pl.col('mto_pim') - pl.col('Devengado.').alias('Saldo Dev.')),
# )



# # df1['Compromiso.'] = df1['mto_at_comp_01'] + df1['mto_at_comp_02']+df1['mto_at_comp_03']+df1['mto_at_comp_04']+df1['mto_at_comp_05']+df1['mto_at_comp_06']+df1['mto_at_comp_07']+df1['mto_at_comp_08']+df1['mto_at_comp_09']+df1['mto_at_comp_10']+df1['mto_at_comp_11']+df1['mto_at_comp_12']
# df1['Devengado.'] = df1['mto_devenga_01'] + df1['mto_devenga_02']+df1['mto_devenga_03']+df1['mto_devenga_04']+df1['mto_devenga_05']+df1['mto_devenga_06']+df1['mto_devenga_07']+df1['mto_devenga_08']+df1['mto_devenga_09']+df1['mto_devenga_10']+df1['mto_devenga_11']+df1['mto_devenga_12']
# # df1['Girado.'] = df1['mto_girado_01'] + df1['mto_girado_02']+df1['mto_girado_03']+df1['mto_girado_04']+df1['mto_girado_05']+df1['mto_girado_06']+df1['mto_girado_07']+df1['mto_girado_08']+df1['mto_girado_09']+df1['mto_girado_10']+df1['mto_girado_11']+df1['mto_girado_12']

# df1['Saldo Cert.'] = df1['mto_pim'] - df1['mto_certificado']
# # df1['Saldo Compromiso anual.'] = df1['mto_certificado'] - df1['mto_compro_anual']
# # df1['Saldo Compromiso.'] = df1['mto_compro_anual'] - df1['Compromiso.']
# # df1['Saldo Devengado.'] = df1['Compromiso.'] - df1['Devengado.']
# # df1['Saldo Girado.'] = df1['Devengado.'] - df1['Girado.']
# df1['Saldo Dev.'] = df1['mto_pim'] - df1['Devengado.']
print("--- %s seconds --- CALCULOS" % (time.time() - start_time))
#########################         GASTOS     #########################

# if isinstance(df1, pl.DataFrame):
#     print("df1 es un DataFrame de Polars.")
# else:
#     print("df1 NO es un DataFrame de Polars.")


# df1 = df1.with_columns(lamina6 = pl.lit('')) 
df1 = df1.with_columns([pl.col('producto_proyecto').str.slice(0, 7).alias('dnpp')])
df1 = df1.with_columns(gastos = pl.lit('')) 
df1 = df1.with_columns(pl.when((pl.col('activ_obra_accinv').str.slice(0,7) == '5006269')).then(pl.lit('b. Covid-19')).otherwise(df1['gastos']).alias('gastos'))
df1 = df1.with_columns(pl.when((pl.col('activ_obra_accinv').str.slice(0,7) == '5002003')).then(pl.lit('c. Mundial')).otherwise(df1['gastos']).alias('gastos'))
df1 = df1.with_columns(pl.when((pl.col('dnpp').str.slice(0,7) == '2563242')).then(pl.lit('c. Mundial')).otherwise(df1['gastos']).alias('gastos'))
df1 = df1.with_columns(pl.when((pl.col('dnpp').str.slice(0,7) == '2563335')).then(pl.lit('c. Mundial')).otherwise(df1['gastos']).alias('gastos'))
df1 = df1.with_columns(pl.when((pl.col('dnpp').str.slice(0,7) == '2563048')).then(pl.lit('c. Mundial')).otherwise(df1['gastos']).alias('gastos'))
df1 = df1.with_columns(pl.when((pl.col('dnpp').str.slice(0,7) == '2570801')).then(pl.lit('c. Mundial')).otherwise(df1['gastos']).alias('gastos'))
df1 = df1.with_columns(pl.when((pl.col('dnpp').str.slice(0,7) == '2564414')).then(pl.lit('c. Mundial')).otherwise(df1['gastos']).alias('gastos'))
df1 = df1.with_columns(pl.when((pl.col('dnpp').str.slice(0,7) == '2564412')).then(pl.lit('c. Mundial')).otherwise(df1['gastos']).alias('gastos'))
df1 = df1.with_columns(pl.when((pl.col('gastos') == '')).then(pl.lit('a. Institucional')).otherwise(df1['gastos']).alias('gastos'))

# df1 = df1.with_columns(pl.col('producto_proyecto').alias('dnpp'))              ###############     DUPLICA COLUMNA

# print(df1)

# df1['dnpp'] = df1['producto_proyecto'].str[:7]
# df1['gastos']=''
# df1.loc[df1['activ_obra_accinv'].str[:7] == '5002003', 'gastos'] = 'c. Mundial'
# df1.loc[df1['dnpp'] == '2563242', 'gastos'] = 'c. Mundial'
# df1.loc[df1['dnpp'] == '2563335', 'gastos'] = 'c. Mundial'
# df1.loc[df1['dnpp'] == '2563048', 'gastos'] = 'c. Mundial'
# df1.loc[df1['dnpp'] == '2570801', 'gastos'] = 'c. Mundial'
# df1.loc[df1['dnpp'] == '2564414', 'gastos'] = 'c. Mundial'
# df1.loc[df1['dnpp'] == '2564412', 'gastos'] = 'c. Mundial'
# df1.loc[df1['activ_obra_accinv'].str[:7] == '5006269', 'gastos'] = 'b. Covid-19'

# df1['gastos'] = df1['gastos'].fillna(value='a. Institucional')
# df1.loc[df1['gastos'] == '', 'gastos'] = 'a. Institucional'

#########################         ENTIDADES     #########################
df1 = df1.with_columns(pl.col('departamento_meta').str.replace('07. PROVINCIA CONSTITUCIONAL DEL CALLAO','07. CALLAO'))#.alias('Rango'),
df1 = df1.with_columns(pl.col('departamento_ejecutora').str.replace('07. PROVINCIA CONSTITUCIONAL DEL CALLAO','07. CALLAO'))#.alias('Rango'),

df1 = df1.with_columns(UE = pl.lit(''))
df1 = df1.with_columns(pliego_Nombre_abrev = pl.lit(''))
df1 = df1.with_columns(Tipo_de_gasto = pl.lit(''))
df1 = df1.rename({'pliego_Nombre_abrev':'pliego Nombre abrev', 'Tipo_de_gasto':'Tipo de gasto' })

df1 = df1.with_columns(pl.col('sec_ejec').cast(pl.Int64, strict=False))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036') & (pl.col('sec_ejec') == 1072)).then(pl.lit('001 MTC')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036') & (pl.col('sec_ejec') == 1078)).then(pl.lit('007 PVN')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036') & (pl.col('sec_ejec') == 1250)).then(pl.lit('010 PVD')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036') & (pl.col('sec_ejec') == 1260)).then(pl.lit('011 FITEL')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036') & (pl.col('sec_ejec') == 1338)).then(pl.lit('012 AATE')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036') & (pl.col('sec_ejec') == 1669)).then(pl.lit('013 LEGADO')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036') & (pl.col('sec_ejec') == 1720)).then(pl.lit('014 PRONATEL')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036') & (pl.col('sec_ejec') == 1747)).then(pl.lit('015 TRUJILLO')).otherwise(df1['UE']).alias('UE'))

df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '202') & (pl.col('sec_ejec') == 1346)).then(pl.lit('001 SUTRAN')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '203') & (pl.col('sec_ejec') == 1717)).then(pl.lit('001 ATU')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '214') & (pl.col('sec_ejec') == 1205)).then(pl.lit('001 APN')).otherwise(df1['UE']).alias('UE'))

df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036')).then(pl.lit('036 MTC')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '202')).then(pl.lit('202 SUTRAN')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '203')).then(pl.lit('203 ATU')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '214')).then(pl.lit('214 APN')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))

df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '012')).then(pl.lit('012 MTPE')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '121')).then(pl.lit('121 SUNAFIL')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '012') & (pl.col('sec_ejec') == 154)).then(pl.lit('001 MT-OGA')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '012') & (pl.col('sec_ejec') == 993)).then(pl.lit('002 J. PRODUCTIVOS')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '012') & (pl.col('sec_ejec') == 1066)).then(pl.lit('005 TRABAJA PERU')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '121') & (pl.col('sec_ejec') == 1510)).then(pl.lit('001 SUNAFIL')).otherwise(df1['UE']).alias('UE'))

df1 = df1.with_columns(pl.when(pl.col('tipo_prod_proy') == '0.OTROS').then(pl.lit('Otros')).otherwise(df1['Tipo de gasto']).alias('Tipo de gasto'))
df1 = df1.with_columns(pl.when((pl.col('tipo_prod_proy') == '2.PROYECTO')).then(pl.lit('Inversiones')).otherwise(df1['Tipo de gasto']).alias('Tipo de gasto'))
df1 = df1.with_columns(pl.when((pl.col('tipo_prod_proy') == '3.PRODUCTO')).then(pl.lit('Actividades')).otherwise(df1['Tipo de gasto']).alias('Tipo de gasto'))

print("--- %s seconds --- ENTIDADES" % (time.time() - start_time))

#####################      ENTIDADES UES GRP      #######################

# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '440').then(pl.lit('GORE AMAZONAS')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '441').then(pl.lit('GORE ANCASH')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '442').then(pl.lit('GORE APURIMAC')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '443').then(pl.lit('GORE AREQUIPA')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '444').then(pl.lit('GORE AYACUCHO')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '445').then(pl.lit('GORE CAJAMARCA')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '446').then(pl.lit('GORE CUSCO')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '447').then(pl.lit('GORE HUANCAVELICA')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '448').then(pl.lit('GORE HUANUCO')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '449').then(pl.lit('GORE ICA')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '450').then(pl.lit('GORE JUNIN')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '451').then(pl.lit('GORE LA LIBERTAD')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '452').then(pl.lit('GORE LAMBAYEQUE')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '453').then(pl.lit('GORE LORETO')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '454').then(pl.lit('GORE MADRE DE DIOS')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '455').then(pl.lit('GORE MOQUEGUA')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '456').then(pl.lit('GORE PASCO')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '457').then(pl.lit('GORE PIURA')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '458').then(pl.lit('GORE PUNO')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '459').then(pl.lit('GORE SAN MARTIN')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '460').then(pl.lit('GORE TACNA')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '461').then(pl.lit('GORE TUMBES')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '462').then(pl.lit('GORE UCAYALI')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '463').then(pl.lit('GORE LIMA')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '464').then(pl.lit('GORE CALLAO')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '465').then(pl.lit('GORE LIMA MML')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
# print(df1)

# Diccionario de pliegos y sus nombres abreviados
pliegos_dict = {
    '440': 'GORE AMAZONAS',
    '441': 'GORE ANCASH',
    '442': 'GORE APURIMAC',
    '443': 'GORE AREQUIPA',
    '444': 'GORE AYACUCHO',
    '445': 'GORE CAJAMARCA',
    '446': 'GORE CUSCO',
    '447': 'GORE HUANCAVELICA',
    '448': 'GORE HUANUCO',
    '449': 'GORE ICA',
    '450': 'GORE JUNIN',
    '451': 'GORE LA LIBERTAD',
    '452': 'GORE LAMBAYEQUE',
    '453': 'GORE LORETO',
    '454': 'GORE MADRE DE DIOS',
    '455': 'GORE MOQUEGUA',
    '456': 'GORE PASCO',
    '457': 'GORE PIURA',
    '458': 'GORE PUNO',
    '459': 'GORE SAN MARTIN',
    '460': 'GORE TACNA',
    '461': 'GORE TUMBES',
    '462': 'GORE UCAYALI',
    '463': 'GORE LIMA',
    '464': 'GORE CALLAO',
    '465': 'GORE LIMA MML'
}
# Crear una lista de condiciones y resultados
conditions = [
    (pl.col('pliego').str.slice(0, 3) == key, pl.lit(value))
    for key, value in pliegos_dict.items()
]
# Aplicar las condiciones en un bucle
expr = pl.when(conditions[0][0]).then(conditions[0][1])
for condition, result in conditions[1:]:
    expr = expr.when(condition).then(result)
# Finalizar con otherwise para mantener el valor original si no se cumple ninguna condición
df1 = df1.with_columns(
    expr.otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev')
)
# print(df1['pliego Nombre abrev'])

df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '440').then(pl.lit('GORE AMAZONAS')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '441').then(pl.lit('GORE ANCASH')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '442').then(pl.lit('GORE APURIMAC')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '443').then(pl.lit('GORE AREQUIPA')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '444').then(pl.lit('GORE AYACUCHO')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '445').then(pl.lit('GORE CAJAMARCA')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '446').then(pl.lit('GORE CUSCO')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '447').then(pl.lit('GORE HUANCAVELICA')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '448').then(pl.lit('GORE HUANUCO')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '449').then(pl.lit('GORE ICA')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '450').then(pl.lit('GORE JUNIN')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '451').then(pl.lit('GORE LA LIBERTAD')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '452').then(pl.lit('GORE LAMBAYEQUE')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '453').then(pl.lit('GORE LORETO')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '454').then(pl.lit('GORE MADRE DE DIOS')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '455').then(pl.lit('GORE MOQUEGUA')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '456').then(pl.lit('GORE PASCO')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '457').then(pl.lit('GORE PIURA')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '458').then(pl.lit('GORE PUNO')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '459').then(pl.lit('GORE SAN MARTIN')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '460').then(pl.lit('GORE TACNA')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '461').then(pl.lit('GORE TUMBES')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '462').then(pl.lit('GORE UCAYALI')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '463').then(pl.lit('GORE LIMA')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '464').then(pl.lit('GORE CALLAO')).otherwise(df1['pliego']).alias('pliego'))
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '465').then(pl.lit('GORE LIMA MML')).otherwise(df1['pliego']).alias('pliego'))

#########        USAR REPLACE EN POLAR CODIGO DE ARRIBA

###############################################################################
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 902)).then(pl.lit('001 SEDE')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 903)).then(pl.lit('002 PESQUERIA')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 904)).then(pl.lit('003 RIEGO & DRENAJE')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 906)).then(pl.lit('005 DESARROLLO RURAL')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 908)).then(pl.lit('100 AGRICULTURA')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 909)).then(pl.lit('200 TRANSPORTES')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 910)).then(pl.lit('300 GRE PUNO')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 911)).then(pl.lit('301 GRE SAN ROMAN')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 912)).then(pl.lit('302 GRE MELGAR')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 913)).then(pl.lit('303 GRE AZANGARO')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 914)).then(pl.lit('400 SALUD LAMPA')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 915)).then(pl.lit('401 SALUD MELGAR')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 916)).then(pl.lit('402 SALUD AZANGARO')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 917)).then(pl.lit('403 SALUD SAN ROMAN')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 918)).then(pl.lit('404 SALUD HUANCANE')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 919)).then(pl.lit('405 SALUD PUNO')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 920)).then(pl.lit('406 SALUD CHUCUITO')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 967)).then(pl.lit('407 SALUD YUNGUYO')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 968)).then(pl.lit('408 SALUD COLLAO')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1004)).then(pl.lit('304 GRE HUANCANE')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1005)).then(pl.lit('305 GRE PUTINA')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1006)).then(pl.lit('409 SALUD MACUSANI')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1007)).then(pl.lit('410 SALUD SANDIA')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1053)).then(pl.lit('306 GRE COLLAO')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1054)).then(pl.lit('307 GRE CHUCUITO')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1055)).then(pl.lit('308 GRE YUNGUYO')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1056)).then(pl.lit('309 GRE CARABAYA')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1339)).then(pl.lit('310 GRE SANDIA')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1434)).then(pl.lit('311 UGEL PUNO')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1435)).then(pl.lit('411 HOSPITAL MNB')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1504)).then(pl.lit('312 GRE LAMPA')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1505)).then(pl.lit('313 GRE MOHO')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1514)).then(pl.lit('314 GRE CRUCERO')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1621)).then(pl.lit('412 SALUD MELGAR')).otherwise(df1['UE']).alias('UE'))


##################################################################
print("--- %s seconds --- ENTIDADES PUNO" % (time.time() - start_time))

#####################      ENTIDADES UES GRT      #######################
df1 = df1.with_columns(pl.when(pl.col('pliego').str.slice(0, 3) == '460').then(pl.lit('GORE TACNA')).otherwise(df1['pliego Nombre abrev']).alias('pliego Nombre abrev'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 931)).then(pl.lit('001 SEDE')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 932)).then(pl.lit('100 AGRICULTURA')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 933)).then(pl.lit('200 TRANSPORTES')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 934)).then(pl.lit('300 EDUCACION')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 935)).then(pl.lit('400 SALUD')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 970)).then(pl.lit('401 HOSPITAL HIPOLITO UNANUE')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1210)).then(pl.lit('002 PROYECTO ESPECIAL')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1464)).then(pl.lit('301 UGEL TACNA')).otherwise(df1['UE']).alias('UE'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1622)).then(pl.lit('402 RED DE SALUD')).otherwise(df1['UE']).alias('UE'))
# print(df1)

##################################################################
print("--- %s seconds --- ENTIDADES TACNA" % (time.time() - start_time))

#########################         CLAVES     #########################
df1 = df1.with_columns((pl.col('ano_eje').cast(pl.Utf8, strict=False) + pl.col('pliego').str.slice(0,3) + pl.col("sec_ejec").cast(pl.Utf8, strict=False)+ pl.col("sec_func").cast(pl.Utf8, strict=False)+ pl.col("fuente_financ").str.slice(0, 1)+ pl.col("clasif_m")).alias('key'))
df1 = df1.with_columns((pl.col('ano_eje').cast(pl.Utf8, strict=False) + pl.col("sec_ejec").cast(pl.Utf8, strict=False)+ pl.col("dnpp").cast(pl.Utf8, strict=False)+ pl.col("fuente_financ").str.slice(0, 1)+ pl.col("rubro").str.slice(0, 2)).alias('key_proy'))
# df1 = df1.with_columns((pl.col('ano_eje').cast(pl.Utf8, strict=False) + pl.col("sec_ejec").cast(pl.Utf8, strict=False)+ pl.col("dnpp").cast(pl.Utf8, strict=False)+ pl.col("fuente_financ").str.slice(0, 1)).alias('key_proy'))
df1 = df1.with_columns((pl.col('ano_eje').cast(pl.Utf8, strict=False) + pl.col("sec_ejec").cast(pl.Utf8, strict=False)+ pl.col("sec_func").cast(pl.Utf8, strict=False)).alias('key_office'))
df1 = df1.with_columns((pl.col('ano_eje').cast(pl.Utf8, strict=False) + pl.col("sec_ejec").cast(pl.Utf8, strict=False)+ pl.col("dnpp").cast(pl.Utf8, strict=False)).alias('key_iri'))
df1 = df1.with_columns(pl.when((pl.col('key') == '2023107820612. 6.  8.  1.  4.  3')).then(pl.lit('Inversiones - IRI')).otherwise(df1['Concepto2']).alias('Concepto2')) # proyecto 2234985 de 1078 es IRI
df1 = df1.with_columns(pl.when((pl.col('Tipo de gasto') == 'Actividades') & (pl.col('Genérica').str.slice(0,2) == '24')).then(pl.lit('Transferencias diversas')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1072) & (pl.col('ano_eje') == 2023) & (pl.col('activ_obra_accinv').str.slice(0, 7) == '5000003')& (pl.col('sec_func') == 220)&(pl.col('Especifica det')=='23 2 71199')).then(pl.lit('Negociaciones colectivas')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when((pl.col('sec_ejec') == 1072) & (pl.col('ano_eje') == 2024) & (pl.col('activ_obra_accinv').str.slice(0, 7) == '5000003')& (pl.col('sec_func') == 292)&(pl.col('Especifica det')=='23 2 71199')).then(pl.lit('Negociaciones colectivas')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when((pl.col('Tipo de gasto')== 'Actividades')).then((pl.lit(''))).otherwise(df1['key_proy']).alias('key_proy'))
df1 = df1.with_columns(pl.when((pl.col('Tipo de gasto')== 'Inversiones')).then((pl.lit(''))).otherwise(df1['key']).alias('key'))
df1 = df1.with_columns((pl.col('key')+pl.col('key_proy')).alias('key_total'))
# print(df1)

##df1['key'] = df1['ano_eje'].astype(str) + df1['sec_ejec'].astype(str)+df1['sec_func'].astype(str)+df1['fuente_financ'].str[:1]+df1['clasif_m']
##df1['key_proy'] = df1['ano_eje'].astype(str) + df1['sec_ejec'].astype(str) + df1['dnpp'].astype(str) + df1['fuente_financ'].str[:1] + df1['rubro'].str[:2]
##df1['key_office'] = df1['ano_eje'].astype(str) + df1['sec_ejec'].astype(str) + df1['sec_func'].astype(str)
##df1['key_iri'] = df1['ano_eje'].astype(str) + df1['sec_ejec'].astype(str) + df1['dnpp'].astype(str)
# df1.loc[df1['key'] == '2023107820612. 6.  8.  1.  4.  3', 'Concepto2'] = 'Inversiones - IRI' # proyecto 2234985 de 1078 es IRI
# df1.loc[(df1['Tipo de gasto']== 'Actividades')&(df1['Genérica'].str[:2] == '24'),'Concepto2'] = 'Transferencias diversas'
# df1.loc[(df1['sec_ejec']== 1072)&(df1['ano_eje'] == 2023)&(df1['activ_obra_accinv'].str[:7]== '5000003')&(df1['sec_func'] == 220)&(df1['Especifica det'] == '23 2 71199'),'Concepto2'] = 'Negociaciones colectivas'
# df1.loc[(df1['sec_ejec']== 1072)&(df1['ano_eje'] == 2024)&(df1['activ_obra_accinv'].str[:7]== '5000003')&(df1['sec_func'] == 292)&(df1['Especifica det'] == '23 2 71199'),'Concepto2'] = 'Negociaciones colectivas'
# df1.loc[df1['Tipo de gasto'] == 'Actividades', 'key_proy'] = ''
# df1.loc[df1['Tipo de gasto'] == 'Inversiones', 'key'] = ''
# df1['key_total'] = df1['key'].astype(str) + df1['key_proy'].astype(str)
# print(df1.groupby(['nivel'])['mto_pim'].sum())


# prn_nivel = (
#     df1.group_by(['nivel',])
#     .agg([pl.sum('mto_pia'), pl.sum('mto_pim')])
#     # .agg(pl.col('MONTO_PIM').sum().cast(pl.Int64))#, pl.col('Devengado.').sum()
#     .sort(['nivel']))
# print(prn_nivel)

# print('Lectura inicial :', df1.shape)
print("--- %s seconds --- CLAVES" % (time.time() - start_time))
#########################         PROGRAMACION MTC     #########################
# dfp = pl.read_excel("D:/mjda/2023/Ejecucion/programacion/programacion_1669.xlsx")
dfp = dfp.filter(pl.col('TOTAL PROGG') != 0)
dfp = dfp.rename({
    'PROG_ENE':'prog_01',
    'PROG_FEB':'prog_02',
    'PROG_MAR':'prog_03',
    'PROG_ABR':'prog_04',
    'PROG_MAY':'prog_05',
    'PROG_JUN':'prog_06',
    'PROG_JUL':'prog_07',
    'PROG_AGO':'prog_08',
    'PROG_SET':'prog_09',
    'PROG_OCT':'prog_10',
    'PROG_NOV':'prog_11',
    'PROG_DIC':'prog_12',
    'TOTAL PROGG':'prog_total',
    })
df1 = df1.join(dfp, on='key_total', how='left')

df1 = df1.drop(['key_right','key_proy_right','FILT'])


# prn_nivel = (
#     df1.group_by(['nivel',])
#     .agg([pl.sum('mto_pia'), pl.sum('mto_pim')])
#     # .agg(pl.col('MONTO_PIM').sum().cast(pl.Int64))#, pl.col('Devengado.').sum()
#     .sort(['nivel']))
# print(prn_nivel)

# print('Lectura inicial :', df1.shape)
print("--- %s seconds --- PROGRAMACION MTC" % (time.time() - start_time))

#########################          PROGRAMACION MEF     #########################
# dfpi = pl.read_excel("D:/mjda/2023/Ejecucion/programacion/Consolidado_UEs_1912_36036.xlsx", dtype={'key_total':str})
# dfpi.drop(['dnpp','sec_ejec','año','PLIEGO','FORMATO','FICHA','N°','NIVEL_GOBIERNO','SECTOR','UNIDAD_EJECUTORA','SALDO_2023(A)-(B)','SALDO_ROOC_EXTERNO_2023','MODALIDAD_EJECUCION','ESTADO_SITUACION_PROYECTO','COMENTARIO','CORREO-CELULAR_RESPONSABLE','FECHA_CARGA','PIA_SIAF_2023','PIM_PROYECTADO_2023_(A)'],axis=1,inplace=True)
# print('dfpi :\n',dfpi)
dfpi = dfpi.rename({
    'PROYECCION_DEVENGADO_ENERO':'prog_mef_01',
    'PROYECCION_DEVENGADO_FEBRERO':'prog_mef_02',
    'PROYECCION_DEVENGADO_MARZO':'prog_mef_03',
    'PROYECCION_DEVENGADO_ABRIL':'prog_mef_04',
    'PROYECCION_DEVENGADO_MAYO':'prog_mef_05',
    'PROYECCION_DEVENGADO_JUNIO':'prog_mef_06',
    'PROYECCION_DEVENGADO_JULIO':'prog_mef_07',
    'PROYECCION_DEVENGADO_AGOSTO':'prog_mef_08',
    'PROYECCION_DEVENGADO_SETIEMBRE':'prog_mef_09',
    'PROYECCION_DEVENGADO_OCTUBRE':'prog_mef_10',
    'PROYECCION_DEVENGADO_NOVIEMBRE':'prog_mef_11',
    'PROYECCION_DEVENGADO_DICIEMBRE':'prog_mef_12',
    # 'DEVENGADO_TOTAL_2023_(B)':'prog_mef'
    })
df1 = df1.join(dfpi, on='key_total', how='left')
# df1 = df1.unique()
df1 = df1.fill_null(0)

# prn_nivel = (
#     df1.group_by(['nivel',])
#     .agg([pl.sum('mto_pia'), pl.sum('mto_pim')])
#     # .agg(pl.col('MONTO_PIM').sum().cast(pl.Int64))#, pl.col('Devengado.').sum()
#     .sort(['nivel']))
# print(prn_nivel)
# print('Lectura inicial :', df1.shape)

print("--- %s seconds --- PROGRAMACION MEF" % (time.time() - start_time))
###################                DE POLARS A PANDAS   Y    PANDAS A POLARS


# ######################################
# df1 = df1.to_pandas()
# df1[['prog_mef_01','prog_mef_02','prog_mef_03','prog_mef_04','prog_mef_05','prog_mef_06','prog_mef_07','prog_mef_08','prog_mef_09','prog_mef_10','prog_mef_11','prog_mef_12','prog_mef','prog_01','prog_02','prog_03','prog_04','prog_05','prog_06','prog_07','prog_08','prog_09','prog_10','prog_11','prog_12','prog_total']] = df1[['prog_mef_01','prog_mef_02','prog_mef_03','prog_mef_04','prog_mef_05','prog_mef_06','prog_mef_07','prog_mef_08','prog_mef_09','prog_mef_10','prog_mef_11','prog_mef_12','prog_mef','prog_01','prog_02','prog_03','prog_04','prog_05','prog_06','prog_07','prog_08','prog_09','prog_10','prog_11','prog_12','prog_total']].mask(df1[['prog_mef_01','prog_mef_02','prog_mef_03','prog_mef_04','prog_mef_05','prog_mef_06','prog_mef_07','prog_mef_08','prog_mef_09','prog_mef_10','prog_mef_11','prog_mef_12','prog_mef','prog_01','prog_02','prog_03','prog_04','prog_05','prog_06','prog_07','prog_08','prog_09','prog_10','prog_11','prog_12','prog_total']].duplicated(), 0)
# # print('pandas',df1)
# print("--- %s seconds --- DE POLARS A PANDAS" % (time.time() - start_time))
# df1 = pl.from_pandas(df1)
# # print('polars',df1)
# print("--- %s seconds --- PANDAS A POLARS" % (time.time() - start_time))
# ######################################



print("--- %s seconds --- ANTES - DUPLICADOS DE PROGRAMACION A CERO" % (time.time() - start_time))
# Supongamos que ya tienes un DataFrame 'df1' en Polars con las columnas mencionadas
columns_to_check = [
    'key_total','prog_mef_01','prog_mef_02','prog_mef_03','prog_mef_04','prog_mef_05','prog_mef_06','prog_mef_07','prog_mef_08','prog_mef_09','prog_mef_10','prog_mef_11','prog_mef_12',
    'prog_mef','prog_01','prog_02','prog_03','prog_04','prog_05','prog_06','prog_07','prog_08','prog_09','prog_10','prog_11','prog_12','prog_total'
]

# Crear una columna de comparación uniendo las columnas a verificar en una única columna de string
df1 = df1.with_columns(
    pl.concat_str([pl.col(col).cast(pl.Utf8) for col in columns_to_check], separator="_").alias("combined_cols")
)
# print(df1['combined_cols'])
# Identificar duplicados usando la columna combinada
df1 = df1.with_columns(
    pl.col("combined_cols").is_first_distinct().alias("is_duplicated")
)
# print(df1['combined_cols','is_duplicated'])
# Reemplazar los duplicados con 0 en las columnas seleccionadas
df1 = df1.with_columns([
    pl.when(~pl.col("is_duplicated"))
    .then(pl.lit(0))
    .otherwise(pl.col(col)).alias(col) for col in columns_to_check[1:]
])
# Eliminar la columna auxiliar 'combined_cols' y 'is_duplicated'
df1 = df1.drop(["combined_cols", "is_duplicated"])
# Mostrar el DataFrame resultante
# print(df1)

print("--- %s seconds --- DESPUES - DUPLICADOS DE PROGRAMACION A CERO" % (time.time() - start_time))





#########################          OFICINAS       #########################
dfo = dfo.with_columns(pl.col('key_office').cast(pl.Utf8, strict=False))

df1 = df1.join(dfo, on='key_office', how='left')
df1 = df1.with_columns(pl.col('ROF').fill_null(pl.lit('Por definir')))
df1 = df1.with_columns(pl.col('Dispositivo legal').fill_null(pl.lit('Por definir')))
df1 = df1.with_columns(pl.when((pl.col('key_total') == '2023107213212. 4.  1.  2.  1. 99')).then(pl.lit('Transferencias a OACI')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when((pl.col('Genérica') == '24 DONACIONES Y TRANSFERENCIAS')).then(pl.lit('Transferencias diversas')).otherwise(df1['Concepto2']).alias('Concepto2'))#        para verificar concepto2 =0

# df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '202')).then(pl.lit('DM')).otherwise(df1['VICE']).alias('VICE'))
# df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '203')).then(pl.lit('DM')).otherwise(df1['VICE']).alias('VICE'))
# df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '214')).then(pl.lit('DM')).otherwise(df1['VICE']).alias('VICE'))

# df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036')&(pl.col('sec_ejec') == 1078)).then(pl.lit('VMT')).otherwise(df1['VICE']).alias('VICE'))
# df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036')&(pl.col('sec_ejec') == 1250)).then(pl.lit('VMT')).otherwise(df1['VICE']).alias('VICE'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036')&(pl.col('sec_ejec') == 1669)).then(pl.lit('DM')).otherwise(df1['VICE']).alias('VICE'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036')&(pl.col('sec_ejec') == 1720)).then(pl.lit('VMC')).otherwise(df1['VICE']).alias('VICE'))
# df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036')&(pl.col('sec_ejec') == 1747)).then(pl.lit('VMT')).otherwise(df1['VICE']).alias('VICE'))

df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '012')&(pl.col('sec_ejec') == 993)).then(pl.lit('VMPE')).otherwise(df1['VICE']).alias('VICE'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '012')&(pl.col('sec_ejec') == 1066)).then(pl.lit('VMPE')).otherwise(df1['VICE']).alias('VICE'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '012')&(pl.col('sec_ejec') == 993)).then(pl.lit('J. PRODUCTIVOS')).otherwise(df1['ROF']).alias('ROF'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '012')&(pl.col('sec_ejec') == 1066)).then(pl.lit('TRABAJA PERU')).otherwise(df1['ROF']).alias('ROF'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '012')&(pl.col('sec_ejec') == 993)).then(pl.lit('J. PRODUCTIVOS')).otherwise(df1['DEP']).alias('DEP'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '012')&(pl.col('sec_ejec') == 1066)).then(pl.lit('TRABAJA PERU')).otherwise(df1['DEP']).alias('DEP'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '121')&(pl.col('sec_ejec') == 1510)).then(pl.lit('DM')).otherwise(df1['VICE']).alias('VICE'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '121')&(pl.col('sec_ejec') == 1510)).then(pl.lit('SUNAFIL')).otherwise(df1['ROF']).alias('ROF'))
df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '121')&(pl.col('sec_ejec') == 1510)).then(pl.lit('SUNAFIL')).otherwise(df1['DEP']).alias('DEP'))

valores_vice = ['202', '203', '214']
df1 = df1.with_columns(
    pl.when(pl.col('pliego').str.slice(0, 3).is_in(valores_vice))
    .then(pl.lit('DM'))
    .otherwise(df1['VICE'])
    .alias('VICE')
)

valores_vice_ejec = [1078, 1250, 1747]
df1 = df1.with_columns(
    pl.when(pl.col('sec_ejec').is_in(valores_vice_ejec))
    .then(pl.lit('VMT'))
    .otherwise(df1['VICE'])
    .alias('VICE')
)


# df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '202')).then(pl.lit('DIRECCION EJECUTIVA')).otherwise(df1['ROF']).alias('ROF'))
# df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '203')).then(pl.lit('DIRECCION EJECUTIVA')).otherwise(df1['ROF']).alias('ROF'))
# df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '214')).then(pl.lit('DIRECCION EJECUTIVA')).otherwise(df1['ROF']).alias('ROF'))

# df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036')&(pl.col('sec_ejec') == 1078)).then(pl.lit('DIRECCION EJECUTIVA')).otherwise(df1['ROF']).alias('ROF'))
# df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036')&(pl.col('sec_ejec') == 1250)).then(pl.lit('DIRECCION EJECUTIVA')).otherwise(df1['ROF']).alias('ROF'))
# df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036')&(pl.col('sec_ejec') == 1669)).then(pl.lit('DIRECCION EJECUTIVA')).otherwise(df1['ROF']).alias('ROF'))
# df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036')&(pl.col('sec_ejec') == 1720)).then(pl.lit('DIRECCION EJECUTIVA')).otherwise(df1['ROF']).alias('ROF'))
# df1 = df1.with_columns(pl.when((pl.col('pliego').str.slice(0, 3) == '036')&(pl.col('sec_ejec') == 1747)).then(pl.lit('DIRECCION EJECUTIVA')).otherwise(df1['ROF']).alias('ROF'))

valores_rof = ['202', '203', '214']
# Actualizar la columna "modalidad" en función de la coincidencia
df1 = df1.with_columns(
    pl.when(pl.col('pliego').str.slice(0, 3).is_in(valores_rof))
    .then(pl.lit('DIRECCION EJECUTIVA'))
    .otherwise(df1['ROF'])
    .alias('ROF')
)

valores_rof_ejec = [1078, 1250, 1669, 1720, 1747]
df1 = df1.with_columns(
    pl.when(pl.col('sec_ejec').is_in(valores_rof_ejec))
    .then(pl.lit('DIRECCION EJECUTIVA'))
    .otherwise(df1['ROF'])
    .alias('ROF')
)

df1 = df1.drop(['clasif_4','clasif_5','key_office','año_o','sec_ejec_o','sec_func_o','clasif_m'])

# df1 = df1.unique()
print("--- %s seconds --- OFICINAS" % (time.time() - start_time))

#########################          INVERSIONES IRI       #########################
df_i = df_i.with_columns(pl.col('key_iri').cast(pl.Utf8, strict=False))

df1 = df1.join(df_i, on='key_iri', how='left')
##null_count_df = df1.null_count()
##print('null_count_df',null_count_df)
df1 = df1.fill_null(0)
df1 = df1.with_columns(pl.when((pl.col('tip_inv') == 'Inversiones - IRI')&(pl.col('Genérica').str.slice(0,2) == '26')).then(pl.lit('Inversiones - IRI')).otherwise(df1['Concepto2']).alias('Concepto2'))
df1 = df1.with_columns(pl.when((pl.col('Genérica').str.slice(0,2) == '26')).then(pl.lit('Inversiones - TyC')).otherwise(df1['Concepto2']).alias('Concepto2'))#        para verificar concepto2 =0
df1 = df1.with_columns(pl.when((pl.col('dnpp').str.slice(0,7) == '2001621')).then(pl.lit('Pre inversión')).otherwise(df1['Concepto2']).alias('Concepto2'))#        para verificar concepto2 =0
df1 = df1.with_columns(pl.when((pl.col('producto_proyecto').str.slice(0,1) == '3')&(pl.col('Genérica').str.slice(0,2) == '26')).then(pl.lit('Inversiones - activo fijo')).otherwise(df1['Concepto2']).alias('Concepto2'))#        para verificar concepto2 =0
df1 = df1.drop(['key_iri','año_i','sec_ejec_i','dnpp_i','tip_inv'])

# df1 = df1.unique()

# ### COLUMNAS NEW ###
# df1 = df1.with_columns(Fecha_Gene = pl.lit(fecha))
# df1 = df1.with_columns(provincia_meta = pl.lit(''))
# df1 = df1.with_columns(distrito_meta = pl.lit(''))
# df1 = df1.with_columns(mto_girado_01 = pl.lit(0))
# df1 = df1.with_columns(mto_girado_02 = pl.lit(0))
# df1 = df1.with_columns(mto_girado_03 = pl.lit(0))
# df1 = df1.with_columns(mto_girado_04 = pl.lit(0))
# df1 = df1.with_columns(mto_girado_05 = pl.lit(0))
# df1 = df1.with_columns(mto_girado_06 = pl.lit(0))
# df1 = df1.with_columns(mto_girado_07 = pl.lit(0))
# df1 = df1.with_columns(mto_girado_08 = pl.lit(0))
# df1 = df1.with_columns(mto_girado_09 = pl.lit(0))
# df1 = df1.with_columns(mto_girado_10 = pl.lit(0))
# df1 = df1.with_columns(mto_girado_11 = pl.lit(0))
# df1 = df1.with_columns(mto_girado_12 = pl.lit(0))
# df1 = df1.with_columns(Girado_ = pl.lit(0))
# df1 = df1.with_columns(Saldo_Girado_ = pl.lit(0))

# Añadir todas las columnas con valores constantes en una sola llamada
df1 = df1.with_columns([
    pl.lit(fecha).alias('Fecha_Gene'),
    pl.lit('').alias('provincia_meta'),
    pl.lit('').alias('distrito_meta'),
    pl.lit(0).alias('mto_girado_01'),
    pl.lit(0).alias('mto_girado_02'),
    pl.lit(0).alias('mto_girado_03'),
    pl.lit(0).alias('mto_girado_04'),
    pl.lit(0).alias('mto_girado_05'),
    pl.lit(0).alias('mto_girado_06'),
    pl.lit(0).alias('mto_girado_07'),
    pl.lit(0).alias('mto_girado_08'),
    pl.lit(0).alias('mto_girado_09'),
    pl.lit(0).alias('mto_girado_10'),
    pl.lit(0).alias('mto_girado_11'),
    pl.lit(0).alias('mto_girado_12'),
    pl.lit(0).alias('Girado_'),
    pl.lit(0).alias('Saldo_Girado_')
])
df1 = df1.rename({'Fecha_Gene':'Fecha_Gene.','Girado_':'Girado.','Saldo_Girado_':'Saldo Girado.'})
# print(df1['Fecha_Gene'])
print("--- %s seconds --- PROGRAMACION IRI 1" % (time.time() - start_time))

################ INVERSIONES MODALIDAD #############
df_im = df_im.with_columns(pl.col('dnpp').cast(pl.Utf8, strict=False))
df_im = df_im.select(pl.col('dnpp', 'modalidad'))
df1 = df1.join(df_im, on='dnpp', how='left')
# df1 = df1.fill_null(0)

df1 = df1.with_columns(pl.when((pl.col('finalidad').str.slice(0,7) == '0285850')).then(pl.lit('6 COMPROMISOS CORRIENTES POR G2G')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0
df1 = df1.with_columns(pl.when((pl.col('finalidad').str.slice(0,7) == '0285850')&(pl.col('Genérica').str.slice(0,2) == '26')).then(pl.lit('8 EQUIPAMIENTO Y OTROS')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0
df1 = df1.with_columns(pl.when((pl.col('finalidad').str.slice(0,7) == '0285851')).then(pl.lit('6 COMPROMISOS CORRIENTES POR G2G')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0
df1 = df1.with_columns(pl.when((pl.col('finalidad').str.slice(0,7) == '0285851')&(pl.col('Genérica').str.slice(0,2) == '26')).then(pl.lit('8 EQUIPAMIENTO Y OTROS')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0
df1 = df1.with_columns(pl.when((pl.col('finalidad').str.slice(0,7) == '0285852')).then(pl.lit('6 COMPROMISOS CORRIENTES POR G2G')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0
df1 = df1.with_columns(pl.when((pl.col('finalidad').str.slice(0,7) == '0285852')&(pl.col('Genérica').str.slice(0,2) == '26')).then(pl.lit('8 EQUIPAMIENTO Y OTROS')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0
df1 = df1.with_columns(pl.when((pl.col('finalidad').str.slice(0,7) == '0285853')).then(pl.lit('4 COMPROMISOS CORRIENTES DE OXI')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0
df1 = df1.with_columns(pl.when((pl.col('finalidad').str.slice(0,7) == '0285853')&(pl.col('Genérica').str.slice(0,2) == '26')).then(pl.lit('8 EQUIPAMIENTO Y OTROS')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0
df1 = df1.with_columns(pl.when((pl.col('finalidad').str.slice(0,7) == '0285854')).then(pl.lit('6 COMPROMISOS CORRIENTES POR G2G')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0
df1 = df1.with_columns(pl.when((pl.col('finalidad').str.slice(0,7) == '0285854')&(pl.col('Genérica').str.slice(0,2) == '26')).then(pl.lit('8 EQUIPAMIENTO Y OTROS')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0
df1 = df1.with_columns(pl.when((pl.col('finalidad').str.slice(0,7) == '0285971')).then(pl.lit('6 COMPROMISOS CORRIENTES POR G2G')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0


# df1 = df1.with_columns(pl.when((pl.col('finalidad').str.slice(0,7) == '0053638')).then(pl.lit('1 COMPROMISOS CORRIENTES DE APP')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0
############################################################################################################################     OPTIMIZADO POR CHAT GPT
# Lista de valores a comparar
valores_finalidad = [
    '0053638', '0053655', '0053656', '0053657', '0053658', '0053659', '0053660', '0053661', '0053824', '0053825', '0053826', '0053827', '0053828', '0053829', '0053830', '0053831', '0053832', '0067484',
    '0067489', '0067490', '0075546', '0078330', '0078331', '0101901', '0107087', '0141811', '0167302', '0209913', '0237397', '0321426', '0343338', '0241908', '1630837', '1630838',
]# '0053654' 
# Actualizar la columna "modalidad" en función de la coincidencia
df1 = df1.with_columns(
    pl.when(pl.col('finalidad').str.slice(0, 7).is_in(valores_finalidad))
    .then(pl.lit('2 COMPROMISOS CORRIENTES DE APP'))
    .otherwise(df1['modalidad'])
    .alias('modalidad')
)

# df1 = df1.with_columns(ejecutora_nombre=pl.col("ejecutora_nombre").str.strip_chars())
# df1 = df1.with_columns((pl.col('modalidad')).alias('modalidad1'))
# valores_finalidads = [
#     '0001880', '0002100', '0052362', '0052383', '0052389', '0052394', '0052395', '0052396', '0052399', '0052402', '0052403', '0052406', '0052416', '0052419', '0052420', '0052421', '0052422', '0052423',
#     '0052424', '0052425', '0052974', '0052975', '0053818', '0067493', '0067505', '0095084', '0106294', '0141834', '0181910', '0190519', '0193665', '0224358', '0252051', '0285003', '0334677', '0416819',
#     '0417777', '0422922', '0433279',
# ]

# df1 = df1.with_columns(
#     pl.when(pl.col('finalidad').str.slice(0, 7).is_in(valores_finalidads))
#     .then(pl.lit('3 APP'))
#     .otherwise(df1['modalidad1'])
#     .alias('modalidad1')
# )
############################################################################################################################     OPTIMIZADO POR CHAT GPT


df1 = df1.with_columns(pl.when((pl.col('finalidad').str.slice(0,7) == '0053963')).then(pl.lit('3 COMPROMISOS CORRIENTES DE PROYECTOS EN ACTIVOS')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0
df1 = df1.with_columns(pl.when((pl.col('finalidad').str.slice(0,7) == '0290562')).then(pl.lit('3 COMPROMISOS CORRIENTES DE PROYECTOS EN ACTIVOS')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0
df1 = df1.with_columns(pl.when((pl.col('finalidad').str.slice(0,7) == '0290562')&(pl.col('Genérica').str.slice(0,2) == '26')).then(pl.lit('8 EQUIPAMIENTO Y OTROS')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0
df1 = df1.with_columns(pl.when((pl.col('Tipo de gasto') == 'Actividades')&(pl.col('Genérica').str.slice(0,2) == '26')).then(pl.lit('8 EQUIPAMIENTO Y OTROS')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0
df1 = df1.with_columns(pl.when((pl.col('Tipo de gasto') == 'Actividades')&(pl.col('Genérica').str.slice(0,2) == '24')).then(pl.lit('7 COMPROMISOS CORRIENTES POR TRANSFERENCIAS')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0
df1 = df1.with_columns(pl.when((pl.col('Tipo de gasto') == 'Actividades')&(pl.col('Genérica').str.slice(0,2) == '24')&(pl.col('categoria_gasto').str.slice(0,1) == '6')).then(pl.lit('7 TRANSFERENCIAS')).otherwise(df1['modalidad']).alias('modalidad'))#        para verificar concepto2 =0

df1 = df1.with_columns(pl.col('modalidad').fill_null(pl.lit('1 CONTINUA')))

# df1 = df1.merge(df_im, how='left', on='dnpp')
# df1 = df1.fillna(0)
# df1.drop_duplicates(inplace=True)#############
# print(df1.shape)
##################
print("--- %s seconds --- MODALIDAD" % (time.time() - start_time))

################ INVERSIONES ESTRATEGICAS #############
df_estrategica = df_estrategica.with_columns(pl.col('dnpp').cast(pl.Utf8, strict=False))

df_estrategica = df_estrategica.select(pl.col('dnpp', 'anexo1'))
df1 = df1.join(df_estrategica, on='dnpp', how='left')
df1 = df1.with_columns(pl.col('anexo1').fill_null(pl.lit('No estratégica')))

# df1 = df1.unique()

print("--- %s seconds --- ESTRATEGICA" % (time.time() - start_time))
##################

################ INVERSIONES ANEXOS 4 - PIRCC #############
df_a4 = df_a4.with_columns(pl.col('dnpp').cast(pl.Utf8, strict=False))

df_a4 = df_a4.select(pl.col('dnpp', 'anexo4'))
df1 = df1.join(df_a4, on='dnpp', how='left')
df1 = df1.with_columns(pl.col('anexo4').fill_null(pl.lit('No PIRCC')))
# df1 = df1.fill_null(0)
# df1 = df1.unique()
print("--- %s seconds --- INVERSIONES ANEXOS 4 - PIRCC" % (time.time() - start_time))
##################

################ INVERSIONES ANEXOS VI - INV GR GL #############
df_a6 = df_a6.with_columns(pl.col('dnpp').cast(pl.Utf8, strict=False))

df_a6 = df_a6.select(pl.col('dnpp', 'monto_VI','anexo6'))
df1 = df1.join(df_a6, on='dnpp', how='left')
df1 = df1.with_columns(pl.col('anexo6').fill_null(pl.lit('No Anexo VI')))
# df1 = df1.fill_null(0)
# df1 = df1.unique()
print("--- %s seconds --- INVERSIONES ANEXOS VI - TRANSF A GR-GL" % (time.time() - start_time))
##################

################ INVERSIONES EJECUCION MENSUAL PARA COMPARATIVO DE MEF #############
# df01e = df1.select(pl.col('key_total', 'sec_ejec', 'dnpp','mto_devenga_01'))
# df01e = df01e.rename({'mto_devenga_01':'Ejecución'})
# df01e = df01e.with_columns(mes = pl.lit(f'31/01/{anio[0]}'))
# data_12e = pl.concat([df01e, df02e, df03e, df04e, df05e, df06e, df07e, df08e, df09e, df10e, df11e, df122e], how='vertical_relaxed')
############################################################################################################################     OPTIMIZADO POR CHAT GPT
# Lista con los nombres de las columnas de devengado por mes y las fechas correspondientes
devengados = [
    ('mto_devenga_01', f'31/01/{anio[0]}'),
    ('mto_devenga_02', f'28/02/{anio[0]}'),
    ('mto_devenga_03', f'31/03/{anio[0]}'),
    ('mto_devenga_04', f'30/04/{anio[0]}'),
    ('mto_devenga_05', f'31/05/{anio[0]}'),
    ('mto_devenga_06', f'30/06/{anio[0]}'),
    ('mto_devenga_07', f'31/07/{anio[0]}'),
    ('mto_devenga_08', f'31/08/{anio[0]}'),
    ('mto_devenga_09', f'30/09/{anio[0]}'),
    ('mto_devenga_10', f'31/10/{anio[0]}'),
    ('mto_devenga_11', f'30/11/{anio[0]}'),
    ('mto_devenga_12', f'31/12/{anio[0]}')
]
# Crear una lista de dataframes
dfs = []
for col, fechas in devengados:
    df_temp = df1.select(pl.col('key_total', 'sec_ejec', 'dnpp', col))
    df_temp = df_temp.rename({col: 'Ejecución'})
    df_temp = df_temp.with_columns(mes=pl.lit(fechas))
    dfs.append(df_temp)

# Concatenar todos los dataframes en uno solo
data_12e = pl.concat(dfs, how='vertical_relaxed')
# Renombrar, seleccionar y ordenar en una sola cadena de operaciones
data_12e = (
    data_12e
    .rename({'key_total': 'key_dnpp'})
    .select(['mes', 'key_dnpp', 'sec_ejec', 'dnpp', 'Ejecución'])
    .sort(['mes', 'key_dnpp', 'sec_ejec', 'dnpp', 'Ejecución'])
)
############################################################################################################################     OPTIMIZADO POR CHAT GPT

# df1 = df1.unique()
print("--- %s seconds --- INVERSIONES EJECUCION MENSUAL PARA COMPARATIVO DE MEF" % (time.time() - start_time))


################ INVERSIONES PROGRAMACION MENSUAL PARA COMPARATIVO DE MEF #############
# df01 = df1.select(pl.col('key_total', 'sec_ejec', 'dnpp','prog_01'))
# df01 = df01.rename({'prog_01':'Mef'})
# df01 = df01.with_columns(mes = pl.lit(f'31/01/{anio[0]}'))
# data_12 = pl.concat([df01, df02, df03, df04, df05, df06, df07, df08, df09, df10, df11, df122])
############################################################################################################################     OPTIMIZADO POR CHAT GPT
# Lista con los nombres de las columnas de devengado por mes y las fechas correspondientes
programaciones = [
    ('prog_01', f'31/01/{anio[0]}'),
    ('prog_02', f'28/02/{anio[0]}'),
    ('prog_03', f'31/03/{anio[0]}'),
    ('prog_04', f'30/04/{anio[0]}'),
    ('prog_05', f'31/05/{anio[0]}'),
    ('prog_06', f'30/06/{anio[0]}'),
    ('prog_07', f'31/07/{anio[0]}'),
    ('prog_08', f'31/08/{anio[0]}'),
    ('prog_09', f'30/09/{anio[0]}'),
    ('prog_10', f'31/10/{anio[0]}'),
    ('prog_11', f'30/11/{anio[0]}'),
    ('prog_12', f'31/12/{anio[0]}')
]
# Crear una lista de dataframes
dfs = []
for col, fechas in programaciones:
    df_temp = df1.select(pl.col('key_total', 'sec_ejec', 'dnpp', col))
    df_temp = df_temp.rename({col: 'Mef'})
    df_temp = df_temp.with_columns(mes=pl.lit(fechas))
    dfs.append(df_temp)

# Concatenar todos los dataframes en uno solo
data_12 = pl.concat(dfs, how='vertical_relaxed')
# Renombrar, filtrar, seleccionar y ordenar en una sola cadena de operaciones
data_12 = (
    data_12
    .rename({'key_total': 'key_dnpp'})
    .filter(pl.col('Mef') != 0)
    .select(['mes', 'key_dnpp', 'sec_ejec', 'dnpp', 'Mef'])
    .sort(['mes', 'key_dnpp', 'sec_ejec', 'dnpp', 'Mef'])
)
# data_12 = data_12.unique(maintain_order=True)
############################################################################################################################     OPTIMIZADO POR CHAT GPT

# df1 = df1.unique(maintain_order=True)
df1 = df1.select([
    # "ano_eje", "nivel", "sector", "pliego", "sec_ejec", "unidad_ejecutora", "departamento_ejecutora", "provincia_ejecutora", "distrito_ejecutora", "sec_func", "programa_pptal", "tipo_act_obra_ac", "tipo_prod_proy", "producto_proyecto", "activ_obra_accinv", "funcion", "division_fn", "grupo_fn", "meta", "finalidad", "departamento_meta", "fuente_financ", "rubro", "tipo_recurso", "categoria_gasto", "tipo_transaccion", "Genérica", "Sub generica", "Subgenerica det", "Especifica.", "Especifica det", "Clasificador", "mto_pia", "mto_pim", "mto_certificado", "mto_compro_anual", "mto_devenga_01", "mto_devenga_02", "mto_devenga_03", "mto_devenga_04", "mto_devenga_05", "mto_devenga_06", "mto_devenga_07", "mto_devenga_08", "mto_devenga_09", "mto_devenga_10", "mto_devenga_11", "mto_devenga_12", "Devengado.", "Saldo Cert.", "Saldo Dev.", "Tipo de gasto", "clasificador_m", "Rango","Concepto1", "Concepto2", "lamina 1", "lamina 6", "pliego Nombre abrev", "UE", "images", "restring", "key_total", "dnpp", "gastos", "ROF", "DEP", "VICE", "UEI", "Dispositivo legal", "modalidad", "anexo1", "anexo4", "anexo6", "monto_VI", "provincia_meta", "distrito_meta", "Fecha_Gene.", "prog_01", "prog_02", "prog_03", "prog_04", "prog_05", "prog_06", "prog_07", "prog_08", "prog_09", "prog_10", "prog_11", "prog_12", "prog_total", "prog_mef_01", "prog_mef_02", "prog_mef_03", "prog_mef_04", "prog_mef_05", "prog_mef_06", "prog_mef_07", "prog_mef_08", "prog_mef_09", "prog_mef_10", "prog_mef_11", "prog_mef_12", "prog_mef", "mto_girado_01", "mto_girado_02", "mto_girado_03", "mto_girado_04", "mto_girado_05", "mto_girado_06", "mto_girado_07", "mto_girado_08", "mto_girado_09", "mto_girado_10", "mto_girado_11", "mto_girado_12", "Girado.", "Saldo Girado.", "Gasto critico", "GC numeral"]
    "ano_eje", "nivel", "sector", "pliego", "sec_ejec", "unidad_ejecutora", "departamento_ejecutora", "provincia_ejecutora", "distrito_ejecutora", "sec_func", "programa_pptal", "tipo_act_obra_ac", "tipo_prod_proy", "producto_proyecto", "activ_obra_accinv", "funcion", "division_fn", "grupo_fn", "meta", "finalidad", "departamento_meta", "fuente_financ", "rubro", "tipo_recurso", "categoria_gasto", "tipo_transaccion", "Genérica", "Sub generica", "Subgenerica det", "Especifica.", "Especifica det", "Clasificador", "mto_pia", "mto_pim", "mto_certificado", "mto_compro_anual", "mto_devenga_01", "mto_devenga_02", "mto_devenga_03", "mto_devenga_04", "mto_devenga_05", "mto_devenga_06", "mto_devenga_07", "mto_devenga_08", "mto_devenga_09", "mto_devenga_10", "mto_devenga_11", "mto_devenga_12", "Devengado.", "Saldo Cert.", "Saldo Dev.", "Tipo de gasto", "clasificador_m", "Rango","Concepto1", "Concepto2", "lamina 1", "lamina 6", "pliego Nombre abrev", "UE", "images", "restring", "key_total", "dnpp", "gastos", "ROF", "DEP", "VICE", "UEI", "Dispositivo legal", "modalidad", "anexo1", "anexo4", "anexo6", "monto_VI", "provincia_meta", "distrito_meta", "Fecha_Gene.", "prog_01", "prog_02", "prog_03", "prog_04", "prog_05", "prog_06", "prog_07", "prog_08", "prog_09", "prog_10", "prog_11", "prog_12", "prog_total", "prog_mef_01", "prog_mef_02", "prog_mef_03", "prog_mef_04", "prog_mef_05", "prog_mef_06", "prog_mef_07", "prog_mef_08", "prog_mef_09", "prog_mef_10", "prog_mef_11", "prog_mef_12", "prog_mef", "mto_girado_01", "mto_girado_02", "mto_girado_03", "mto_girado_04", "mto_girado_05", "mto_girado_06", "mto_girado_07", "mto_girado_08", "mto_girado_09", "mto_girado_10", "mto_girado_11", "mto_girado_12", "Girado.", "Saldo Girado.", "Gasto critico", "GC numeral", 'coordinador', 'coordinador_email']
)
df1 = df1.sort(['pliego','sec_ejec','sec_func'])
# print(df1.sort('nivel').group_by('nivel').agg(pl.sum('mto_pim'), pl.sum('Devengado.').sort(descending=True)))

prn_ejec = (
    df1.group_by(['nivel',])
    .agg(pl.col('mto_pia').sum().cast(pl.Int64), pl.col('mto_pim').sum().cast(pl.Int64), (pl.col('Devengado.').sum().cast(pl.Int64)))
    # .agg(pl.col('mto_pim').sum().cast(pl.Int64), (pl.col('Devengado.').sum().cast(pl.Float64).round(2)))
    .sort(['nivel']))
print(prn_ejec)

# prn_nivel = (
#     df1.group_by(['NIVEL_GOBIERNO',])
#     .agg([pl.sum('MONTO_PIA'), pl.sum('MONTO_PIM')])
#     # .agg(pl.col('MONTO_PIM').sum().cast(pl.Int64))#, pl.col('Devengado.').sum()
#     .sort(['NIVEL_GOBIERNO']))
# print(prn_nivel)
print('Lectura inicial :', df1.shape)
print("--- %s seconds --- INVERSIONES PROGRAMACION MENSUAL PARA COMPARATIVO DE MEF" % (time.time() - start_time))

# df = df.loc[df['key_total']!=0]#Sistema de Seguimiento de Inversiones SSI
df1 = df1.filter(pl.col('ano_eje') == anio[0])
df1 = df1.with_columns(pl.col('ano_eje').cast(pl.Utf8, strict=False))
df1 = df1.with_columns(pl.col('sec_ejec').cast(pl.Utf8, strict=False))


data_12 = data_12.with_columns(
    pl.col('sec_ejec').cast(pl.Utf8, strict=False),
    pl.col('Mef').cast(pl.Float64)
    )
formato_fecha = '%d/%m/%Y'


data_12e = data_12e.with_columns(
    # pl.col("mes").str.to_date(format=formato_fecha),
    pl.col('sec_ejec').cast(pl.Int64),
    pl.col('dnpp').cast(pl.Int64)
    )
# print('data_12e',data_12e)
print("--- %s seconds --- ANTES DE GRABAR" % (time.time() - start_time))

# df1 = df1.fill_null(0)
###############################
dfss = (df1
       .select(['nivel','dnpp','sector','pliego','sec_ejec','unidad_ejecutora','producto_proyecto', 'pliego Nombre abrev'])
       .with_columns(clave = pl.col('sec_ejec') + pl.col('dnpp'))
       .with_columns(ruc = pl.lit(''))
       )
dfss = (dfss
 .filter(pl.col('nivel').str.slice(0,1) == 'E')
#  .filter(pl.col('sector').str.slice(0,2)=='36')
 .filter(pl.col('producto_proyecto').str.slice(0,1)=='2')
 .group_by(['dnpp','sector','pliego','sec_ejec','unidad_ejecutora', 'producto_proyecto','pliego Nombre abrev','clave', 'ruc'])
 .agg(pl.col('dnpp').unique().count().alias('entidades'))
)
dfss = (dfss
       .with_columns(pl.when(pl.col('dnpp') == '2001621').then(pl.lit('PRE INVERSION')).otherwise(pl.lit('')).alias('Tipo de Inversión'))#        para verificar concepto2 =0
       .with_columns(pl.col("sector").str.split(": ").alias("sector_split"))
       .with_columns(pl.col("pliego").str.split(": ").alias("pliego_split"))
       .with_columns(pl.col("unidad_ejecutora").str.split(". ").alias("unidad_ejecutora_split"))
       .with_columns(pl.col("sector_split").list.get(0).alias("sector"),
                     (pl.col("sector_split").list.get(1).alias("sector Nombre")))
       .with_columns(pl.col("pliego_split").list.get(0).alias("pliego"),
                     (pl.col("pliego_split").list.get(1).alias("pliego Nombre")))
       .with_columns(pl.col("unidad_ejecutora_split").list.get(0).alias("ejec"),
                     (pl.col("unidad_ejecutora_split").list.get(1).alias("unidad ejecutora")))
        .drop(["sector_split", 'pliego_split','unidad_ejecutora_split'])
.select(['dnpp','sector','sector Nombre','pliego','pliego Nombre','sec_ejec','ejec','unidad ejecutora','producto_proyecto', 'Tipo de Inversión', 'clave','entidades','ruc'])
.sort(['pliego', 'sec_ejec', 'dnpp'],descending=[False, False, False])
)
###############################

###############################            keys_dnpp     únicos
keys_dnpp = (df1
.filter(pl.col('Tipo de gasto') == 'Inversiones')
.select(pl.col('key_total'))
)
###############################

# archivof = archivo_zip_local[:-4]
# print('archivof', archivof)

if sector == '1':
    # print("--- %s seconds --- antes csv" % (time.time() - start_time))
    # df1.write_csv(f'{sector}_ejecucion_gastos_{anio[0]}_diciembre_da.csv' )
    print("--- %s seconds --- antes de parquet" % (time.time() - start_time))
    # df1.write_parquet(f'E:/escritorio/conadis_e/1/{sector}_ejecucion_gastos_{anio[0]}_diciembre_da.parquet', compression='lz4')
    df1.write_parquet(f'{cwd}/archivos_parquet/{filename_base}.parquet',  compression='lz4')
    print("--- %s seconds --- sheet 1" % (time.time() - start_time))

else:


    df1 = df1.to_pandas()
    data_12e = data_12e.to_pandas()
    with pd.ExcelWriter(f'{sector}_ejecucion_gastos_{anio[0]}_diciembre_da.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:  # doctest: +SKIP
        df1.to_excel(writer, sheet_name='Sheet1',index=False)#, engine='xlsxwriter')
        print("--- %s seconds --- sheet 1" % (time.time() - start_time))
        data_12e.to_excel(writer, sheet_name='T_ejec_mensual',index=False)
        print("--- %s seconds --- sheet 2" % (time.time() - start_time))
        # data_12.to_excel(writer, sheet_name='T_mef_mensual',index=False)
        print("--- %s seconds --- sheet 3" % (time.time() - start_time))
        pd.ExcelWriter.close



print('subido :', df1.shape)
print("--- %s seconds ---" % (time.time() - start_time))


# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

###                SHEET1
Sheet1 = 'abfss://at_so_peru@onelake.dfs.fabric.microsoft.com/lh_2025_peru.Lakehouse/Tables/Sheet1'
df1.write_delta(
    Sheet1,
    mode="overwrite",
    # storage_options=storage_options,
    delta_write_options={"schema_mode": "overwrite"},
    )


# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

###                DATA DEL MEF
df_19 = 'abfss://at_so_peru@onelake.dfs.fabric.microsoft.com/lh_2025_peru.Lakehouse/Tables/df_19'
data_12.write_delta(
    df_19,
    mode="overwrite",
    # storage_options=storage_options,
    delta_write_options={"schema_mode": "overwrite"},
    )

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

###                DATA DE EJECUCION MENSUAL
df_ejecucion_mensual = 'abfss://at_so_peru@onelake.dfs.fabric.microsoft.com/lh_2025_peru.Lakehouse/Tables/df_ejecucion_mensual'
data_12e.write_delta(
    df_ejecucion_mensual,
    mode="overwrite",
    # storage_options=storage_options,
    delta_write_options={"schema_mode": "overwrite"},
    )

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

###                KEYS_DNPP
t_keys_dnpp = 'abfss://at_so_peru@onelake.dfs.fabric.microsoft.com/lh_2025_peru.Lakehouse/Tables/keys_dnpp'
keys_dnpp.write_delta(
    t_keys_dnpp,
    mode="overwrite",
    # storage_options=storage_options,
    delta_write_options={"schema_mode": "overwrite"},
    )

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

###                KEYS_DNPP

# Crear un rango de fechas
fecha = pl.date_range(date(2025, 1, 1), date(2025, 12, 31), "1d", eager=True).alias("Fecha")
# Convertir a Polars DataFrame
df_fecha = pl.DataFrame({"Date": fecha})

# Crear columnas adicionales
calendar_df = df_fecha.with_columns(
    pl.col("Date").dt.year().alias("Año"),
    pl.col("Date").dt.month().alias("mes_numero"),
    pl.col("Date").dt.quarter().alias("trimestre"),
    pl.col("Date").dt.week().alias("semana_numero"),
)
calendar_df = calendar_df.with_columns(mes = pl.lit('')) 

mes_dict = {
    1 : 'Enero', 2 : 'Febrero', 3 : 'Marzo', 4 : 'Abril',
    5 : 'Mayo', 6 : 'Junio', 7 : 'Julio', 8 : 'Agosto',
    9 : 'Septiembre', 10 : 'Octubre', 11 : 'Noviembre', 12 : 'Diciembre',
}
# Crear una lista de condiciones y resultados
conditions = [
    (pl.col('mes_numero') == key, pl.lit(value))
    for key, value in mes_dict.items()
]
# Aplicar las condiciones en un bucle
expr = pl.when(conditions[0][0]).then(conditions[0][1])
for condition, result in conditions[1:]:
    expr = expr.when(condition).then(result)
# Finalizar con otherwise para mantener el valor original si no se cumple ninguna condición
calendar_df = calendar_df.with_columns(
    expr.otherwise(calendar_df['mes']).alias('mes')
)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

###                Fecha
path_table = 'abfss://at_so_peru@onelake.dfs.fabric.microsoft.com/lh_2025_peru.Lakehouse/Tables/Fecha'
calendar_df.write_delta(
    path_table,
    mode="overwrite",
    delta_write_options={"schema_mode": "overwrite"},
)


# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

###                UNE PARQUET 2023AL2025

from deltalake.writer import write_deltalake

# Ruta a la carpeta con los archivos
folder_path ='/lakehouse/default/Files/archivos_parquet'

# Filtrar los primeros 3 archivos que contengan 'parque' en el nombre
archivos = sorted([
    os.path.join(folder_path, f)
    for f in os.listdir(folder_path)
    if "parque" in f.lower() and f.endswith(('.csv', '.parquet'))
])[:3]
print(archivos)
# Leer y concatenar los archivos con Polars
dfs = [pl.read_csv(f) if f.endswith(".csv") else pl.read_parquet(f) for f in archivos]
df_total = pl.concat(dfs)

# Guardar como tabla Delta (requiere que el DataFrame tenga tipos compatibles con Arrow)
output_path = 'abfss://at_so_peru@onelake.dfs.fabric.microsoft.com/lh_2025_peru.Lakehouse/Tables/data_total'

# output_path = "ruta/salida/tabla_delta"

write_deltalake(output_path, df_total.to_arrow(), mode="overwrite")


# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************




# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }
