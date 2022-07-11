#Realizadao por Davir Rivera en Python 3.7
from time import sleep
from tkinter.tix import Tree
from openpyxl import load_workbook
import pandas as pd 

#Leectura de fuente de datos y guardado en dataframe
df=pd.read_csv('Datos a trabajar.txt',delimiter='\t')  #Se modifica datos a trabajar DatosAtrabajar.txt
print(df)
#Conversión de tipo de datos
df['Fecha']=pd.to_datetime(df['Fecha'],errors='coerce')
df['Clave_estación']=df['Clave_estación'].apply(str)
df['Químico']=df['Químico'].apply(str)

#Renombre de columnas
df=df.rename(columns={'Hora01':1,'Hora02':2,
'Hora03':3,'Hora04':4,'Hora05':5,'Hora06':6,'Hora07':7,'Hora08':8,'Hora09':9,
'Hora10':10,'Hora11':11,'Hora12':12,'Hora13':13,'Hora14':14,
'Hora15':15,'Hora16':16,'Hora17':17,'Hora18':18,'Hora19':19,'Hora20':20,
'Hora21':21,'Hora22':22,'Hora23':23,'Hora24':24})

#Sustitución de los valoes nulos por NaN

df.fillna('NaN', inplace=True)

#Anulación de la dinamización y reordenamiento del dataframe
df_unpivot=df.set_index(['Fecha','Clave_estación','Químico']).stack().unstack(2).reset_index()
# print(df_unpivot.head(10)) #->Solo para revisar

#Asignar nombre a la columna de horas
df_unpivot.rename(columns={'level_2':'Horas'},inplace=True)

# print(df_unpivot.head(25))

# Selección de las columnas a pertenecer al dataframe
df_unpivot=df_unpivot.reindex(columns=['Fecha','Clave_estación','Horas','CO','NO2','O3','SO2','RH','TMP','WDR','WSP'])
#Asingar nuevo indice para la eliminación del índice por defecto
df_unpivot=df_unpivot.set_index('Fecha')
print(df_unpivot.head(25))
print('Creando el archivo de excel "Datos_Ambientales_Sort " en el que se guardará la información')
print('Espere unos segundos...')
# Creación del archivo en excel donde se guardará la información
escrito = pd.ExcelWriter('Datos_Ambientales_Sort.xlsx')
# Escritura del DataFrame en el archivo creado
df_unpivot.to_excel(escrito)
# Guardado el archivo excel
escrito.save()

print('Agregando filtros a las columnas...')

#Carga del archivo creado
librox=load_workbook('Datos_Ambientales_Sort.xlsx')
hoja=librox.active
#Adición de filtros a las columnas del archivo
hoja.auto_filter.ref='A1:K69889'
#Guardado del archivo
librox.save('Datos_Ambientales_Sort.xlsx')
print('El DataFrame se ha escrito con éxito en el archivo de Excel.\n')