
# %%
import shutil
import os
import datetime
import pandas as pd
import openpyxl




libro_EAF=openpyxl.load_workbook("C:\\Users\\luis.lizana\\OneDrive - Coordinador Eléctrico Nacional\\Archivos actualizables para automatización\\Indice Estudios Análisis de Fallas_2019-2020-2021-2022 v9.xlsx")

nom_hojas=libro_EAF.sheetnames
df=pd.read_excel("C:\\Users\\luis.lizana\\OneDrive - Coordinador Eléctrico Nacional\\Archivos actualizables para automatización\\Indice Estudios Análisis de Fallas_2019-2020-2021-2022 v9.xlsx", sheet_name=nom_hojas[4])
columns=df.loc[2]
columns=columns.reset_index(drop=True)
df=df.rename(columns=dict(zip(df.columns, columns)))
df=df.fillna("")
df=df.drop(df.index[:3])
df=df.reset_index(drop=True)
df=df.loc[~(df["TITULO"]=="")]

df=df.loc[df["Autor"].str.contains("LLQ")]
df=df.reset_index(drop=True)
df=df.loc[~(df["Fecha Emisión"]=="Anulado")]
df=df.reset_index(drop=True)
df2=pd.DataFrame()
df2["Asunto"]="EAF "+df["EAF N°"]+": "+df["TITULO"]
df2["Fecha de inicio"]=df["Fecha de la Falla"]
df2["Fecha de vencimiento"]=df["Fecha Max. Inf. SEC"]
df2["Prioridad"]=0
df2["% Completado"]=0
df2["Estado"]="No comenzada"
df2["Categoría"]="Categoría verde"
df2["Mensaje"]=df["EMPRESAS INVOLUC./ IF"]+"\nHora de la falla: "+df["Hora inicio"].astype(str)
df2.to_excel("salida_indice_eaf.xlsx", index=False)

# %%
