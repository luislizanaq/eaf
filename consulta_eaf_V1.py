
# %%
import shutil
import os
import datetime
import pandas as pd
import openpyxl
import streamlit as st

st.title("📄 Índice de EAF V1.0")
uploaded_file = st.file_uploader("Sube el índice de EAF V9 (.xlsx)", type=("xlsx"))

libro_EAF=openpyxl.load_workbook(uploaded_file)

nom_hojas=libro_EAF.sheetnames
df=pd.read_excel(uploaded_file, sheet_name=nom_hojas[4])
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

st.table(df2)


# %%
