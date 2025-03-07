
# %%
import pandas as pd
import openpyxl
import streamlit as st

st.title("📄 Índice de EAF V1.0")
file = st.file_uploader("Sube el índice de EAF V9 (.xlsx)", type=("xlsx"))


if file is not None and file.name.startswith("Indice Estudios Análisis de Fallas"):


  libro_EAF=openpyxl.load_workbook(file)

  nom_hojas=libro_EAF.sheetnames

  option = st.selectbox(
    "Selecciona el Mes",
    (nom_hojas[4], nom_hojas[5], nom_hojas[6],nom_hojas[7], nom_hojas[8], nom_hojas[9]), )

  df=pd.read_excel(file, sheet_name=option)
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
  st.dataframe(df2)
  df3=pd.DataFrame()
  df3["N° EAF"]=df["EAF N°"]
  df3["TITULO"]=df["TITULO"]
  df3["Decha de Falla"]=df["Fecha de la Falla"]
  df3["Hora Falla"]=df["Hora inicio"]
  df3["Plazo Artículo 6-44"]=df["Fecha Max. Art. 6-44"]
  st.dataframe(df3)



# %%
