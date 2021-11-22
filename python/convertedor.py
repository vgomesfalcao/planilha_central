from datetime import date
import pandas as pd
import numpy as np
import re
import os
pd.set_option('max_colwidth', None)

def extract_filename(path:str):
  return re.findall('(?<=\/)(\w+)(?=\.xlsx)', path)[0]

def convertedor_dexion(path:str):
  dexion_df = pd.read_excel(path)
  dexion_df = dexion_df.drop(range(8))
  dexion_df = dexion_df.drop(dexion_df.index[-1])
  dexion_df = dexion_df.dropna(how='all', axis=1)
  dexion_df.columns = ['Data', 'Lançamento','Contra-Partida','Parenteses','Atalho','Histórico','Valor de Débito','Valor de Crédito', 'Saldo','Deb/Cred']
  dexion_df['Atalho'] = dexion_df['Parenteses'] + dexion_df['Atalho']
  dexion_df = dexion_df.drop(columns=['Parenteses'])
  dexion_df = dexion_df.reset_index(drop=True)


  for row in dexion_df.itertuples():
    if(pd.isnull(row.Data)):
      subindex = 1
      while True:
        if((row.Index-subindex) in dexion_df.index):
          linha_anterior = dexion_df.loc[[row.Index-subindex],'Histórico'].values[0]
          break
        else:
          subindex += 1

      dexion_df.loc[[row.Index-subindex],'Histórico'] =  linha_anterior + row.Histórico
      dexion_df = dexion_df.drop(row.Index)
  dexion_df = dexion_df[dexion_df['Data'] != 'CENTRAL ORGANIZACAO CONTABIL S/C LTDA']
  dexion_df = dexion_df.dropna(how='all', axis=0)
  dexion_df = dexion_df.reset_index(drop=True)

  dexion_df['Cheques'] = dexion_df['Histórico'].str.extract(r'((?<=CH:)[\d,\s]{3,10})')
  dexion_df['Cheques'] = dexion_df['Cheques'].str.replace(' ','')

  if(not os.path.exists('convertido')):
    os.makedirs('./convertido')

  filename = extract_filename(path)
  name = 'convertido/cleaned_dexion_' + filename + '.xlsx'
  if(os.path.exists(name)):
    for n in range(2,100):
      name = 'convertido/cleaned_dexion_' + filename + '_' + str(n) + '.xlsx'
      if(not os.path.exists(name)):
        break
  datatoexcel = pd.ExcelWriter(name)
  dexion_df.to_excel(datatoexcel)
  datatoexcel.save()

