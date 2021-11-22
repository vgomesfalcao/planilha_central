import pandas as pd
import numpy as np
pd.set_option('max_colwidth', None)

path = 'itau_convertido.xlsx'
planilha = pd.read_excel(path)
planilha = planilha.drop(range(8))
planilha = planilha.drop(planilha.index[-1])
planilha = planilha.dropna(how='all', axis=1)
planilha.columns = ['Data', 'Lançamento','Contra-Partida','Parenteses','Atalho','Histórico','Valor de Débito','Valor de Crédito', 'Saldo','Deb/Cred']
planilha['Atalho'] = planilha['Parenteses'] + planilha['Atalho']
planilha = planilha.drop(columns=['Parenteses'])
planilha = planilha.reset_index(drop=True)


for row in planilha.itertuples():
  if(pd.isnull(row.Data)):
    subindex = 1
    while True:
      if((row.Index-subindex) in planilha.index):
        linha_anterior = planilha.loc[[row.Index-subindex],'Histórico'].values[0]
        break
      else:
        subindex += 1

    planilha.loc[[row.Index-subindex],'Histórico'] =  linha_anterior + row.Histórico
    planilha = planilha.drop(row.Index)
planilha = planilha[planilha['Data'] != 'CENTRAL ORGANIZACAO CONTABIL S/C LTDA']
planilha = planilha.dropna(how='all', axis=0)
planilha = planilha.reset_index(drop=True)

planilha['Cheques'] = planilha['Histórico'].str.extract(r'((?<=CH:)[\d,\s]{3,10})')
planilha['Cheques'] = planilha['Cheques'].str.replace(' ','')

datatoexcel = pd.ExcelWriter('processado/cleaned_'+path)
planilha.to_excel(datatoexcel)
datatoexcel.save()
