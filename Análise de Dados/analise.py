import pandas as pd
import plotly.express as px

#Abrir e visualizar a tabela na extensão csv
tabela = pd.read_csv("telecom_users.csv")
print(tabela)

#Correção dos problemas na base de dados
tabela = tabela.drop("Unnamed: 0", axis = 1) #exclusão da coluna Unnamed 0 da base de dados; axis = 0 (linha); axis = 1 (coluna)
tabela["TotalGasto"] = pd.to_numeric(tabela["TotalGasto"], errors="coerce") #correção de reconhecimento texto para um float
tabela = tabela.dropna(how= "all", axis= 1) #excluir colunas vazias
tabela = tabela.dropna(how= "any", axis= 0) #excluir linhas vazias

print(tabela.info())

#Análise Inicial
print(tabela["Churn"].value_counts())
print(tabela["Churn"].value_counts(normalize= True).map("{:.2%}".format))

#Análise Detalhada dos Clientes
for coluna in tabela.columns:
    grafico = px.histogram(tabela, x= coluna, color= "Churn")
    grafico.show()