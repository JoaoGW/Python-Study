import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor
from sklearn import metrics

#Importar a Base de Dados para o Código
tabela = pd.read_csv("advertising.csv")
print (tabela)

#Tratar os Dados encontrados na Base de Dados
print(tabela.info()) #A tabela já está tratada

#Análise Exploratória
sns.heatmap(tabela.corr(), cmap="Wistia", annot=True) #Criação do Gráfico
plt.show() #Exibição do Gráfico

#Separação de Treino para Teste
y = tabela["Vendas"]
x = tabela[["TV", "Radio", "Jornal"]]
x_treino, x_teste, y_treino, y_teste = train_test_split(x,y, test_size= 0.3)

#Modelos de Inteligência Artificial
modelo_regressaolinear = LinearRegression()
modelo_arvoredecisao = RandomForestRegressor()

modelo_regressaolinear.fit(x_treino, y_treino)
modelo_arvoredecisao.fit(x_treino, y_treino)

#Teste dos Modelos
previsao_regressaolinear = modelo_regressaolinear.predict(x_teste)
previsao_arvoredecisao = modelo_arvoredecisao.predict(x_teste)
print(metrics.r2_score(y_teste, previsao_regressaolinear))
print(metrics.r2_score(y_teste, previsao_arvoredecisao))

#Visualização Gráfica
tabela_auxiliar = pd.DataFrame()
tabela_auxiliar['y_teste'] = y_teste
tabela_auxiliar['Previsao Regressao Linear'] = previsao_regressaolinear
tabela_auxiliar['Previsao Arvore Decisao'] = previsao_arvoredecisao
sns.lineplot(data=tabela_auxiliar)
plt.show()

#Nova previsão
novos = pd.read_csv("novos.csv")
print(novos)
previsao = modelo_arvoredecisao(novos)
print(previsao)