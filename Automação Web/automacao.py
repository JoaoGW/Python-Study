from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd

#Abrir o Google Chrome
navegador = webdriver.Chrome()

#Links salvos
Link_Google = "https://www.google.com.br/"

#Pegar a cotação atual do Dólar
navegador.get(Link_Google)
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div[2]/div[2]/input').send_keys("Cotação Dólar")
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div[2]/div[2]/input').send_keys(Keys.ENTER)
Cotacao_dolar = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(Cotacao_dolar)

#Pegar a cotação do Euro
navegador.get(Link_Google)
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div[2]/div[2]/input').send_keys("Cotação Euro")
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div[2]/div[2]/input').send_keys(Keys.ENTER)
Cotacao_euro = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(Cotacao_euro)

#Pegar a cotação do Ouro
Link_Cambio_Ouro = "https://www.melhorcambio.com/ouro-hoje"
navegador.get(Link_Cambio_Ouro)
Cotacao_ouro = navegador.find_element(By.XPATH, '//*[@id="comercial"]').get_attribute('value')
Cotacao_ouro = Cotacao_ouro.replace(",", ".")
print(Cotacao_ouro)

#Fechar o Google Chrome
navegador.quit()

#Importar a Base de Dados e Atualizar os Valores das Cotações
tabela = pd.read_excel("Produtos.xlsx")
print(tabela)

#Calcular os Novos Preços e Salvar a Base de Dados
tabela.loc[tabela['Moeda'] == 'Dólar', 'Cotação'] = float(Cotacao_dolar)
tabela.loc[tabela['Moeda'] == 'Euro', 'Cotação'] = float(Cotacao_euro)
tabela.loc[tabela['Moeda'] == 'Ouro', 'Cotação'] = float(Cotacao_ouro)
print(tabela)

#Atualização Final
tabela["Preço de Compra"] = tabela["Preço Original"] * tabela["Cotação"]
tabela["Preço de Venda"] = tabela["Preço de Compra"] * tabela["Margem"]
print(tabela)

#Atualizar na Base de Dados no Excel
tabela.to_excel("Produtos_Atualizado.xlsx", index= False)