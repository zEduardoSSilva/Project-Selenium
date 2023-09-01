#Dolar
#Euro
#Ouro

#pip install selenium
#pip install pandas
#pip install openpyxl

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import openpyxl

#baixar web driver e colocar ele na mesma pasta do seu codigo
#chrome - > chromedriver
#firefox - > geckodrive

# Passo 1 - Pegar Cotação do Dolar
navegador = webdriver.Chrome("chromedriver.exe")
#Entrar no Navegador
navegador.get("https://www.google.com.br/")
#Pesquisar Cotação Dolar
navegador.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotação do Dolar")
navegador.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_Dolar = navegador.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
print(cotacao_Dolar)
navegador.quit()

# Passo 2 - Pegar Cotação do Euro
navegador = webdriver.Chrome("chromedriver.exe")
#Entrar no Navegador
navegador.get("https://www.google.com.br/")
#Pesquisar Cotação Euro
navegador.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotação do Euro")
navegador.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_Euro = navegador.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
print(cotacao_Euro)
navegador.quit()

# Passo 3 - Pegar Cotação do Ouro
navegador = webdriver.Chrome("chromedriver.exe")
#Entrar no Navegador
navegador.get("https://www.melhorcambio.com/ouro-hoje")
#Pesquisar Cotação Ouro
cotacao_Ouro = navegador.find_element_by_xpath('//*[@id="comercial"]').get_attribute("value")
cotacao_Ouro = cotacao_Ouro.replace(",", ".")
print(cotacao_Ouro)
navegador.quit()

# Passo 4 - Importar Base de Dados

tabela = pd.read_excel("Produtos.xlsx")
#print(tabela)

# Passo 5 - Atualizar a Cotação, Preço de Compra e de Venda

#Atualizar Cotação

tabela.loc[tabela["Moeda"] == "Dolar", "Cotação"] = float(cotacao_Dolar)
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"] = float(cotacao_Euro)
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"] = float(cotacao_Ouro)

#Atualizar Preço de Compra = Preço Original * Cotação

tabela["Preço Base Reais"] = tabela["Preço Base Original"] * tabela["Cotação"]

#Atualizar Preço de Venda = Preço Compra * Margem

tabela["Preço Final"] = tabela["Preço Base Reais"] * tabela["Margem"]
print(tabela)

# Passo 6 - Exportar o Relatorio Atualizado
tabela.to_excel("Produtos.xlsx", index=False)
navegador.quit()


