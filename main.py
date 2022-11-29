from selenium import webdriver
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import time

servico = Service(EdgeChromiumDriverManager().install())
navegador = webdriver.Edge(service=servico)

navegador.get("https://www.google.com.br")

navegador.find_element("xpath",'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação dólar")
navegador.find_element("xpath",'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_dolar = navegador.find_element("xpath",'/html/body/div[7]/div/div[10]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div/div/div/div/div[3]/div[1]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_dolar)

navegador.get("https://www.google.com.br")
navegador.find_element("xpath",'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação euro")
navegador.find_element("xpath",'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_euro = navegador.find_element("xpath",'/html/body/div[7]/div/div[10]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div/div/div/div/div[3]/div[1]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_euro)

navegador.get("https://www.melhorcambio.com/ouro-hoje")
cotacao_ouro = navegador.find_element("xpath",'/html/body/div[5]/div[1]/div/div/input[2]').get_attribute("value")
cotacao_ouro = cotacao_ouro.replace(',','.')
print(cotacao_ouro)

tabela = pd.read_excel("Produtos.xlsx")

tabela.loc[tabela["Moeda"] == "Dólar", "Cotação"] = float(cotacao_dolar)
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"] = float(cotacao_euro)
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"] = float(cotacao_ouro)

tabela["Preço de Compra"] = tabela["Preço Original"] * tabela["Cotação"]

tabela["Preço de Venda"] = tabela["Preço de Compra"] * tabela["Margem"]

tabela["Preço Original"] = tabela["Preço Original"].map("R${:.2f}".format)
tabela["Cotação"] = tabela["Cotação"].map("R${:.2f}".format)
tabela["Preço de Compra"] = tabela["Preço de Compra"].map("R${:.2f}".format)
tabela["Preço de Venda"] = tabela["Preço de Venda"].map("R${:.2f}".format)
print(tabela)

tabela.to_excel("ProdutosAtualizados.xlsx", index=False)
