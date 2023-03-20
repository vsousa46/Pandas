from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time

navegador = webdriver.Chrome()

tabela = pd.read_excel("Emitir.xlsx")
print(tabela)

for i, cpf in enumerate(tabela["CPF"]):
    email = tabela.loc[i, "Email"]
    descrição = tabela.loc[i, "Descrição"]
    valor = tabela.loc[i, "Valor"]
    navegador.get("https://forms.office.com/Pages/ResponsePage.aspx?id=DQSIkWdsW0yxEjajBLZtrQAAAAAAAAAAAAYAAOS0dsdUQlJTR09TSTRIWTZWRlZYRVU0S1hMWFFMMC4u")
    time.sleep(0.5)
    navegador.find_element(By.XPATH, '//*[@id="form-main-content"]/div/div/div[5]/button/div').click()
    time.sleep(1)
    #Cpf
    navegador.find_element(By.XPATH, '//*[@id="cover-page-root"]/div[1]/div[2]/div[2]/div[1]/div/div[3]/div/span/input').send_keys(cpf)

    #email
    navegador.find_element(By.XPATH, '//*[@id="cover-page-root"]/div[1]/div[2]/div[2]/div[2]/div/div[3]/div/span/input').send_keys(email)

    #preencher descrição
    navegador.find_element(By.XPATH, '//*[@id="cover-page-root"]/div[1]/div[2]/div[2]/div[3]/div/div[3]/div/span/input').send_keys(descrição)

    #preencher Valor
    navegador.find_element(By.XPATH, '//*[@id="cover-page-root"]/div[1]/div[2]/div[2]/div[4]/div/div[3]/div/span/input').send_keys(str (valor))

    #clicar no botão
    navegador.find_element(By.XPATH, '//*[@id="cover-page-root"]/div[1]/div[2]/div[3]/div/button').click()
    time.sleep(1)
    navegador.find_element(By.XPATH, '//*[@id="form-main-content"]/div[1]/div[2]/div[2]/div[2]/a').click()

    time.sleep(30)
