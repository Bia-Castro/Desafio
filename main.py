import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from time import sleep

# Ler a Planilha Excel
clientes = pd.read_excel('clientes.xlsx').fillna(value='').to_dict(orient='records')
print(clientes)

# Inicia o Driver
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)

# Clientes que não foram cadastrados
unregistered_customer = []

# Prencher o formulário
for cliente in clientes:
    nome = cliente['NOME']
    cpf_cnpj = str(cliente['CPF/CNPJ'])
    endereco = cliente['ENDEREÇO']
    tipo_instalacao = cliente['IMÓVEL']

    if not nome or not cpf_cnpj or not endereco or not tipo_instalacao:
        unregistered_customer.append(nome)
        continue

    # Acessar o Site
    driver.get('https://docs.google.com/forms/d/1ORS1KenIq98_jdMo1GK8RV2aK5U0JnizPDn3OuQ6TeI/viewform?edit_requested=true')

    # Espera
    sleep(3)

    # Preencher o campo Nome
    campo_nome = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input')
    campo_nome.send_keys(nome)

    # Preencher o campo CPF/CNPJ
    campo_cpf_cnpj = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input')
    campo_cpf_cnpj.send_keys(cpf_cnpj)

    # Preencher o campo Endereço
    campo_endereco = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input')
    campo_endereco.send_keys(endereco)

    # Selecionar o Tipo de Instalação
    if tipo_instalacao.upper() == 'C':
        driver.find_element(By.XPATH, '//*[@id="i17"]/div[3]/div').click()
    elif tipo_instalacao.upper() == 'A':
        driver.find_element(By.XPATH, '//*[@id="i20"]/div[3]/div').click()
    elif tipo_instalacao.upper() == 'R':
        driver.find_element(By.XPATH, '//*[@id="i23"]/div[3]/div').click()
    elif tipo_instalacao.upper() == 'I':
        driver.find_element(By.XPATH, '//*[@id="i26"]/div[3]/div').click()

    # Botaõ de enviar o formulário
    botao_enviar = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span').click()

print('teste', unregistered_customer)

driver.quit()

from discord_webhook import DiscordWebhook, DiscordEmbed

# if rate_limit_retry is True then in the event that you are being rate 
# limited by Discord your webhook will automatically be sent once the 
# rate limit has been lifted
webhook = DiscordWebhook(
    url='https://discord.com/api/webhooks/1102292423696207965/DErsEE-EreK9azvnF1Hzj5MzcfkI60vFMmEktWZ_muMM1R9SsxkjWUyXH9ItY1szPbrD',
    rate_limit_retry=True,
    content='Webhook Message')

embed = DiscordEmbed(title='Clientes que não foram cadastrados', description='\n'.join(unregistered_customer), color='03b2f8')

# add embed object to webhook
webhook.add_embed(embed)
response = webhook.execute()

input('<<<')