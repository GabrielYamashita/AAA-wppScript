import time
import pandas as pd

from config import CHROME_PROFILE_PATH # Importa Configurações do Chrome

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

# Recebendo Dados do Excel
excel_data = pd.read_excel('Hoteis.xlsx', sheet_name='Hotéis')

# Recebe Parâmetros pré armazenados, como o profile do wpp
options = webdriver.ChromeOptions() # Incializa o Objeto de Options()
options.add_argument(CHROME_PROFILE_PATH) # Insere o Chrome Profile
options.add_experimental_option('excludeSwitches', ['enable-logging']) # Desabilita Logs

# Seleciona o Driver do Chrome
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Entra no Whatsapp Web
driver.get('https://web.whatsapp.com')

# Espera input do Usuário para Começar
input("\nAperte ENTER para mandar as Mensagens.\n--> ")
print("\nLOG das Mensagens:")

# Passa por cada Linha do Excel com cada Informação
for i, column in enumerate(excel_data['Contato'].tolist()):
   try:
      # Recebe Dados das Colunas do Excel e Separa em Variáveis para o Texto
      contato = str(excel_data['Contato'][i])
      nome = excel_data['Nome'][i]
      saudacao = excel_data['Saudação'][i]
      adeus = excel_data['Adeus'][i]

      # Faz o Corpo da Mensagem a ser Enviada
      texto = f'''
{saudacao}, {nome}.%0A
{adeus}, {nome}.
'''

      # Cria a URL para o Contato
      url = 'https://web.whatsapp.com/send?phone=' + contato + '&text=' + texto

      # Entra no Link de Contato do Whatsapp
      driver.get(url)
      try:
         # Espera 35s ou até se tornar Clicável
         click_btn = WebDriverWait(driver, 35).until(EC.element_to_be_clickable((By.CLASS_NAME, '_4sWnG')))

      except Exception as e:
         # Retorna Erro caso a mensagem não tenha sido enviada
         print(f"- Desculpe, mensagem não enviada para {excel_data['Nome'][i]}")

      else:
         time.sleep(2) # Tempo para Carregar Página
         click_btn.click() # Aperta o Botão de Enviar
         time.sleep(1) # Tempo para Carregar Página

         # Retorna Mensagem de Sucesso para a Pessoa
         print(f"- Mensagem enviada para: {excel_data['Nome'][i]}")

   except Exception as e:
      # Retorna Mensagem de Erro 
      print(f"Mensagem falhou para: {excel_data['Nome'][i]} | {str(e)} ")

driver.quit() # Sai do Navegador
print("\nO script foi executado com sucesso.") # Mostra Mensagem de Sucesso da Execução do Script
