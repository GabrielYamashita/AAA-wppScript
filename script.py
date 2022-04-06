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
excel_data = pd.read_excel('Pagamento de quarto pro codigo.xlsx', sheet_name='Quartos')

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
for i, column in enumerate(excel_data['Nome Principal'].tolist()):
   try:
      # Recebe Dados das Colunas do Excel e Separa em Variáveis para o Texto
      contato = str(excel_data['Telefone Principal'][i]).replace('.0', '')
      nome = excel_data['Nome Principal'][i]
      valor = str(excel_data['Valor'][i])

      if len(contato) < 9:
         print(f"- Contato não encontrado de: {nome}")
         continue

      # Faz o Corpo da Mensagem
      texto = f'''
Olá {nome}, aqui é a Luisa Bragaia da atlética do Insper! Estou passando pra te lembrar que você reservou um quarto de hotel para 2 pessoas com a gente para o Economíadas 2022 em São Carlos! É necessário efetuar o pagamento do valor total do quarto, de R$ {valor},00, presencialmente no Insper até quarta feira (06/04) e me mandar nome, telefone e RG de cada hóspede do quarto.%0A%0A
Caso já tenha efetuado o pagamento, peço desculpas! Vou precisar somente de um comprovante de pagamento e que você me confirme o nome, telefone e RG de cada hóspede do quarto.  %0A%0A
Qualquer dúvida ou problema, pode me chamar!!
'''

      # Cria a URL para o Contato
      url = 'https://web.whatsapp.com/send?phone=55' + contato + '&text=' + texto

      # Entra no Link de Contato do Whatsapp
      driver.get(url)
      try:
         # Espera 35s ou até se tornar Clicável
         click_btn = WebDriverWait(driver, 35).until(EC.element_to_be_clickable((By.CLASS_NAME, '_4sWnG')))

      except Exception as e:
         # Retorna Erro caso a mensagem não tenha sido enviada
         print(f"- Desculpe, mensagem não enviada para {nome}")

      else:
         time.sleep(3) # Tempo para Carregar Página
         click_btn.click() # Aperta o Botão de Enviar
         time.sleep(2) # Tempo para Mandar Mensagem

         # Retorna Mensagem de Sucesso para a Pessoa
         print(f"- Mensagem enviada para: {nome}")

   except Exception as e:
      # Retorna Mensagem de Erro 
      print(f"Mensagem falhou para: {nome} | {str(e)} ")

print("Encerrando o script...")
time.sleep(5)
driver.quit() # Sai do Navegador
print("\nO script foi executado com sucesso.") # Mostra Mensagem de Sucesso da Execução do Script
