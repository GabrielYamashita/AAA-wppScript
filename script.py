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
excel_data = pd.read_excel('./Atletas Carimbados.xlsx', sheet_name='Planilha1')
# print(excel_data)

# Recebe Parâmetros pré armazenados, como o profile do wpp
options = webdriver.ChromeOptions() # Incializa o Objeto de Options()
options.add_argument(CHROME_PROFILE_PATH) # Insere o Chrome Profile
options.add_experimental_option('excludeSwitches', ['enable-logging']) # Desabilita Logs

# Seleciona o Driver do Chrome
# driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver = webdriver.Chrome(executable_path=ChromeDriverManager(version="114.0.5735.16").install(), options=options)

# Entra no Whatsapp Web
driver.get('https://web.whatsapp.com')

# Espera input do Usuário para Começar
input("\nAperte ENTER para mandar as Mensagens.\n--> ")
print("\nLOG das Mensagens:")

# Passa por cada Linha do Excel com cada Informação
for i, column in enumerate(excel_data['Telefone'].tolist()):
   try:
      # Recebe Dados das Colunas do Excel e Separa em Variáveis para o Texto
      nome = str(excel_data['Nome'][i])
      contato = str(excel_data['Telefone'][i]).replace('.0', '').replace("(", "").replace(")", "").replace(" ", "").replace("+", "").replace("-", "").replace("+55", "")
      # pago = str(excel_data['Pagou (sim/não)'][i]).lower().strip()

      if len(contato) < 9:
         erro = f"| (contato não encontrado)"

      # Faz o Corpo da Mensagem
      texto = f'''
Oie, td bem?%0A%0A

Nós da Atlética precisamos da sua ajuda! Até quarta feira, dia 16/08, respondam esse forms com sua matrícula atualizada de 23.2.%0A%0A

Se você recebeu isso e não vai jogar, desconsidere, por favor!%0A%0A

Se você já preencheu, obrigada!%0A%0A

Como retirar a matrícula?%0A%0A 

aluno on-line > secretaria virtual > solicitação de serviços > autoatendimento - declaração de matrícula > clicar em confirmar > descer a tela e clicar em confirmar solicitação > pdf%0A%0A

Qualquer dúvida, entre em contato com:%0A%0A

Beatriz Bozzo - 11 96442-1343%0A
Julia Benatti - 11 98797-4030%0A
Luana Vargas - 11 98570-9900%0A%0A%0A


https://docs.google.com/forms/d/e/1FAIpQLSfrcSaq80S_5_hi3kcM_mbrHhkvQyhvET11AKsHymXwhsu_QA/viewform
'''

      # Cria a URL para o Contato
      url = 'https://web.whatsapp.com/send?phone=55' + contato + '&text=' + texto

      # Entra no Link de Contato do Whatsapp
      driver.get(url)
      if len(contato) >= 9:
         try:
            # Espera 35s ou até se tornar Clicável
            click_btn = WebDriverWait(driver, 35).until(EC.element_to_be_clickable((By.CLASS_NAME, '_3XKXx')))

         except Exception as e:
            # Retorna Erro caso a mensagem não tenha sido enviada
            print(f"{i+1}) Desculpe, mensagem não enviada para {nome}")

         else:
            time.sleep(3) # Tempo para Carregar Página
            click_btn.click() # Aperta o Botão de Enviar
            time.sleep(2) # Tempo para Mandar Mensagem

            # Retorna Mensagem de Sucesso para a Pessoa
            print(f"{i+1}) Mensagem enviada para: {nome} | ({contato})")

      else:
         print(f"{i+1}) Desculpe, mensagem não enviada para {nome} {erro}")

   except Exception as e:
      # Retorna Mensagem de Erro 
      print(f"{i+1}) Mensagem falhou para: {nome} | {str(e)}")

print("Encerrando o script...")
time.sleep(5)
driver.quit() # Sai do Navegador
print("\nO script foi executado com sucesso.") # Mostra Mensagem de Sucesso da Execução do Script