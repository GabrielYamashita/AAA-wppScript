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
excel_data = pd.read_excel('./Vendas Onibus.xlsx', sheet_name='Onibus')
# print(excel_data)

# Recebe Parâmetros pré armazenados, como o profile do wpp
options = webdriver.ChromeOptions() # Incializa o Objeto de Options()
options.add_argument(CHROME_PROFILE_PATH) # Insere o Chrome Profile
options.add_experimental_option('excludeSwitches', ['enable-logging']) # Desabilita Logs

# Seleciona o Driver do Chrome
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
# driver = webdriver.Chrome(executable_path=ChromeDriverManager(version="114.0.5735.16").install(), options=options)

# Entra no Whatsapp Web
driver.get('https://web.whatsapp.com')

# Espera input do Usuário para Começar
input("\nAperte ENTER para mandar as Mensagens.\n--> ")
print("\nLOG das Mensagens:")

# Passa por cada Linha do Excel com cada Informação
for i, column in enumerate(excel_data['Telefone2'].tolist()):
   try:
      # Recebe Dados das Colunas do Excel e Separa em Variáveis para o Texto
      nome = str(excel_data['Nome2'][i])
      contato = str(excel_data['Telefone2'][i]).replace('.0', '').replace("(", "").replace(")", "").replace(" ", "").replace("+", "").replace("-", "").replace("+55", "")
      saida = str(excel_data['Saída'][i])

      saida = 'h' in saida or '00' in saida

      if len(contato) < 9:
         erro = f"| (contato não encontrado)"

            # Faz o Corpo da Mensagem
      texto = f'''
Oi {nome} tudo bem?%0A
Iniciamos a semana do economiadas!%0A
Passando aqui para perguntar qual horário você gostaria de sair do insper?%0A
16h,18h ou 19h?%0A%0A

E sobre o horário de volta (São Carlos para São Paulo)? 14h, 16h, 17h ou 18h?%0A
caso já tenha definido o seu horário, desconsidere essa mensagem
'''

            # Cria a URL para o Contato
      url = 'https://web.whatsapp.com/send?phone=55' + contato + '&text=' + texto

      # Entra no Link de Contato do Whatsapp
      driver.get(url)
      if not saida:
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

      else:
            print("Já possui Horário Alocado")
            continue


   except Exception as e:
      # Retorna Mensagem de Erro 
      print(f"{i+1}) Mensagem falhou para: {nome} | {str(e)}")

print("Encerrando o script...")
time.sleep(5)
driver.quit() # Sai do Navegador
print("\nO script foi executado com sucesso.") # Mostra Mensagem de Sucesso da Execução do Script