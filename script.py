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
excel_data = pd.read_excel('./Precisam preencher o forms da carimbação.xlsx', sheet_name='Planilha1')
# print(excel_data)

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
Boa tarde, pessoal!%0A%0A

Nós, da Atlética, precisamos muito da ajuda de vocês para organizarmos a inscrição de vocês no Economíadas. %0A%0A

O que precisam fazer?%0A
Até essa sexta-feira 12h, dia 23/06, respondam esse forms com seu documento com foto (escaneado ou digital) e sua rede social!%0A%0A%0A


Quem pode fazer?%0A
TODOS os atletas do time que estarão cursando regularmente o Insper no semestre que vem (2023.2)%0A%0A

OBS: Atletas que se formaram agora em 2023.1, NÃO poderão jogar.%0A%0A

Me inscrever nesse forms garante que eu estarei jogando no Econo?%0A
Não! Isso sempre é critério da Comissão Técnica, mas é necessário que todos os atletas preencham o formulário de forma a garantir sua inscrição. Quem não responder até sexta-feira, não terá a possibilidade de jogar o Economíadas de 2023.%0A%0A

Está em dúvida se vai ao jogos ou não?%0A
Responda mesmo assim para garantir que caso você vá, você esteja liberado para jogar!%0A%0A

https://forms.gle/FUVQTSMPkfpoidv3A
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