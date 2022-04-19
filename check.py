import pandas as pd

excel_data = pd.read_excel('Vendas Econo 2022.xlsx', sheet_name='Vendas KITS')
nomes = excel_data['NOME'].tolist()

for i in range(len(nomes)):
   nomes[i] = str(nomes[i]).strip()

with open("C:/Users/acer/Desktop/Vendas KITS.txt", "r", encoding="utf-8") as f:
   kit = []
   for line in f:
      kit.append( line.rstrip().replace("- ", "") )

naoEnviado = []
for k in kit:
   if "Desculpe" in k and "nan" not in k:
      naoEnviado.append(k.replace("Desculpe, mensagem não enviada para", "").strip())

# for k in naoEnviado:
#    print(f"({k})")

with open("naoEnviados.txt", "w", encoding="utf-8") as f:
   for nome in naoEnviado:
      tamanhoNome = len(nome)

      index = nomes.index(nome)
      tel = str(excel_data['TELEFONE'][index]).replace('.0', '').replace("(", "").replace(")", "").replace(" ", "").replace("+", "").replace("-", "").strip()
      
      classificacao = 'NÃO CLASSIFICADO'

      if len(tel) > 11:
         classificacao = 'FORMATO DO NÚMERO'
      
      if tel[3] == '.':
         classificacao = 'FOI USADO CPF'
         


      info = f"{nome}{' ' * (46 - tamanhoNome)} - ({tel}) - {classificacao}\n"
      f.write(info)


