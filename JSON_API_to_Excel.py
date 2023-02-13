#Precisamos inicialmente importar as bibliotecas requests e pandas
#Pandas para manipular dataframes
#Requests para requisitar a API online
#Para esse código funcionar, é preciso ter instalada também a biblioteca xlsxwritter

import pandas as pd
import requests
from datetime import datetime

# Vamos definir o link onde estão os dados não estruturados:
api = requests.get("http://ferramenta_empresa/json/itens.php")

#depois forçamos a formatação dos dados para JSON, essa etapa garante a funcionalidade
api_em_dict = api.json()

#Agora vamos criar uma lista vazia, uma tabela com as colunas pré-definidas
#E também variáveis de data e hora para visualizarmos o horário de execução
lista = []
tabela = pd.DataFrame(lista, columns=['TICKET', 'CIDADE', 'COD_OPER_NEWMONITOR', 'EMPRESA',  'URA', 'CRN', 'DT_INI', 'DT_PREV', 'SINTOMA', 'NATUREZA', 'GMUD', 'GRUPO', 'USUARIO', 'TITULO','OBS', 'ABRANGENCIAS', 'REC', 'REG','COD_IMOVEL_MDU','TITULO_TICKET','TICKET_DESCRICAO'],index=None)
data_e_hora_atuais = datetime.now()
data_e_hora_em_texto = data_e_hora_atuais.strftime('%d/%m/%Y %H:%M:%S')

#agora vamos percorrer os dados no dicionario gerado
#Alguns campos como "outage" sofrerão uma transformação forçada de texto em número
#Campos do tipo timestamp precisam de um cálculo para ficar no formato de data do excel
for dado in api_em_dict:
	lista.append([int(outage['ticket']),
								outage['cidade'],
								int(outage['cidade_id']),
								outage['empresa'],
								outage['ura'],
								outage['crn'],
								25569+(outage['data_ini']/86400000)-0.125,
								25569+(outage['data_prev']/86400000)-0.125,
								outage['sintoma'],
								outage['natureza'],
								outage['gmud'],
								outage['grupo'],
								outage['usuario'],
								outage['titulo'],
								outage['obs'],
								outage['abrangencias'],
								outage['rec'],
								outage['reg'],
								outage['cod_imovel_mdu'],
								outage['titulo_ticket'],
								outage['ticket_descricao']])


#Por fim criamos a tabela virtual e exportamos o resultado com duas linhas:
#A primeira recria a tabela, dessa vez inserindo os dados da nossa lista
#A segunda linha exporta a tabela virtual, é necessário aqui explicitar o caminho do arquivo
tabela = pd.DataFrame(lista, columns=['TICKET', 'CIDADE', 'COD_OPER_NEWMONITOR', 'EMPRESA',  'URA', 'CRN', 'DT_INI', 'DT_PREV', 'SINTOMA', 'NATUREZA', 'GMUD', 'GRUPO', 'USUARIO', 'TITULO','OBS', 'ABRANGENCIAS', 'REC', 'REG','COD_IMOVEL_MDU','TITULO_TICKET','TICKET_DESCRICAO'],index=None)
tabela.to_excel('C:/Users/alex.fraga/Desktop/Backlog_MDU/backlog/backlog_newmonitor.xlsx', engine='xlsxwriter', index=None)
