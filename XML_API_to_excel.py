#O código abaixo é utilizado na minha empresa para consumir uma API corporativa e retornar a lista de serviços pendentes de uma determinada atividade
#Para fins de compliance, ocultei os dados sensíveis, assim como o nome da ferramenta
#Adaptando para sua finalidade, deve funcionar para qualquer API que retorne dados em XML


#Começamos importando as bibliotecas necessárias
#No meu caso, é necessária uma autenticação Digest, então foi necessário importar a HTTPDigestAuth
import requests
from requests.auth import HTTPDigestAuth
import xmltodict
from datetime import datetime
import pandas as pd
from time import sleep

#Em URL, você deverá fornecer o endereço da API
print('Iniciando extração de Serviços')
url = 'http://nome_da_ferramento/Ocorrencia/ajaxLoadExtend.html?iID_TIPO_OCORRENCIA=14,84,85,9191'

#Aqui eu criei um looping para caso a API esteja sendo altamente demandada, ele irá tentar diversas vezes até conseguir logar
resposta = 0
while resposta == 0:
    try:
        response = requests.get(url, auth=HTTPDigestAuth('Usuario_login', 'senha_123'),timeout=20)
        resposta = 1
    except:
        resposta = 0
        sleep (5)

#após realizar o login com sucesso, vamos começar a tratar os dados XML, transformando em um dicionário python:
r = response.text
d = xmltodict.parse(r)

#no meu caso, os serviços estão dentro de uma chave do dicionário chamada 'Servicos', na verdade não é esse o nome, mas para ocultar o nome original utilize esse:
portal = d['Servicos']
backlog = portal['Ocorr']


#vamos criar uma lista vazia e um dataframe vazio do pandas, nesse dataframe vamos já deixar o nome das colunas definidos, para o meu caso ficou assim:
lista = []
tabela = pd.DataFrame(lista, columns=['CID_CONTRATO','DT_OCORRENCIA','COD_OPERADORA','CONTRATO', 'ID_OCORRENCIA', 'NODE', 'COD_CELULA','ID_NOTIFICACAO','COD_IMOVEL','DT_CADASTRO','PRE_DT_AGENDA','AGENDA_DESCR','ID_USR','ID_USR_PF','OBS'],index=None)

#agora vamos percorrer a API XML, no meu caso eu só quero os casos em que o campo de "NOTIFICACAO" estiver zerado ou vazio, então chamei o IF para aplicar essa regra de negócio
#a cada "passo" do laço for, vamos adicionar uma lista de dados ordenados na nossa lista que inicialmente estava vazia
for ie in backlog:
    if len(ie['@ID_NOTIFICACAO']) >= 1:
        lista.append([str(ie['@CID_CONTRATO']),ie['@DT_OCORRENCIA'],int(ie['@COD_OPERADORA']), int(ie['@NUM_CONTRATO']), int(ie['@ID_OCORRENCIA']), str(ie['@COD_NODE']), str(ie['@COD_CELULA']),int(ie['@ID_NOTIFICACAO']),int(ie['@COD_IMOVEL']), ie['@OBS'], ie['@DT_OCORRENCIA'],ie['@PRE_DT_AGENDA'],ie['@PRE_AGENDA_DESCR'],ie['@USR_ATEND'],'',ie['@OBS']])

#Após concluirmos a etapa de percorrer a API, vamos popular nossa tabela inicialmente criada vazia com os dados que estão presentes na nossa lista
tabela = pd.DataFrame(lista, columns=['CID_CONTRATO', 'DT_OCORRENCIA','COD_OPERADORA', 'CONTRATO', 'ID_OCORRENCIA', 'NODE', 'COD_CELULA','ID_NOTIFICACAO','COD_IMOVEL', 'OBS', 'DT_CADASTRO','PRE_DT_AGENDA','AGENDA_DESCR','ID_USR','ID_USR_PF','OBS'],index=None)
#Por fim, vamos exportar nossa tabela virtualizada que existe no momento apenas no nosso código, em uma tabela Excel, é necessário definir o caminho e o nome da sheet:
tabela.to_excel('C:/Relatorios/DataBases/backlog_servicos.xlsx', engine='xlsxwriter', index=None, sheet_name='backlog_servicos')

#Fim da aplicação, você leu a API com dados XML não estruturados e estruturou eles em uma tabela Excel, tudo isso em poucos segundos
print('Extração de Serviços concluída')
sleep(3)
