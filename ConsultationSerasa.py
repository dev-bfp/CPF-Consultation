import time
import pyautogui
import pprint
import requests
import gspread # LIB DO GOOGLE SHEETS
from oauth2client.service_account import ServiceAccountCredentials # LIB DE AUTENTICAÇÃO DA GOOGLE
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from datetime import datetime
from Tokens import *

# Configuração credencial GoogleApis
scope = ["https://www.googleapis.com/auth/drive"]
credentials_google = ServiceAccountCredentials.from_json_keyfile_name(caminho_local_credentials, scope)
client = gspread.authorize(credentials_google)
# Endereçamento GoogleSheets
nmSheets = "Vendas Automaticas"
sheet = client.open(nmSheets).worksheet("Consulta CPF")


class Selenium():    

    def click(element,value_id):

        element = element.lower().replace(' ','')
        if element == 'id':
            driver.find_element(by=By.ID, value=value_id).click() 
        elif element == 'name':
            driver.find_element(by=By.NAME, value=value_id).click()
        elif element == 'xpath':
            driver.find_element(by=By.XPATH, value=value_id).click()
        elif element == 'class':
            driver.find_element(by=By.CLASS_NAME, value=value_id).click()
        else:
            print('error')


    def sendtext(element,value_id,text):

        element = element.lower().replace(' ','')
        if element == 'id':
            driver.find_element(by=By.ID, value=value_id).send_keys(text)
        elif element == 'name':
            driver.find_element(by=By.NAME, value=value_id).send_keys(text)
        elif element == 'xpath':
            driver.find_element(by=By.XPATH, value=value_id).send_keys(text)
        elif element == 'class':
            driver.find_element(by=By.CLASS_NAME, value=value_id).send_keys(text)
        else:
            print('error')


    def clear(element,value_id):

        element = element.lower().replace(' ','')
        if element == 'id':
            driver.find_element(by=By.ID, value=value_id).clear()
        elif element == 'name':
            driver.find_element(by=By.NAME, value=value_id).clear()
        elif element == 'xpath':
            driver.find_element(by=By.XPATH, value=value_id).clear()
        elif element == 'class':
            driver.find_element(by=By.CLASS_NAME, value=value_id).clear()
        else:
            print('error')


    def find(element,value_id):

        element = element.lower().replace(' ','')
        if element == 'id':
            return driver.find_element(by=By.ID, value=value_id).get_attribute("innerHTML")
        elif element == 'name':
            return driver.find_element(by=By.NAME, value=value_id).get_attribute("innerHTML")
        elif element == 'xpath':
            return driver.find_element(by=By.XPATH, value=value_id).get_attribute("innerHTML")
        elif element == 'class':
            return driver.find_element(by=By.CLASS_NAME, value=value_id).get_attribute("innerHTML")
        else:
            print('error')
    # ------------------- Fim -------------------

def start():
    # ----------- Options chrome -----------
    options = Options()
    prefs = {"credentials_enable_service": False, "profile.password_manager_enabled": False}
    options.add_experimental_option("prefs", prefs)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    # ----------- Configura Webdriver -----------
    driver = webdriver.Chrome(options=options, service=Service(diretorio_chromedriver))
    return driver
    # ------------------- Fim -------------------


def send_msg2(bot_message):
    send_text = 'https://api.telegram.org/bot' + bot_token + '/sendMessage?chat_id=' + bot_chatID + '&parse_mode=Markdown&text=' + bot_message
    response = requests.get(send_text)
    res = response.json()
    print(res)
    if res['ok']:
        return res['result']['message_id']
    else:
        return 0

def delete_msg(id):
    if id != 0:
        send_text = 'https://api.telegram.org/bot' + bot_token + '/deleteMessage?chat_id=' + bot_chatID + '&message_id=' + str(
            id)
        response = requests.get(send_text)
        res = response.json()
        print(res)
        return res


def login_ixc(driver):
    # Realiza login no sistema IXC
    link = url_ixc
    driver.get(link)
    driver.maximize_window()

    Selenium.sendtext('id', 'email', email_ixc)  # insere email
    Selenium.sendtext('id', 'senha', senha_ixc)  # insere senha
    Selenium.click('id', 'entrar')  # clica em entrar
    time.sleep(1)
    try:
        try:
            Selenium.click('id','entrar')  # clica em entrar
            time.sleep(1)
        except:
            time.sleep(0.5)
            
        try:
            driver.implicitly_wait(10)  # seconds
            Selenium.click('xpath','//*[@id="slide_0"]/div[4]/vg-button')  # fecha pop-up
            time.sleep(0.5)
            Selenium.click('xpath','//*[@id="slide_1"]/div[4]/vg-button[2]')
        except:
            time.sleep(0.5)
            
        try:
            driver.implicitly_wait(5)  # seconds
            Selenium.click('xpath','//*[@id="warning"]/vg-body/div/vg-button[2]')  # fecha pop-up
        except:
            time.sleep(0.5)

        usuario = Selenium.find('xpath','//*[@id="layout_painel"]/div[2]/span').strip().split(' ')
        print(f'Welcome, {usuario[0]} {usuario[-1]}')
        
    except:
        print('Erro login')
        
    # ------------------- Fim -------------------

def consulta_serasa(CPF):
    driver.implicitly_wait(10)  # seconds
    Selenium.click('id','id_input_menu')  # abre pesquisa
    Selenium.sendtext('id','id_input_menu', 'cliente')  # insere pesquisa
    Selenium.click('xpath','//*[@id="grupo_menu04400d48d04acd3599cf545dafbb90ed"]/ul/li[1]/a') # clica na opção
    Selenium.click('xpath','//*[@id="1_grid"]/div/div[2]/div[1]/button[1]') # clica em novo
    Selenium.clear('id','cnpj_cpf') # limpa o campo
    time.sleep(0.1)
    for x in CPF:
        Selenium.sendtext('id','cnpj_cpf',x) # insere o CPF
    Selenium.sendtext('id','cnpj_cpf',Keys.TAB) # Tecla TAB
    Selenium.click('id','consultar_serasa') # clica em Consultar análise de crédito
    Selenium.click('id','consultar') # clica em consultar
    time.sleep(3)
    driver.switch_to.frame(0)
    time.sleep(1)


def extracao_dados(tabela):
    '''
    Tabelas Cadastradas:
    - dados cadastrais
    - ocorrencias
'''
    def dados_cadastrais():
        dados_cadastrais = {}
        try:
            dados_cadastrais['Data da Consulta'] = Selenium.find('xpath','//*[@id="QUADRO_ENTRADA-T1"]/tbody/tr[2]/td')
            dados_cadastrais['Código de Resposta'] = Selenium.find('xpath','//*[@id="QUADRO_ENTRADA-T1"]/tbody/tr[3]/td')
        except:
            dados_cadastrais['Data da Consulta'] = datetime.today().strftime('%d/%m/%Y %H:%M')
            dados_cadastrais['Código de Resposta'] = 'Não informado'
                    
        for x in range(1,10):
            driver.implicitly_wait(0.1)  # seconds
            try:
                xpa = '//*[@id="CONFIRMEI-T1"]/tbody/tr[{}]/th'.format(x)
                tabela = Selenium.find('xpath',xpa).replace(':','')
                # print(tabela)
                result = Selenium.find('xpath','//*[@id="CONFIRMEI-T1"]/tbody/tr[{}]/td'.format(x))
                print(f'{tabela}: {result}')
                dados_cadastrais[tabela] = result
            except:
                print('End')

        return dados_cadastrais

    def ocorrencias():
        resultado = {}
        for x in range(1,8):
            try:
                xpa = '//*[@id="QUADRO_RESUMO_CONSTA-T1"]/tbody/tr[{}]/td[1]'.format(x)
                tabela = Selenium.find('xpath',xpa)
                ocorrencias = []
                for y in range(2,5):
                    xpa = '//*[@id="QUADRO_RESUMO_CONSTA-T1"]/tbody/tr[{}]/td[{}]'.format(x,y)
                    info = Selenium.find('xpath',xpa)
                    ocorrencias.append(info)
                resultado[tabela] = {'Quantidade': ocorrencias[0],
                                     'Valor': ocorrencias[1],
                                     'Ultimo Registro': ocorrencias[2]}
            except:
                print('Fim')
                
        pprint.pp(resultado)
        return resultado
        

    if tabela == 'ocorrencias':
        resultado = ocorrencias()
        return resultado

    elif tabela == 'dados cadastrais':
        resultado = dados_cadastrais()
        return resultado

    else:
        print('Insira uma tabela válida')

    
    
    # ------------------- Fim -------------------

    
def algoritmo_CPF(array):
    data_ultimo_registro = []
    qtd = 0
    valor_total = 0
    for tipos_pendencia, dados_pendencia in array.items():
        if tipos_pendencia != 'Dados Cadastrais':
            quantidade = array[tipos_pendencia]['Quantidade']
            valor = array[tipos_pendencia]['Valor']
            ultimo_reg = array[tipos_pendencia]['Ultimo Registro']
            try:
                qtd += int(quantidade)
                valor_total += float(valor.replace('R$ ','').replace('.','').replace(',','.'))
                data_ultimo_registro.append(ultimo_reg)
            except:
                ...

    return qtd,valor_total,data_ultimo_registro
    # ------------------- Fim -------------------

def start_serasa_consulting(driver,cpf,numero_linha):

    aprovado = None
    login_ixc(driver)
    time.sleep(1)
    consulta_serasa(cpf)
    
    try:
        erro = Selenium.find('id','ERROS_B900-H2')
        sheet.update_acell('G' + str(numero_linha), erro)
        print(erro)
        driver.quit()
        exit()

    except:
        resumo = {}
        resumo['Dados Cadastrais'] = extracao_dados('dados cadastrais')
        resumo['Informações Restritivas'] = extracao_dados('ocorrencias')
        analise_restricao = algoritmo_CPF(resumo['Informações Restritivas'])
        qtd_registros = analise_restricao[0] if analise_restricao[0] > 0 else 0
        valor_aberto = analise_restricao[1]
        if valor_aberto > 0 or qtd_registros > 0:
            aprovado = '🚫🚫🚫 Com restrição 🚫🚫🚫'
            ultimo_registro = analise_restricao[2]
            valor_aberto = f'R$ {valor_aberto:_.2f}'
            valor_aberto = valor_aberto.replace('.',',').replace('_','.')
            resumo_final = f'Possui {qtd_registros} registros no valor total de {valor_aberto}'
        else:
            aprovado = '✅✅✅ Sem restrição ✅✅✅'
            resumo_final = ''
            
        dados_para_mensagem = ['CPF','Nome/Razão Social','Data de Nascimento','Situação', 'Data da Consulta','Código de Resposta']

        for tipo_dados in dados_para_mensagem:
            try:
                dado = resumo['Dados Cadastrais'][tipo_dados]
                print(dado)
            except:
                try:    
                    if tipo_dados == 'CPF':
                        info = Selenium.find('xpath','//*[@id="QUADRO_ENTRADA-T1"]/tbody/tr[4]/td')
                        resumo['Dados Cadastrais']['CPF'] = info.split('<br>')[0].replace('CPF ','')

                    elif tipo_dados == 'Nome/Razão Social':
                        info = Selenium.find('xpath','//*[@id="QUADRO_ENTRADA-T1"]/tbody/tr[4]/td')
                        resumo['Dados Cadastrais']['CPF'] = info.split('<br>')[1]
                        
                except:
                    if tipo_dados == 'CPF':
                        resumo['Dados Cadastrais']['CPF'] = 'Não informado'

                    elif tipo_dados == 'Nome/Razão Social':
                        resumo['Dados Cadastrais']['Nome/Razão Social'] = 'Não informado'
                    
                    
                if tipo_dados == 'Data de Nascimento':
                    resumo['Dados Cadastrais']['Data de Nascimento'] = 'Não informado'
                    
                elif tipo_dados == 'Situação':
                    resumo['Dados Cadastrais']['Situação'] = 'Não informado'

                elif tipo_dados == 'Data da Consulta':
                    resumo['Dados Cadastrais']['Data da Consulta'] = datetime.today().strftime('%d/%m/%Y %H:%M')

                elif tipo_dados == 'Código de Resposta':
                    resumo['Dados Cadastrais']['Código de Resposta'] = 'Não informado'


        mensagem = f'''Consulta *2* de *2* - *SERASA*

    Data e hora: {resumo['Dados Cadastrais']['Data da Consulta']}
    Código consulta: {resumo['Dados Cadastrais']['Código de Resposta']}

    CPF Consultado: {resumo['Dados Cadastrais']['CPF']}
    Nome: {resumo['Dados Cadastrais']['Nome/Razão Social']}
    Data de nascimento: {resumo['Dados Cadastrais']['Data de Nascimento']}
    Receita Federal: {resumo['Dados Cadastrais']['Situação']}

    Status: {aprovado}

    {resumo_final}'''
        sheet.update_acell('G' + str(numero_linha), aprovado)
        send_msg2(mensagem)
        driver.quit()
        exit()
    

driver = start()

