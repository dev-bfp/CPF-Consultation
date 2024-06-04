import time
import pyautogui
import pprint
import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from datetime import datetime
from Tokens import *


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

        Selenium.click('id','entrar')  # clica em entrar
        try:
            driver.implicitly_wait(10)  # seconds
            Selenium.click('xpath','//*[@id="slide_0"]/div[4]/vg-button')  # fecha pop-up
            Selenium.click('xpath','//*[@id="slide_1"]/div[4]/vg-button[2]')

        except:
            time.sleep(0.5)
    except:
        try:
            driver.implicitly_wait(10)  # seconds
            Selenium.click('xpath','//*[@id="warning"]/vg-body/div/vg-button[2]')  # fecha pop-up
        except:
            time.sleep(1)
            print('Erro login')
            driver.quit()
    # ------------------- Fim -------------------

def consulta_serasa(CPF):
    driver.implicitly_wait(10)  # seconds
    Selenium.click('id','id_input_menu')  # abre pesquisa
    Selenium.sendtext('id','id_input_menu', 'cliente')  # insere pesquisa
    Selenium.click('xpath','//*[@id="grupo_menu04400d48d04acd3599cf545dafbb90ed"]/ul/li[1]/a') # clica na opção
    Selenium.click('xpath','//*[@id="1_grid"]/div/div[2]/div[1]/button[1]') # clica em novo
    Selenium.clear('id','cnpj_cpf') # limpa o campo
    time.sleep(1)
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
        for x in range(1,10):
            driver.implicitly_wait(0.1)  # seconds
            try:
                xpa = '//*[@id="CONFIRMEI-T1"]/tbody/tr[{}]/th'.format(x)
                tabela = Selenium.find('xpath',xpa).replace(':','')
                print(tabela)
                result = Selenium.find('xpath','//*[@id="CONFIRMEI-T1"]/tbody/tr[{}]/td'.format(x))
                print(result)
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

aprovado = None
driver = start()
login_ixc(driver)
time.sleep(2)
cpf = input('Insira o CPF: ')
consulta_serasa(cpf)
dados_cadastrais = extracao_dados('dados cadastrais')
ocorrencias = extracao_dados('ocorrencias')
resumo = {}
resumo['Dados Cadastrais'] = dados_cadastrais
resumo['Informações Restritivas'] = ocorrencias
analise_restricao = algoritmo_CPF(resumo['Informações Restritivas'])

qtd_registros = analise_restricao[0] if analise_restricao[0] > 0 else ''
valor_aberto = analise_restricao[1]
if valor_aberto > 0 or qtd_registros > 0:
    aprovado = '🚫🚫🚫 Com restrição 🚫🚫🚫'
else:
    aprovado = '✅✅✅ Sem restrição ✅✅✅'

ultimo_registro = analise_restricao[2]
valor_aberto = f'R$ {valor_aberto:_.2f}'
valor_aberto = valor.replace('.',',').replace('_','.')
if aprovado == '🚫🚫🚫 Com restrição 🚫🚫🚫':
    resumo_final = f'Possui {qtd_registros} registros no valor total de {valor_aberto}' 
else:
    resumo_final = ''

mensagem = (f'''------- *Serasa* -------

Data: {'adefinir'}
Código consulta: {'adefinir'}
CPF: {resumo['Dados Cadastrais']['CPF']}
Nome: {resumo['Dados Cadastrais']['Nome/Razão Social']}
Data de nascimento: {resumo['Dados Cadastrais']['Data de Nascimento']}
Situação Receita Federal: {resumo['Dados Cadastrais']['Situação']}

Status: {aprovado}

{resumo_final}
''')
send_msg2(mensagem)

    
