import time
import requests
import pyautogui
import pprint
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
    driver.maximize_window()
    return driver
    # ------------------- Fim -------------------
    

def send_msg2(bot_message):
    send_text = 'https://api.telegram.org/bot' + bot_token + '/sendMessage?chat_id=' + bot_chatID + '&parse_mode=Markdown&text=' + bot_message
    response = requests.get(send_text)
    res = response.json()
    if res['ok']:
        status_msg = 'Sim'
        return f'Enviado: {status_msg}  -  Mensagem ID:',res['result']['message_id']
    else:
        return 0

def delete_msg(id):
    if id != 0:
        send_text = 'https://api.telegram.org/bot' + bot_token + '/deleteMessage?chat_id=' + bot_chatID + '&message_id=' + str(id)
        response = requests.get(send_text)
        res = response.json()
        if res['ok']:
            status_msg = 'Sim'
            return f'Apagado? {status_msg}'
        else:
            return f'Not deleted  -  {res['description']}'


def coleta_dados_consulta():
    array = {}
    cabecalho = sheet.row_values(1)
    ultima_linha = len(sheet.get_all_values())
    for dados_para_consulta in range((ultima_linha-3),ultima_linha+1):
        print(dados_para_consulta)
        dados_cliente = sheet.row_values(dados_para_consulta)
        tabela_consulta = {}
        for i, dado in enumerate(dados_cliente):
            tabela_consulta[cabecalho[i]] = dado
        tabela_consulta['Número linha'] = dados_para_consulta
        array[dados_cliente[4]] = tabela_consulta
        

    return array

    #------------------- Fim -------------------

def start_consultation_SCPC(d,numero_linha):
    dados = get_info_SCPC(d[0],d[1],d[2],d[3],numero_linha)
    cpf = d[2]
    aprovado = None
    aprovado = 'Sim' if dados['Quantidade de Débitos'] == '-' or '' else 'Não'

    status = '✅✅✅ Sem restrição ✅✅✅' if aprovado == 'Sim' else '🚫🚫🚫 Com restrição 🚫🚫🚫'
    sheet.update_acell('F' + str(numero_linha), status)

    resumo_final = f'''
    Foram encontratos {dados['Quantidade de Débitos']} registros com
    valor total de {dados['Valor Negativação']},
    último registro em {dados['Última Negativação']}
    ''' if aprovado == 'Não' else ''

    mensagem = f''' Consulta *1* de *2* - *SCPC*

    Data e hora: {dados['Data e Hora Consulta']}
    Código consulta: {dados['Código Resposta']}

    CPF Consultado: {dados['CPF']}
    Nome: {dados['Nome Completo']}
    Data de nascimento: {dados['Data de Nascimento']}

    Status: {status}

    {resumo_final}'''

    send_msg2(mensagem)

    return aprovado,cpf,numero_linha

    #------------------- Fim -------------------

def get_info_SCPC(solicitante,nome_cliente,cpf,data_nascimento,numero_linha):
    driver.implicitly_wait(10)
    driver.get('https://www.acibarueri.com.br/')
    time.sleep(2)
    Selenium.sendtext('xpath','//*[@id="codigo"]',usuario_acib)
    Selenium.sendtext('xpath','//*[@id="senha"]',senha_acib)
    Selenium.click('name','logar')
    time.sleep(1)
    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)
    time.sleep(3)
    try:
        log_in = Selenium.find('class','m-nav__link-text')
        print('Welcome, ' + log_in.strip())
    except:
        try:
            mensagem_validacao = Selenium.find('xpath','//*[@id="localQuestoes"]/div[1]/div/div/label/b')
            if mensagem_validacao == 'Digite a sua frase de segurança':
                Selenium.sendtext('xpath', '//*[@id="chave_seg"]', pergunta_seguranca)
                Selenium.sendtext('xpath', '//*[@id="chave_seg"]', Keys.TAB)
                Selenium.click('xpath','//*[@id="btn_enviar"]')
                time.sleep(3)
            else:
                print('ACIB: Erro no login, falar com Brian')
                send_msg('ACIB: Erro no login, falar com Brian')
                time.sleep(1)
                driver.close()
                exit()
        except:
            try:
                erro = Selenium.find('xpath','//*[@id="formulario"]/div[1]/span')
                print('ACIB: ' + erro.strip())
                send_msg2('ACIB: ' + erro)
                sheet.update_acell('G' + str(numero_linha), erro)
                driver.quit()
                exit()
            except:
                print('Acesso liberado')

    try: Selenium.click('xpath','/html/body/div[4]/div/div/a')
    except: pass
    Selenium.click('xpath','/html/body/div[2]/div/div[1]/div/ul/li[1]/a/span')    
    Selenium.click('xpath','/html/body/div[2]/div/div[2]/div[2]/div[2]/div[2]/div/div/div[2]/div/ul/li[3]')
    Selenium.sendtext('id','solicita',solicitante)
    Selenium.sendtext('id','nome',nome_cliente)
    Selenium.sendtext('id','nasc',data_nascimento)
    Selenium.sendtext('id','nasc',Keys.TAB)
    Selenium.sendtext('id','documento',cpf)
    Selenium.sendtext('id','documento',Keys.TAB)
    try:
        erro = Selenium.find('xpath','//*[@id="documento-error"]')
        sheet.update_acell('F' + str(numero_linha), erro)
        sheet.update_acell('G' + str(numero_linha), erro)
        msg = ("Cliente: " + nome_cliente + "\n" +
               "Data de Nascimento: " + data_nascimento + "\n" +
               "CPF: " + cpf + "\n" +
               "Solicitante: " + solicitante + "\n" + "\n" +
               "Resultado: " + "😓 " + erro)
        send_msg2(msg)
        print(erro)
        driver.quit()
        exit()
    except:
        Selenium.click('id','btn_consultar')
        try:
            erro = Selenium.find('id','swal2-content')
            print(erro)
            msg = ("Cliente: " + nome_cliente + "\n" +
               "Data de Nascimento: " + data_nascimento + "\n" +
               "CPF: " + cpf + "\n" +
               "Solicitante: " + solicitante + "\n" + "\n" +
               "Resultado: " + "😓 " + erro)
            send_msg2(msg)
            sheet.update_acell('F' + str(numero_linha), erro)
            sheet.update_acell('G' + str(numero_linha), erro)
            driver.quit()
            exit()
            # update cell
        except:
            time.sleep(3)
            driver.switch_to.frame(0)
            time.sleep(2)

            dados = {}
            nome_completo = Selenium.find('xpath','//*[@id="container"]/div[1]/div/div[6]/table/tbody/tr[1]/td[1]')
            data_nascimento_scpc = Selenium.find('xpath','//*[@id="container"]/div[1]/div/div[6]/table/tbody/tr[2]/td[1]')
            cpf_scpc = Selenium.find('xpath','//*[@id="container"]/div[1]/div/div[6]/table/tbody/tr[1]/td[2]')
            data_hora_consulta = Selenium.find('class','help-back')
            codigo_resposta_scpc = Selenium.find('class','nroresp')
            
            dados['Nome Completo'] = nome_completo.split('<br>')[-1].strip()
            dados['CPF'] = cpf_scpc.split('CPF ')[-1].strip()
            dados['Data de Nascimento'] = data_nascimento_scpc.split('<br>')[-1].strip()
            dados['Data e Hora Consulta'] = data_hora_consulta.split(',')[1].split(' Ho')[0].strip()
            dados['Código Resposta'] = codigo_resposta_scpc.split(': ')[1].strip()
            dados['Quantidade de Débitos'] = Selenium.find('xpath','//*[@id="idtabelas"]/tbody/tr[2]/td[3]').strip()
            dados['Última Negativação'] = Selenium.find('xpath','//*[@id="idtabelas"]/tbody/tr[2]/td[4]').strip()
            dados['Valor Negativação'] = Selenium.find('xpath','//*[@id="idtabelas"]/tbody/tr[2]/td[5]').strip()
            dados['Quantidade de Consultas'] = Selenium.find('xpath','//*[@id="idtabelas"]/tbody/tr[3]/td[3]').strip()
            dados['Última Consulta'] = Selenium.find('xpath','//*[@id="idtabelas"]/tbody/tr[3]/td[4]').strip()
            dados['Valor da Consulta'] = Selenium.find('xpath','//*[@id="idtabelas"]/tbody/tr[3]/td[5]').strip()
            return dados
    
    #------------------- Fim -------------------


def consulting_verificator(dados):
    for x in dados.items():
        try:
            validador = x[1]['Resultado SCPC']
            cpf = x[0]
            print(f'O CPF {cpf} já consultado {validador}')
        except:
            cpf = x[0]
            validador = 'NC'

        if validador == 'NC':
            print(cpf + ' realizar consulta')
            cpf = x[0]
            sorted_dados = (x[1]['Solicitante'],
                            x[1]['Nome Cliente'],
                            x[1]['CPF - Sem pontos ou traços'],
                            x[1]['Data de Nascimento'])
            numero_linha =  x[1]['Número linha']
            return sorted_dados,numero_linha
            
        else:
            pass


# ----------------------- Inicio altoritmo final -----------------------

print('Starting Check ' + datetime.today().strftime('%H:%M'))
msg_inicio = send_msg2("🔎🔎")

dados = coleta_dados_consulta()
'''
resp = consulting_verificator(dados)
if resp == None:
    send_msg2('Sem CPF para consulta')
    delete_msg(msg_inicio[1])
    exit()
else:
    driver = start()
    resp = start_consultation_SCPC(resp[0],resp[1])

    if resp[0] == 'Sim':
        driver.quit()
        time.sleep(1)
        from ConsultationSerasa import *
        start_serasa_consulting(driver,resp[1],resp[2])
    else:
        driver.quit()
        exit()

    print('Check Completed ' + str(datetime.today().strftime('%d/%m/%Y %H:%M')))
    send_msg2("🆗!")
    exit()
'''

