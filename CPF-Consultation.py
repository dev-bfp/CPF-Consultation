import time
import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from datetime import datetime
from Tokens import *

import pyautogui
# LIB DO GOOGLE SHEETS
import gspread
# LIB DE AUTENTICAÇÃO DA GOOGLE
from oauth2client.service_account import ServiceAccountCredentials

# URL GOOGLE
scope = ["https://www.googleapis.com/auth/drive"]
credentials_google = ServiceAccountCredentials.from_json_keyfile_name(caminho_local_credentials, scope)
client = gspread.authorize(credentials_google)
# CÓDIGO MATRIZ
nmSheets = "VENDAS REDFIBRA"
sheet = client.open(nmSheets).worksheet("Consulta de CPF")


CPF_error = ""
date_notfound = ""


# ------------------------ Webdriver ------------------------
driver = webdriver.Chrome(service=Service(diretorio_chromedriver))

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


# ------------------------ Sheets ------------------------
def lphr():
    i = 0
    while True:

        if rz5[i].value == "" and rz6[i].value == "" and rz1[i].value != "" and rz2[i].value != "" \
                and rz3[i].value != "" and rz4[i].value != "":
            msg_consult = ("🔎 CONSULTANDO... 🔎" + "\n" + "\n" +
                           "Solicitante: " + rz1[i].value + "\n" +
                           "Cliente: " + rz2[i].value + "\n" +
                           "Data de Nascimento: " + rz3[i].value + "\n" +
                           "CPF: " + rz4[i].value)
            send_msg2(msg_consult)
            selenium(i + 2)
            continue

        elif rz1[i].value == "" and rz2[i].value == "" and rz3[i].value == "" and rz4[i].value == "":
            date_notfound = send_msg2("Sem CPF para consulta")
            print(date_notfound)
            # send_msg(msg)
            break
        elif rz1[i].value == "" or rz2[i].value == "" or rz3[i].value == "" or rz4[i].value == "":
            if rz5[i].value != "Dados insuficientes para consulta" and rz6[
                i].value != "Dados insuficientes para consulta":
                sheet.update_acell('E' + str(i + 2), "Dados insuficientes para consulta")
                sheet.update_acell('F' + str(i + 2), "Dados insuficientes para consulta")
                msg_consult = ("🔎 CONSULTANDO... 🔎" + "\n" + "\n" +
                               "Solicitante: " + rz1[i].value + "\n" +
                               "Cliente: " + rz2[i].value + "\n" +
                               "Data de Nascimento: " + rz3[i].value + "\n" +
                               "CPF: " + rz4[i].value)
                msg = "Dado insuficientes para consulta"
                delete_msg(msg_start)
                send_msg2(msg_consult)
                send_msg2(msg)
                # send_msg(msg)

        i += 1


# ------------------------ Selenium ------------------------
def selenium(x):
    razao = sheet.range(x, 2, x, 5)
    Solicitante = razao[0].value
    Nome_cliente = razao[1].value
    Data_nascimento = razao[2].value
    CPF = razao[3].value
    print(Solicitante)
    print(Nome_cliente)
    print(Data_nascimento)
    print(CPF)

    driver.maximize_window()
    link = 'https://www.acibarueri.com.br/'
    driver.get(link)
    time.sleep(3)

    driver.find_element(by=By.XPATH, value='//*[@id="codigo"]').send_keys(Keys.ESCAPE)
    time.sleep(1)
    driver.find_element(by=By.XPATH, value='//*[@id="codigo"]').send_keys(usuario_acib)
    time.sleep(1)
    driver.find_element(by=By.XPATH, value='//*[@id="senha"]').send_keys(senha_acib)
    time.sleep(2)
    driver.find_element(by=By.XPATH,
                        value='/html/body/header/div[1]/div/div/div[2]/ul/li[3]/div[1]/form/input[3]').click()
    time.sleep(5)
    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)

    time.sleep(3)

    while True:

        act = ActionChains(driver)
        act.send_keys(Keys.ESCAPE).perform()

        try:
            # driver.find_element(by=By.XPATH, value='//*[@id="m_aside_left_offcanvas_toggle"]').click()
            # driver.find_element(by=By.XPATH, value='/html/body/div[2]/div/div[1]/div/ul/li[1]/a/span').send_keys(Keys.ESCAPE)

            time.sleep(5)
            try: 
                driver.find_element(by=By.XPATH, value='/html/body/div[4]/div/div/a').click()
            except: 
                pass
            driver.find_element(by=By.XPATH, value='/html/body/div[2]/div/div[1]/div/ul/li[1]/a/span').click()
            time.sleep(2)
            driver.find_element(by=By.XPATH,
                                value='/html/body/div[2]/div/div[2]/div[2]/div[2]/div[2]/div/div/div[2]/div/ul/li[3]').click()
            time.sleep(2)
            driver.find_element(by=By.XPATH,
                                value='/html/body/div[2]/div/div[2]/div[2]/div[2]/div[2]/form/div[1]/div[1]/div/div/input').send_keys(
                Solicitante)
            time.sleep(1)
            driver.find_element(by=By.XPATH,
                                value='/html/body/div[2]/div/div[2]/div[2]/div[2]/div[2]/form/div[1]/div[2]/div[1]/div/input').send_keys(
                Nome_cliente)
            time.sleep(1)
            driver.find_element(by=By.XPATH,
                                value='/html/body/div[2]/div/div[2]/div[2]/div[2]/div[2]/form/div[1]/div[2]/div[2]/div/input').send_keys(
                Data_nascimento)
            time.sleep(1)
            driver.find_element(by=By.XPATH,
                                value='/html/body/div[2]/div/div[2]/div[2]/div[2]/div[2]/form/div[1]/div[2]/div[2]/div/input').send_keys(
                Keys.TAB)
            driver.find_element(by=By.XPATH,
                                value='/html/body/div[2]/div/div[2]/div[2]/div[2]/div[2]/form/div[1]/div[3]/div/div/input').clear()
            driver.find_element(by=By.XPATH,
                                value='/html/body/div[2]/div/div[2]/div[2]/div[2]/div[2]/form/div[1]/div[3]/div/div/input').send_keys(
                CPF)
            driver.find_element(by=By.XPATH,
                                value='/html/body/div[2]/div/div[2]/div[2]/div[2]/div[2]/form/div[1]/div[3]/div/div/input').send_keys(
                Keys.TAB)
            try:
                time.sleep(2)
                driver.find_element(by=By.XPATH, value='//*[@id="documento-error"]').get_attribute("innerHTML")
                sheet.update_acell('F' + str(x), 'CPF Inválido')
                sheet.update_acell('G' + str(x), 'CPF Inválido')
                CPF_error = send_msg2("CPF Inválido")
                delete_msg(msg_start)
                print(CPF_error)
                driver.quit()
                window_zero = driver.window_handles[0]
                driver.switch_to.window(window_zero)
                driver.quit()
                print("end code")

            except:
                time.sleep(2)
                driver.find_element(by=By.XPATH, value='//*[@id="btn_consultar"]').click()
                time.sleep(10)
                try:
                    time.sleep(2)
                    msg_erro = driver.find_element(by=By.XPATH, value='//*[@id="swal2-content"]').get_attribute(
                        "innerHTML")
                    print(msg_erro)
                    msg_tlg = ("Cliente: " + Nome_cliente + "\n" +
                               "Data de Nascimento: " + Data_nascimento + "\n" +
                               "CPF: " + CPF + "\n" +
                               "Solicitante: " + Solicitante + "\n" + "\n" +
                               "Resultado: " + "😓 " + msg_erro)
                    send_msg2(msg_tlg)
                    sheet.update_acell('G' + str(x), msg_erro)
                    sheet.update_acell('F' + str(x), msg_erro)
                    driver.quit()
                    window_zero = driver.window_handles[0]
                    driver.switch_to.window(window_zero)
                    driver.quit()
                    time.sleep(3)
                    print("end code")





                except:
                    time.sleep(5)
                    driver.find_element(by=By.XPATH, value='/html/body/div[4]/div/div/div[1]/button[3]').click()
                    time.sleep(2)

                    try:
                        driver.find_element(by=By.XPATH, value='//*[@id="m_aside_left_offcanvas_toggle"]').click()
                        time.sleep(1)
                    except:
                        time.sleep(1)

                    driver.find_element(by=By.XPATH, value='//*[@id="numProg00000000000002"]/i').click()
                    time.sleep(6)
                    driver.find_element(by=By.XPATH,
                                        value='/html/body/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div[1]/div/ul/li[1]').click()
                    time.sleep(6)
                    resultado = driver.find_element(by=By.XPATH,
                                                    value='/html/body/div[2]/div/div[2]/div[2]/div/div[3]/div[1]/div/div[2]/div[2]/div/div/div[2]/div/table/tbody/tr[1]/td[9]').get_attribute(
                        "innerHTML")
                    restri = ""
                    if resultado == "Sim":
                        restri = "🚫🚫🚫 Com restrição 🚫🚫🚫"
                    elif resultado == "Não":
                        restri = "✅✅✅ Sem restrição ✅✅✅"

                    sheet.update_acell('F' + str(x), datetime.today().strftime('%d/%m/%Y %H:%M'))
                    sheet.update_acell('G' + str(x), restri)
                    print(restri)
                    msg = ("Cliente: " + Nome_cliente + "\n" +
                           "Data de Nascimento: " + Data_nascimento + "\n" +
                           "CPF: " + CPF + "\n" +
                           "Solicitante: " + Solicitante + "\n" + "\n" +
                           "Resultado: " + restri)
                    send_msg2(msg)
                    driver.quit()
                    window_zero = driver.window_handles[0]
                    driver.switch_to.window(window_zero)
                    driver.quit()
                    print("end code")

        except:
            time.sleep(4)
            mensagem_validacao = driver.find_element(by=By.XPATH,
                                                     value='//*[@id="localQuestoes"]/div[1]/div/div/label/b').get_attribute(
                "innerHTML")
            if mensagem_validacao == 'Digite a sua frase de segurança':
                driver.find_element(by=By.XPATH, value='//*[@id="chave_seg"]').send_keys(pergunta_seguranca)
                driver.find_element(by=By.XPATH, value='//*[@id="chave_seg"]').send_keys(Keys.TAB)
                driver.find_element(by=By.XPATH, value='//*[@id="btn_enviar"]').click()
                time.sleep(5)
                continue


            else:
                msg = "Validar pergunta de segurança"
                send_msg(msg)
                time.sleep(120)
                driver.close()
        # time.sleep(3)
        # driver.find_element(by=By.XPATH, value='//*[@id="m_aside_left_offcanvas_toggle"]').click()
        # driver.find_element(by=By.XPATH, value='/html/body/div[4]/div').send_keys(Keys.ESCAPE)
        # time.sleep(3)
    else:
        print("end all")


print('Starting Check ' + datetime.today().strftime('%H:%M'))
msg_start = send_msg2("🔎🔎")

rz1 = sheet.range(2, 2, 10000, 2)
rz2 = sheet.range(2, 3, 10000, 3)
rz3 = sheet.range(2, 4, 10000, 4)
rz4 = sheet.range(2, 5, 10000, 5)
rz5 = sheet.range(2, 6, 10000, 6)
rz6 = sheet.range(2, 7, 10000, 7)

lphr()
print('Check Completed ' + str(datetime.today().strftime('%d/%m/%Y %H:%M')))
msg_end = send_msg2("🆗!")
sheet.update_acell('I2', datetime.today().strftime('%d/%m/%Y %H:%M'))

time.sleep(1)
delete_msg(msg_start)
delete_msg(msg_end)
