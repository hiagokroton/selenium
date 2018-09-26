import time
import pyautogui
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By

driver = webdriver.Chrome(executable_path=r"chromedriver.exe")
#driver = webdriver.Firefox()

driver.get("https://www.desafionotamaxima.com.br/login?locale=pt-BR")
driver.maximize_window()

# usuario do Thalles
username = driver.find_element_by_id("nome").send_keys("09261115667")
password = driver.find_element_by_id("senha").send_keys("desafiodoprofessor")

element = driver.find_element_by_xpath('//*[@id="new_user_session"]/fieldset[4]/input').click()

"""
# escrever usuario e senha
time.sleep(3)
pyautogui.moveTo(938, 393)
pyautogui.click()
pyautogui.typewrite('09261115667')

pyautogui.moveTo(930, 487)
pyautogui.click()
pyautogui.typewrite('desafiodoprofessor')

pyautogui.moveTo(1220, 579)
pyautogui.click()
"""





"""
###################################### abrir pagina de chamado
element = driver.find_element_by_xpath('//*[@id="js-dtour-step-two"]')
element.click()

element = driver.find_element_by_xpath('//*[@id="js-new-design-tour"]/div[2]/div/ul/li[2]/ul/li[1]/a')
element.click()
"""


# enquanto o link ta quebrado
driver.get("https://www.desafionotamaxima.com.br/mail_messages?locale=pt-BR")



###################################### filtrar chamados
# filtrar por status
elem_filtro1 = '//*[@id="accordion1"]/div/div[1]/h4/a'
elem_aberto = '//*[@id="collapseStatus"]/div/div/div[1]/div[2]/*[@id="array_status_"]'

time.sleep(0.5)
element = driver.find_element_by_xpath(elem_filtro1).click()

###################################### esperar abrir o filtro
result = None
while result is None:
    try:
        # connect
        element = driver.find_element_by_xpath(elem_aberto).click()
        result = 1
    except:
         pass

time.sleep(0.5)
element = driver.find_element_by_xpath(elem_filtro1).click()

###################################### filtrar por motivo detalhado
elem_filtro2 = '//*[@id="accordion2"]/div/div[1]/h4/a'
elem_dificuldade_acesso = '//*[@id="collapseSMotivo"]/div/div/div[3]/div[5]//*[@id="array_subtype_"]'

time.sleep(0.5)
element = driver.find_element_by_xpath(elem_filtro2).click()

# esperar abrir o filtro
result = None
while result is None:
    try:
        # connect
        element = driver.find_element_by_xpath(elem_dificuldade_acesso).click()
        result = 1
    except:
         pass

time.sleep(3)
element = driver.find_element_by_xpath(elem_filtro2).click()

###################################### filtrar alunos
elem_filtro_alunos = '/html/body/div[3]/div[2]/div/form/div[4]/div[2]/div/div[2]/a[1]'
time.sleep(1)

result = None
while result is None:
    try:
        # connect
        element = driver.find_element_by_xpath(elem_filtro_alunos).click()
        result = 1
    except:
         pass

time.sleep(3)

####################### ordenar por ordem de criação e esperar abrir o filtro
elem_data_criacao = '/html/body/div[3]/div[2]/div/form/div[4]/div[3]/table/thead/tr/th[7]/a'

result = None
while result is None:
    try:
        # connect
        element = driver.find_element_by_xpath(elem_data_criacao).click()
        result = 1
    except:
         pass


###################################### primeiro chamado para resposta
elem_chamado = '//*[@id="call_btn_0"]'
time.sleep(3)
element = driver.find_element_by_xpath(elem_chamado).click()






























######################################### ler cursos
df = pd.DataFrame()

for i in ['Olimpo', 'Siae', 'Colaborar']:
    df_aux = pd.read_excel('cursos_suporte_DNM.xlsx', sheet_name=i)
    df = df.append(df_aux)

df['chave_elegivel'] = df['curso'] + df['semestre'].astype(str)
lista_cursos_elegiveis = list(df['chave_elegivel'])





##################################### verificar data
time.sleep(4)
elem_data_criacao = '/html/body/div[2]/div[2]/div/div[2]/div/div/div/fieldset[4]'
elem_data_modificacao = '/html/body/div[2]/div[2]/div/div[2]/div/div/div/fieldset[5]'

data_criacao = driver.find_element_by_xpath(elem_data_criacao).text
data_criacao = data_criacao.replace('Data de criação\n\n', '')

data_modificacao = driver.find_element_by_xpath(elem_data_modificacao).text
data_modificacao = data_modificacao.replace('Data de última modificação\n\n', '')


if data_criacao == data_modificacao:

    ###################################### pegar o cpf
    elem_cpf = '/html/body/div[2]/div[2]/div/div[3]/div[2]/div/p[3]'
    time.sleep(2)
    cpf = driver.find_element_by_xpath(elem_cpf).text

    for i in ['CPF: ', '.', '-']:
        cpf = cpf.replace(i, '')


    ###################################### buscar cpf
    driver2 = webdriver.Chrome(executable_path=r"chromedriver.exe")
    #driver = webdriver.Firefox()
    time.sleep(3)
    #driver2.get("http://10.100.33.138:8000/signup/")
    driver2.get("http://10.100.34.64:8000/signup")
    driver2.maximize_window()

    username = driver2.find_element_by_xpath('//*[@id="id_cpf"]').send_keys(cpf)
    time.sleep(1)
    element = driver2.find_element_by_xpath('/html/body/div/div/div/form/center/input').click()

    ########### verificar base do cpf
    time.sleep(5)
    elem_primeira_linha_suporte = '/html/body/table/tbody/tr/th'

    try:
        driver2.find_element_by_xpath(elem_primeira_linha_suporte).text
        if 'DNM' in driver2.page_source:
            print('msg4')
        else:
            elem_curso = '/html/body/table/tbody/tr[1]/td[9]'
            elem_semestre = '/html/body/table/tbody/tr[1]/td[10]'
            curso = driver2.find_element_by_xpath(elem_curso).text
            semestre = driver2.find_element_by_xpath(elem_semestre).text

            tag_elegivel = curso + semestre

            if tag_elegivel in lista_cursos_elegiveis: # trazer curso e semestre e verificar se esta no excel
                print('msg2')
            else:
                print('msg1')



    except:
        print('msg3')








else:
    ##### mandar pra analise
    elem_responder = '/html/body/div[2]/div[2]/div/div[1]/div/div[3]/form/div/div/div/button'
    elem_analise = '/html/body/div[2]/div[2]/div/div[1]/div/div[3]/form/div/div/div/ul/li[1]/a'
    element = driver.find_element_by_xpath(elem_responder).click()

    result = None
    while result is None:
        try:
            # connect
            element = driver.find_element_by_xpath(elem_analise).click()
            result = 1
        except:
             pass

    ############################ apertar ok pra fechar chamado
    elem_ok = '/html/body/div[5]/div/div/div[2]/button'

    result = None
    while result is None:
        try:
            # connect
            element = driver.find_element_by_xpath(elem_ok).click()
            result = 1
        except:
             pass
