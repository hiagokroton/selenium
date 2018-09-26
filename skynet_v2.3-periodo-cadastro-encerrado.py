import time
import pandas as pd
import tkinter
import re
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

############################ ABRIR PAGINAS
driver2 = webdriver.Chrome(executable_path=r"chromedriver.exe")
#driver = webdriver.Firefox()
driver2.maximize_window()
pagina_suporte = "http://10.100.37.242:8000/signup/"
driver2.get(pagina_suporte)


driver = webdriver.Chrome(executable_path=r"chromedriver.exe")
#driver = webdriver.Firefox()
driver.maximize_window()
driver.get("https://www.desafionotamaxima.com.br/login?locale=pt-BR")

time.sleep(5)

############################# LER INPUTS
# cursos elegiveis
df = pd.DataFrame()

for i in ['Olimpo', 'Siae', 'Colaborar']:
    df_aux = pd.read_excel('cursos_suporte_DNM.xlsx', sheet_name=i)
    df = df.append(df_aux)

df['chave_elegivel'] = df['curso'] + df['semestre'].astype(str)
lista_cursos_elegiveis = list(df['chave_elegivel'])

# status aluno
df_status_aluno = pd.read_csv('status_aluno.csv')

lista_cpfs_dnm = list(df_status_aluno['user_cpf'])








######################## LOGAR NO DNM E IR PRA PAGINA DE CHAMADO
# usuario do Thalles
username = driver.find_element_by_id("nome").send_keys("123456789000")
password = driver.find_element_by_id("senha").send_keys("123")

element = driver.find_element_by_xpath('//*[@id="new_user_session"]/fieldset[4]/input').click()

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


################################### NUMERO DE CHAMADOS
elem_numero_chamados = '//*[@id="control-mails-table"]/div[3]/div/div/ul/li[7]/a'
elem_numero_chamados2 = '//*[@id="control-mails-table"]/div[3]/div/div/ul/li[6]/a'
#try:
#    num_paginas = int(driver.find_element_by_xpath(elem_numero_chamados).text)
#except:
#    num_paginas = int(driver.find_element_by_xpath(elem_numero_chamados2).text)









######################################### MENSAGENS

msg1 = """E aí {},
Consultamos no sistema e vimos que você já possui acesso.
Acesse o DNM com o seu cpf nos campos login e senha ou clique em esqueci a minha senha, que mandamos uma senha nova para o seu e-mail!
Durante o login, ele atualiza a página e mostra algumas opções de semestres para logar, sempre de uma olhada se está em 2018.2.
Bons estudos!"""

msg2 = """E aí {},
Consultamos no sistema e não encontramos o seu cpf na base.
Verifique se você digitou corretamente o número do cpf, caso esteja, procure seu coordenador pois não te encontramos vinculado no sistema.
Bons estudos!"""

msg3 = """E aí {},
Checamos aqui e você realmente é um aluno DNM, estamos te devendo.
Nós atualizamos diariamente os usuários, verifique dentro de 24h se você conseguirá acessar!
Lembre-se: Para acessar o dnm use o seu cpf nos campos login e senha, sempre de uma olhada se está em 2018.2.
Bons estudos!"""

msg4 = """E aí {},
Infelizmente seu curso e semestre não participam do DNM 2018.2.
Se ficar com alguma dúvida do porque, procure alguém na sua unidade para lhe explicar como funciona a regra.
Estamos ansiosos para te receber nos próximos!
Bons estudos!"""

msg5 = """E aí {},
Verificamos que você é um aluno elegível, mas não está matriculado na turma da disciplina participante do DNM.
Confirme com seu coordenador se está tudo certo com suas disciplinas e turma, que ele poderá lhe ajudar em algum problema mais complexo.
Bons estudos!"""

msg6 = """E aí {},
Verificamos que você não enviou o CPF, favor abrir um novo chamado, informando seu CPF.
Bons estudos!"""

msg7 = """E aí {},
O cadastro de alunos encerrou dia 17/09.
Todos os alunos que se matricularem após a data 17/09, não poderão realizar mais o DNM 2018.2.
Para mais informações procure o seu Coordenador."""

# elementos para se movimentar nas paginas
elem_responder = '//*[@id="mail_message_form"]/div/div/div/button'
elem_ok = '/html/body/div[5]/div/div/div[2]/button'
elem_ok2 = '/html/body/div[6]/div/div/div[2]/button'
elem_analise = '//*[@id="mail_message_form"]/div/div/div/ul/li[1]/a'
xpath_aluno = 2




def apertar_ok():
    ############################ apertar ok pra fechar chamado
    result = None
    while result is None:
        try:
            # connect
            element = driver.find_element_by_xpath(elem_ok).click()
            result = 1
        except:
            pass

        try:
            # connect
            element = driver.find_element_by_xpath(elem_ok2).click()
            result = 1
        except:
            pass

def mandar_pra_analise():
    element = driver.find_element_by_xpath(elem_responder).click()
    result = None
    while result is None:
        try:
            # connect
            element = driver.find_element_by_xpath(elem_analise).click()
            result = 1
        except:
            pass

    apertar_ok()

def verificar_load(elem_verificar):
    result = None
    while result is None:
        try:
            # connect
            element = driver.find_element_by_xpath(elem_verificar).click()
            result = 1
        except:
            pass

















for j in range(10000):
    ################################# IR PRA CHAMADO MANUAL
    # driver.get(j)
    driver2.get(pagina_suporte)

    ##################################### verificar data
    #time.sleep(4)
    elem_data_criacao = '//*[@id="sidebar_adaptativos"]/div/div/div/fieldset[4]'
    elem_data_modificacao = '//*[@id="sidebar_adaptativos"]/div/div/div/fieldset[5]'

    # verificar se carregou
    verificar_load(elem_data_criacao)

    data_criacao = driver.find_element_by_xpath(elem_data_criacao).text
    data_criacao = data_criacao.replace('Data de criação\n\n', '')

    data_modificacao = driver.find_element_by_xpath(elem_data_modificacao).text
    data_modificacao = data_modificacao.replace('Data de última modificação\n\n', '')

    ################################## VERIFICAR SEM ACESSO
    elem_aluno = '//*[@id="sidebar_adaptativos"]/div/div/div/fieldset[3]'
    aluno = driver.find_element_by_xpath(elem_aluno).text
    aluno = aluno.replace('Aluno\n\n', '')

    #if data_criacao == data_modificacao and aluno == 'SEM ACESSO':
    if aluno == 'SEM ACESSO':
        try:
            ##################################### SALVAR NOME DO ALUNO
            elem_nome = '//*[@id="full-page-wrapper"]/div[{}]/div/div[3]/div[2]/div/p[1]'
            nome = driver.find_element_by_xpath(elem_nome.format(xpath_aluno)).text
            nome = nome.replace('Nome: ', '')
            nome = nome.split()[0].title()

            ###################################### PEGAR CPF
            elem_cpf = '//*[@id="full-page-wrapper"]/div[{}]/div/div[3]/div[2]/div/p[3]'
            time.sleep(2)
            cpf = driver.find_element_by_xpath(elem_cpf.format(xpath_aluno)).text
            cpf = re.sub('[^0-9]', '', cpf)
            try:
                cpf_check_dnm = int(cpf)
            except:
                pass

            if cpf == '':
                mensagem = driver.find_element_by_id("mail_message_content").send_keys(msg6.format(nome))

            ################################## ver se o cpf ja esta no DNM
            elif cpf_check_dnm in lista_cpfs_dnm:
                mensagem = driver.find_element_by_id("mail_message_content").send_keys(msg1.format(nome))

            else:
                ###################################### BUSCAR CPF
                username = driver2.find_element_by_xpath('//*[@id="id_cpf"]').send_keys(cpf)
                element = driver2.find_element_by_xpath('/html/body/div/div/div/form/center/input').click()

                ####################### VERIFICAR CPF NO SUPORTE
                elem_primeira_linha_suporte = '/html/body/table/tbody/tr/th'
                elem_head_cpf = '/html/body/table/thead/tr/th[2]'

                ######################### VERIFICAR SE JA CARREGOU
                try:
                    # esperar aparecer a base do suporte
                    result = None
                    while result is None:
                        try:
                            # connect
                            driver2.find_element_by_xpath(elem_head_cpf).text
                            result = 1
                        except:
                            pass

                    driver2.find_element_by_xpath(elem_primeira_linha_suporte).text

                    if 'DNM' in driver2.page_source:
                        ############## ALUNO TEM DISCIPLINA VINCULADA
                        mensagem = driver.find_element_by_id("mail_message_content").send_keys(msg7.format(nome))
                    else:
                        ################### VERIFICAR SE É ELEGIVEL
                        elem_curso = '/html/body/table/tbody/tr[1]/td[9]'
                        elem_semestre = '/html/body/table/tbody/tr[1]/td[10]'
                        curso = driver2.find_element_by_xpath(elem_curso).text
                        semestre = driver2.find_element_by_xpath(elem_semestre).text

                        tag_elegivel = curso + semestre

                        if tag_elegivel in lista_cursos_elegiveis: # trazer curso e semestre e verificar se esta no excel
                            mensagem = driver.find_element_by_id("mail_message_content").send_keys(msg7.format(nome))
                        else:
                            mensagem = driver.find_element_by_id("mail_message_content").send_keys(msg4.format(nome))

                except:
                    ############################# NAO VEIO TABELA NO SUPORTE
                    mensagem = driver.find_element_by_id("mail_message_content").send_keys(msg2.format(nome))


            time.sleep(0.5)
            ######################################### clicar em respondido
            element = driver.find_element_by_xpath(elem_responder).click()
            elem_respondido = '//*[@id="mail_message_form"]/div/div/div/ul/li[3]/a'

            result = None
            while result is None:
                try:
                    # connect
                    element = driver.find_element_by_xpath(elem_respondido).click()
                    result = 1
                except:
                    pass

            apertar_ok()

        except:
            #messagebox.showerror("Error", "Error message")
            #print('erro2')
            #break
            mandar_pra_analise()



    else:
        #messagebox.showerror("Error", "Error message")
        #print('erro1')
        #break
        #################################### mandar pra analise
        mandar_pra_analise()

    if xpath_aluno == 2:
        xpath_aluno += 1
    else:
        pass
