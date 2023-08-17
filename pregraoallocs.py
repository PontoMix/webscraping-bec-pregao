import time

#Importando o debugger do python para criar breakpoints e verificar se está funcionando o código criado
import pdb

#Definir um determinado período de tempo para que um elemento apareça antes de prosseguir é com Web Driver Wait. 
from selenium.webdriver.support.wait import WebDriverWait

#EC significa condições esperadas, que são condições que devem ser cumpridas para que uma determinada ação possa ser tomada
from selenium.webdriver.support import expected_conditions as EC

import os

##Importando o webdriver da biblioteca Selenium para acessar a Internet e carregar páginas
from selenium import webdriver

##Importando o Options para poder manipular o webdriver e suas propriedades, 
#como executar várias operações, desativar extensões, desativar pop-ups, etc
from selenium.webdriver.chrome.options import Options

#Importando as chaves especiais para utilizar teclas no teclado ao usar Selenium e 
#apagar textos pré-preenchido em campos de entradas
from selenium.webdriver.common.keys import Keys

#Importando o By para localizar elementos em uma página web
from selenium.webdriver.common.by import By

#Importando ActionChains para utilizar o mouse e teclas do teclado
from selenium.webdriver.common.action_chains import ActionChains

#Pandas para criar, manipular e visualizar tabelas
import pandas as pd

#OpenXL para escrever arquivos .xls (Excel) com o Pandas por meio do módulo "Xmlt" para arquivos .xls
import openpyxl 

#Time para utilizar datas quando salvar os arquivos .xls
import time

#Pegando a data e horário de hoje, no momento que criou o arquivo .xls 
TodayDate = time.strftime("%d-%m-%Y %H-%M-%S")
DateSheet = time.strftime("%d-%m-%Y")

#Criando um nome padronizado para os arquivos .xls que terá os detalhes de cada OC
excelfilename = "Pregão Eletrônico Completo - " + TodayDate + ".xls"

username = os.getlogin()  # Obtém o nome do usuário da máquina
home_dir = os.path.expanduser("~")

path_with_filename = os.path.join("C:\\", "WebScraping Licitações - Pregão", "Detalhes Produtos - Pregão", excelfilename)

#Uma pasta e caminho específico (absoluto) para salvar os arquivos .xls gerados de cima
#path_with_filename = r"C:/Users/João Pedro/OneDrive - Brasil/Ponto Mix/Webscraping Licitações/Pregão Eletrônico/Detalhes OCs/"+excelfilename  

#Criando um nome padronizado para os arquivos .xls que terá as informações da tabela das OCs
excelfilenameallocs = "Tabela OC Completa - " + TodayDate + ".xls" 

path_with_filenameallocs = os.path.join("C:\\", "WebScraping Licitações - Pregão", "Detalhes Produtos - Pregão", "Tabela OCs - Pregão", excelfilenameallocs)

#Uma pasta e caminho específico (absoluto) para salvar os arquivos .xls gerados de cima
#path_with_filenameallocs = r"C:/Users/João Pedro/OneDrive - Brasil/Ponto Mix/Webscraping Licitações/Pregão Eletrônico/Detalhes OCs/"+excelfilenameallocs

#Criando um nome padronizado para as folhas do .xls
excelsheet = "Pregão Eletrônico Completo - " + DateSheet

#Importando as variáveis de ambiente para utilizar com segurança o login e senha do usuário
from config import database_infos

get_login = database_infos['login']
get_pass = database_infos['password']
get_username_pc = database_infos['username_pc']


                                                        ##########################################################
                                                         ##Função p/ rodar todas as etapas do Pregão Eletrônico##
                                                        ##########################################################
def bec_pregaoeletronico():

    browser_driver = webdriver.Chrome()

    #Fazendo solicitação para abrir e navegar na página da BEC
    browser_driver.get("https://www.bec.sp.gov.br/BECSP/Home/Home.aspx")

    #Inicializando o WebDriverWait
    waitWDW = WebDriverWait(browser_driver, 10)

    #Maximizando a Tela do Browser
    browser_driver.maximize_window()

    #Confirmando que é o site correto aquele que está aberto
    assert "BEC" in browser_driver.title

    #Procurando a tag certa do botão "Negociações Eletrônicas" 
    btn_ne = browser_driver.find_element(By.LINK_TEXT, "Negociações Eletrônicas")
 
    ##Fazendo com que clique no botão
    btn_ne.send_keys(Keys.RETURN)

    #Procurando as tags certas com XPATH e preenchendo os campos "CNJP/CPF" e "Senha"
    login = browser_driver.find_element(By.XPATH, "//input[@id='TextLogin']") #Se parar de funcionar, utilize a class="TextLogin" ou o id="TextLogin"
    login.send_keys(get_login)

    password = browser_driver.find_element(By.XPATH, "//*[@id='TextSenha']") #Se parar de funcionar, utilize a class="TextSenha" ou o id="TextSenha"
    password.send_keys(get_pass)

    #Marcando a caixa de declaração
    statement_box = browser_driver.find_element(By.XPATH, "//*[@id='chkAceite']") #Se parar de funcionar, utilize a class="chkAceite" ou o id="chkAceite"
    statement_box.click()

    #Clicando no botão de entrar
    btn_enter = browser_driver.find_element(By.ID, "Btn_Confirmar") #Se parar de funcionar, utilize o id="Btn_Confirmar"
    btn_enter.click()

    current_url = browser_driver.current_url
    
    if current_url == "https://www.bec.sp.gov.br/fornecedor_ui/TermoResponsabilidade.aspx?Dzqeio6gALuoR%2flQf2tFB6zBkp9ETq5P44%2bgrURdFf66JmFgqUpWHFjTKO2RLNZR":
        waitWDW = WebDriverWait(browser_driver, 10)
        reconfirm_checkbox = browser_driver.find_element(By.ID, "//*[@id='ctl00_c_area_conteudo_chkDeclaracao']")
        reconfirm_checkbox.click()
        ok_button = browser_driver.find_element(By.ID, "//*[@id='ctl00_c_area_conteudo_Button1']")
        ok_button.click()
        #Passando o mouse por cima da lista "Participar"
        join_menu_list = waitWDW.until(EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Participar']")))
        actions = ActionChains(browser_driver)
        actions.move_to_element(join_menu_list).pause(2).perform()
        
        #Escolhendo o item da lista certa, que é o Pregão Eletrônico e clicando
        pe_item_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//li[@id='21003']//a[contains(text(),'Pregão Eletrônico')]")))
        pe_item_list.click()

        pe_btn_search = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='pesquisa']")))  #Se parar de funcionar, utilizar o id="ctl00_conteudo_Pesquisa", css_selector="#pesquisa" ou text_link="Pesquisar"
        pe_btn_search.click() 
    
    else:
        #Passando o mouse por cima da lista "Participar"
        join_menu_list = waitWDW.until(EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Participar']")))
        actions = ActionChains(browser_driver)
        actions.move_to_element(join_menu_list).pause(2).perform()
        
        #Escolhendo o item da lista certa, que é o Pregão Eletrônico e clicando nele
        pe_item_list = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//li[@id='21003']//a[contains(text(),'Pregão Eletrônico')]")))
        pe_item_list.click()

        pe_btn_search = waitWDW.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='pesquisa']")))  #Se parar de funcionar, utilizar o id="ctl00_conteudo_Pesquisa", css_selector="#pesquisa" ou text_link="Pesquisar"
        pe_btn_search.click() 

    #Procurando todos os elementos da tabela do Pregão Eletrônico
    pe_oc_infos = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_conteudo_Tabela']/tbody/tr/td[2]/a")[1:]
    pe_op_forescat_infos = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_conteudo_Tabela']/tbody/tr/td[4]")[1:]
    pe_purchasing_unit_infos = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_conteudo_Tabela']/tbody/tr/td[6]")[1:]

    ###Lista para armazenar os resultados da coleta de dados do Pregão Eletrônico###
    pe_infos_result = []

    #Criando uma iteração (loop) para pegar os elementos específicos da tabela do Pregão Eletrônico
    for i in range(len(pe_oc_infos)):
        temporary_data={"Oferta de Compra":pe_oc_infos[i].text,
                       "Previsão de Abertura":pe_op_forescat_infos[i].text,
                       "Unidade Compradora":pe_purchasing_unit_infos[i].text}

        #Colocando as informações coletadas na lista criada
        pe_infos_result.append(temporary_data)

    #Lista para armazenar os resultados da coleta de dados da descrição, quantidade, uf, telefone e e-mails do Pregão
    pe_details_oc = []
    #Lista para armazenar os emails dos responsáveis
    pe_emails_oc = []
    #Lista para armazenar os telefones dos responsáveis
    pe_tels_oc = []

                                            ######################################################################################
                                            ###Criando um loop para acessar os links da lista e pegar as informações adicionais### 
                                            ######################################################################################
    
    #A variável global de iteração que será usada para percorrer todos os links da tabela de OCs
    global iterator
    iterator = 1
    
    #Encontrando a tabela do Pregão com as OCs para pegar as linhas e criar um loop
    table_oc = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_conteudo_Tabela']")
    rows = table_oc.find_elements(By.XPATH, "//tbody/tr/td[2]")
   
    n = 0
    for n in range(iterator, len(rows)):
        
                        link = waitWDW.until(EC.element_to_be_clickable((By.XPATH, f"//tr[{iterator+1}]/td[2]/a")))

                        #Pressionado CRTL e clicando no link
                        ActionChains(browser_driver).key_down(Keys.CONTROL).click(link).perform()

                        # Mudando para a nova aba aberta
                        browser_driver.switch_to.window(browser_driver.window_handles[-1])
                        time.sleep(2)

                        #Esperando a página carregar até aparecer o botão do Pregão
                        #Para clicar e não carregar uma nova aba, mas permanecer na mesma eu precisarei usar o get
                        element = WebDriverWait(browser_driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='topMenu']/li[3]/a"))) 
                        link_btn_pregao = element.get_attribute("href")
                        browser_driver.get(link_btn_pregao)
                        time.sleep(2)
                        
                        try:

                                    #Tabela dos detalhes da OC
                                    table_oc_details = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_conteudo_dg']")
                                    #Linhas da tabela    
                                    rows_oc_details = table_oc_details.find_elements(By.XPATH, "//*[@id='ctl00_conteudo_dg']/tbody/tr")[1:] #Pula a primeira linha da tabela (Cabeçalho)

                                    number_ocs = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_wucOcFicha_txtOC']")

                                    col_3_value = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_conteudo_dg']/tbody/tr/td[3]") [1:]

                                    if len(rows_oc_details)>1:
                                    
                                        #Criando variáveis para pegar os dados das colunas da tabela
                                        col_4_values = []
                                        col_5_values = []
                                        col_6_values = []
                                        col_7_values = []


                                        #Passando por todas as linhas, uma de cada vez
                                        for row in rows_oc_details:
                                        
                                            #Percorrer todas as filas
                                            col_4_value = row.find_element(By.XPATH, "./td[4]").text
                                            col_5_value = row.find_element(By.XPATH, "./td[5]").text
                                            col_6_value = row.find_element(By.XPATH, "./td[6]").text
                                            col_6_value = col_6_value.replace(".", "") #Tirando o . que representa as casas de mil para cima para conseguir converter string para int.
                                            col_7_value = row.find_element(By.XPATH, "./td[7]").text
                                            number_ocs = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_wucOcFicha_txtOC']").text
                                            number_ocs = number_ocs[:22].strip()

                                            #Anexar os valores às respectivas listas
                                            col_4_values = int(col_4_value)
                                            col_5_values = col_5_value
                                            col_6_values = int(float(col_6_value))
                                            col_7_values = col_7_value            

                                            #Escrever os valores concatenados a uma única célula em Excel usando Pandas
                                            pe_details_oc.append({
                                                                 'OC': number_ocs,
                                                                 'SIAF.': col_4_values, 
                                                                 'Desc.': col_5_values, 
                                                                 'Qtd.': col_6_values,
                                                                 'UN': col_7_values})                

                                        #Botão para clicar na "Fase Preparatória" e uma variável para armazenar o telefone da UC
                                        element2 = WebDriverWait(browser_driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Fase Preparatória']")))
                                        pe_faseprep_button = element2.get_attribute("href")
                                        browser_driver.get(pe_faseprep_button)
                                        time.sleep(2)

                                        #Pegando o telefone da UC
                                        pe_oc_tel_uc = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_conteudo_Wuc_OC_Ficha2_txtTelUge']").text
                                        time.sleep(2)
                                        element3 = browser_driver.find_element(By.XPATH, "//*[@id='_10210']/li/a")
                                        pe_respoperson_buton = element3.get_attribute("href")
                                        browser_driver.get(pe_respoperson_buton)

                                        for tel in range(len(col_3_value)):
                                            pe_tels_oc.append({
                                                   "Tel.": pe_oc_tel_uc})

                                            tel += 1                

                                        #Encontrando a tabela dos responsáveis e suas linhas para pegar o "PREGOEIRO" e "AUTORIDADE DO PREGÃO"
                                        table_responsables = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_conteudo_grd']")
                                        rows_responsables = table_responsables.find_elements(By.XPATH, "//*[@id='ctl00_conteudo_grd']/tbody/tr")

                                        #Os valores serão declarados como vazios para depois preenchê-los
                                        value_email_preg = ""
                                        value_email_autopreg = ""

                                        #Iteração para passar por todas as linhas da segunda coluna e achar a palavra "PREGOEIRO" para ir para a coluna do lado e pegar o texto
                                        for i_row, row in enumerate(rows_responsables):
                                            col_resp = row.find_element(By.XPATH, f"//*[@id='ctl00_conteudo_grd']/tbody/tr[{i_row+1}]/td[3]")

                                            if col_resp.text == "PREGOEIRO":
                                                col_email_preg = row.find_element(By.XPATH, f"//*[@id='ctl00_conteudo_grd']/tbody/tr[{i_row+1}]/td[4]")
                                                value_email_preg = col_email_preg.text
                                                break
                                            
                                        #Iteração para passar por todas as linhas da segunda coluna e achar a palavra "AUTORIDADE DO PREGÃO" para ir para a coluna do lado e pegar o texto
                                        for i_row, row2 in enumerate(rows_responsables):
                                        
                                            col_resp2 = row2.find_element(By.XPATH, f"//*[@id='ctl00_conteudo_grd']/tbody/tr[{i_row+1}]/td[3]")

                                            if col_resp2.text == "AUTORIDADE PREGÃO":
                                            
                                                col_email_autopreg = row2.find_element(By.XPATH, f"//*[@id='ctl00_conteudo_grd']/tbody/tr[{i_row+1}]/td[4]")
                                                value_email_autopreg = col_email_autopreg.text
                                                break
                                            
                                        for email in range(len(col_3_value)):
                                            #Adicionando os detalhes ao buffer
                                            pe_emails_oc.append({
                                                "E-mail Pregoeiro": value_email_preg,
                                                "E-mail Aut. Pregão": value_email_autopreg})

                                            email += 1

                                    else:
                                    
                                        #Passando por todas as linhas, uma de cada vez
                                        for row in rows_oc_details:
                                        
                                            #Percorrer todas as filas
                                            col_4_value = row.find_element(By.XPATH, "./td[4]").text
                                            col_5_value = row.find_element(By.XPATH, "./td[5]").text
                                            col_6_value = row.find_element(By.XPATH, "./td[6]").text
                                            col_6_value = col_6_value.replace(".", "") #Tirando o . que representa as casas de mil para cima para conseguir converter string para int. (Ex: 54.000 -> 54000)
                                            col_7_value = row.find_element(By.XPATH, "./td[7]").text
                                            number_ocs = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_wucOcFicha_txtOC']").text
                                            number_ocs = number_ocs[:22].strip()

                                            #Escrever os valores concatenados a uma única célula em Excel usando Pandas
                                            pe_details_oc.append({
                                                                 'OC': number_ocs,    
                                                                 'SIAF.': int(col_4_value), 
                                                                 'Desc.': col_5_value, 
                                                                 'Qtd.': int(float(col_6_value)),
                                                                 'UN': col_7_value})

                                            print(pe_details_oc)


                                        #Botão para clicar na "Fase Preparatória" e uma variável para armazenar o telefone da UC
                                        element2 = WebDriverWait(browser_driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Fase Preparatória']")))
                                        pe_faseprep_button = element2.get_attribute("href")
                                        browser_driver.get(pe_faseprep_button)
                                        time.sleep(2)

                                        #Pegando o telefone da UC
                                        pe_oc_tel_uc = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_conteudo_Wuc_OC_Ficha2_txtTelUge']").text
                                        time.sleep(2)
                                        element3 = browser_driver.find_element(By.XPATH, "//*[@id='_10210']/li/a")
                                        pe_respoperson_buton = element3.get_attribute("href")
                                        browser_driver.get(pe_respoperson_buton)

                                        for tel in range(len(col_3_value)):
                                            pe_tels_oc.append({
                                                   "Tel.": pe_oc_tel_uc})   
                                            tel += 1  

                                        #Encontrando a tabela dos responsáveis e suas linhas para pegar o "PREGOEIRO" e "AUTORIDADE DO PREGÃO"
                                        table_responsables = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_conteudo_grd']")
                                        rows_responsables = table_responsables.find_elements(By.XPATH, "//*[@id='ctl00_conteudo_grd']/tbody/tr")

                                        #Os valores serão declarados como vazios para depois preenchê-los
                                        value_email_preg = ""
                                        value_email_autopreg = ""

                                        #Iteração para passar por todas as linhas da segunda coluna e achar a palavra "PREGOEIRO" para ir para a coluna do lado e pegar o texto
                                        for i_row, row in enumerate(rows_responsables):
                                            col_resp = row.find_element(By.XPATH, f"//*[@id='ctl00_conteudo_grd']/tbody/tr[{i_row+1}]/td[3]")

                                            if col_resp.text == "PREGOEIRO":
                                                col_email_preg = row.find_element(By.XPATH, f"//*[@id='ctl00_conteudo_grd']/tbody/tr[{i_row+1}]/td[4]")
                                                value_email_preg = col_email_preg.text
                                                break
                                            
                                        #Iteração para passar por todas as linhas da segunda coluna e achar a palavra "AUTORIDADE DO PREGÃO" para ir para a coluna do lado e pegar o texto
                                        for i_row, row2 in enumerate(rows_responsables):
                                        
                                            col_resp2 = row2.find_element(By.XPATH, f"//*[@id='ctl00_conteudo_grd']/tbody/tr[{i_row+1}]/td[3]")

                                            if col_resp2.text == "AUTORIDADE PREGÃO":
                                            
                                                col_email_autopreg = row2.find_element(By.XPATH, f"//*[@id='ctl00_conteudo_grd']/tbody/tr[{i_row+1}]/td[4]")
                                                value_email_autopreg = col_email_autopreg.text
                                                break
                                            
                                        #Adicionando os detalhes ao buffer
                                        for email in range(len(col_3_value)):
                                            #Adicionando os detalhes ao buffer
                                            pe_emails_oc.append({
                                                "E-mail Pregoeiro": value_email_preg,
                                                "E-mail Aut. Pregão": value_email_autopreg})

                                            email += 1     

                                    browser_driver.close()
                                    time.sleep(2)
                                    iterator+=1
                                    #Voltando para a aba principal        
                                    browser_driver.switch_to.window(browser_driver.window_handles[0])
                                    
                                    
                        except:
                                        table_details = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_conteudo_loteGridItens_grdLote']")

                                        #Linhas da tabela    
                                        rows_oc_details = table_details.find_elements(By.XPATH, "//*[@id='ctl00_conteudo_loteGridItens_grdLote']/tbody/tr") [1:] #Pula a primeira linha da tabela (Cabeçalho
                                        number_ocs = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_wucOcFicha_txtOC']") 
                                        col_3_value = browser_driver.find_elements(By.XPATH, "//*[@id='ctl00_conteudo_loteGridItens_grdLote']/tbody/tr/td[3]")[1:]

                                        if len(rows_oc_details)>1:
                                        
                                                        #Criando variáveis para pegar os dados das colunas da tabela
                                                        col_4_values = []
                                                        col_5_values = []
                                                        col_6_values = []
                                                        col_7_values = []


                                                        #Passando por todas as linhas, uma de cada vez
                                                        for row in rows_oc_details:
                                                        
                                                            #Percorrer todas as filas
                                                            col_4_value = row.find_element(By.XPATH, "./td[3]").text
                                                            col_5_value = row.find_element(By.XPATH, "./td[4]").text
                                                            col_6_value = row.find_element(By.XPATH, "./td[5]").text
                                                            col_6_value = col_6_value.replace(".", "") #Tirando o . que representa as casas de mil para cima para conseguir converter string para int.
                                                            col_7_value = row.find_element(By.XPATH, "./td[6]").text
                                                            number_ocs = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_wucOcFicha_txtOC']").text
                                                            number_ocs = number_ocs[:22].strip()

                                                            #Anexar os valores às respectivas listas
                                                            col_4_values = int(col_4_value)
                                                            col_5_values = col_5_value
                                                            col_6_values = int(float(col_6_value))
                                                            col_7_values = col_7_value            

                                                            #Escrever os valores concatenados a uma única célula em Excel usando Pandas
                                                            pe_details_oc.append({
                                                                                 'OC': number_ocs,
                                                                                 'SIAF.': col_4_values, 
                                                                                 'Desc.': col_5_values, 
                                                                                 'Qtd.': col_6_values,
                                                                                 'UN': col_7_values})                

                                                        #Botão para clicar na "Fase Preparatória" e uma variável para armazenar o telefone da UC
                                                        element2 = WebDriverWait(browser_driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Fase Preparatória']")))
                                                        pe_faseprep_button = element2.get_attribute("href")
                                                        browser_driver.get(pe_faseprep_button)
                                                        time.sleep(2)

                                                        #Pegando o telefone da UC
                                                        pe_oc_tel_uc = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_conteudo_Wuc_OC_Ficha2_txtTelUge']").text
                                                        time.sleep(2)
                                                        element3 = browser_driver.find_element(By.XPATH, "//*[@id='_10210']/li/a")
                                                        pe_respoperson_buton = element3.get_attribute("href")
                                                        browser_driver.get(pe_respoperson_buton)

                                                        for tel in range(len(col_3_value)):
                                                            pe_tels_oc.append({
                                                                   "Tel.": pe_oc_tel_uc})

                                                            tel += 1                

                                                        #Encontrando a tabela dos responsáveis e suas linhas para pegar o "PREGOEIRO" e "AUTORIDADE DO PREGÃO"
                                                        table_responsables = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_conteudo_grd']") 
                                                        rows_responsables = table_responsables.find_elements(By.XPATH, "//*[@id='ctl00_conteudo_grd']/tbody/tr")

                                                        #Os valores serão declarados como vazios para depois preenchê-los
                                                        value_email_preg = ""
                                                        value_email_autopreg = ""

                                                        #Iteração para passar por todas as linhas da segunda coluna e achar a palavra "PREGOEIRO" para ir para a coluna do lado e pegar o texto
                                                        for i_row, row in enumerate(rows_responsables):
                                                            col_resp = row.find_element(By.XPATH, f"//*[@id='ctl00_conteudo_grd']/tbody/tr[{i_row+1}]/td[3]")

                                                            if col_resp.text == "PREGOEIRO":
                                                                col_email_preg = row.find_element(By.XPATH, f"//*[@id='ctl00_conteudo_grd']/tbody/tr[{i_row+1}]/td[4]")
                                                                value_email_preg = col_email_preg.text
                                                                break
                                                            
                                                        #Iteração para passar por todas as linhas da segunda coluna e achar a palavra "AUTORIDADE DO PREGÃO" para ir para a coluna do lado e pegar o texto
                                                        for i_row, row2 in enumerate(rows_responsables):
                                                        
                                                            col_resp2 = row2.find_element(By.XPATH, f"//*[@id='ctl00_conteudo_grd']/tbody/tr[{i_row+1}]/td[3]")

                                                            if col_resp2.text == "AUTORIDADE PREGÃO":
                                                            
                                                                col_email_autopreg = row2.find_element(By.XPATH, f"//*[@id='ctl00_conteudo_grd']/tbody/tr[{i_row+1}]/td[4]")
                                                                value_email_autopreg = col_email_autopreg.text
                                                                break
                                                            
                                                        for email in range(len(col_3_value)):
                                                            #Adicionando os detalhes ao buffer
                                                            pe_emails_oc.append({
                                                                "E-mail Pregoeiro": value_email_preg,
                                                                "E-mail Aut. Pregão": value_email_autopreg})

                                                            email += 1

                                        else:
                                        
                                                        #Passando por todas as linhas, uma de cada vez
                                                        for row in rows_oc_details:
                                                        
                                                            #Percorrer todas as filas
                                                            col_4_value = row.find_element(By.XPATH, "./td[3]").text
                                                            col_5_value = row.find_element(By.XPATH, "./td[4]").text
                                                            col_6_value = row.find_element(By.XPATH, "./td[5]").text
                                                            col_6_value = col_6_value.replace(".", "") #Tirando o . que representa as casas de mil para cima para conseguir converter string para int. (Ex: 54.000 -> 54000)
                                                            col_7_value = row.find_element(By.XPATH, "./td[6]").text
                                                            number_ocs = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_wucOcFicha_txtOC']").text
                                                            number_ocs = number_ocs[:22].strip()

                                                            #Escrever os valores concatenados a uma única célula em Excel usando Pandas
                                                            pe_details_oc.append({
                                                                                 'OC': number_ocs,    
                                                                                 'SIAF.': int(col_4_value), 
                                                                                 'Desc.': col_5_value, 
                                                                                 'Qtd.': int(float(col_6_value)),
                                                                                 'UN': col_7_value})

                                                            print(pe_details_oc)


                                                        #Botão para clicar na "Fase Preparatória" e uma variável para armazenar o telefone da UC
                                                        element2 = WebDriverWait(browser_driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Fase Preparatória']")))
                                                        pe_faseprep_button = element2.get_attribute("href")
                                                        browser_driver.get(pe_faseprep_button)
                                                        time.sleep(2)

                                                        #Pegando o telefone da UC
                                                        pe_oc_tel_uc = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_conteudo_Wuc_OC_Ficha2_txtTelUge']").text
                                                        time.sleep(2)
                                                        element3 = browser_driver.find_element(By.XPATH, "//*[@id='_10210']/li/a")
                                                        pe_respoperson_buton = element3.get_attribute("href")
                                                        browser_driver.get(pe_respoperson_buton)

                                                        for tel in range(len(col_3_value)):
                                                            pe_tels_oc.append({
                                                                   "Tel.": pe_oc_tel_uc})   
                                                            tel += 1  

                                                        #Encontrando a tabela dos responsáveis e suas linhas para pegar o "PREGOEIRO" e "AUTORIDADE DO PREGÃO"
                                                        table_responsables = browser_driver.find_element(By.XPATH, "//*[@id='ctl00_conteudo_grd']")
                                                        rows_responsables = table_responsables.find_elements(By.XPATH, "//*[@id='ctl00_conteudo_grd']/tbody/tr")

                                                        #Os valores serão declarados como vazios para depois preenchê-los
                                                        value_email_preg = ""
                                                        value_email_autopreg = ""

                                                        #Iteração para passar por todas as linhas da segunda coluna e achar a palavra "PREGOEIRO" para ir para a coluna do lado e pegar o texto
                                                        for i_row, row in enumerate(rows_responsables):
                                                            col_resp = row.find_element(By.XPATH, f"//*[@id='ctl00_conteudo_grd']/tbody/tr[{i_row+1}]/td[3]")

                                                            if col_resp.text == "PREGOEIRO":
                                                                col_email_preg = row.find_element(By.XPATH, f"//*[@id='ctl00_conteudo_grd']/tbody/tr[{i_row+1}]/td[4]")
                                                                value_email_preg = col_email_preg.text
                                                                break
                                                            
                                                        #Iteração para passar por todas as linhas da segunda coluna e achar a palavra "AUTORIDADE DO PREGÃO" para ir para a coluna do lado e pegar o texto
                                                        for i_row, row2 in enumerate(rows_responsables):
                                                        
                                                            col_resp2 = row2.find_element(By.XPATH, f"//*[@id='ctl00_conteudo_grd']/tbody/tr[{i_row+1}]/td[3]")

                                                            if col_resp2.text == "AUTORIDADE PREGÃO":
                                                            
                                                                col_email_autopreg = row2.find_element(By.XPATH, f"//*[@id='ctl00_conteudo_grd']/tbody/tr[{i_row+1}]/td[4]")
                                                                value_email_autopreg = col_email_autopreg.text
                                                                break
                                                            
                                                        #Adicionando os detalhes ao buffer
                                                        for email in range(len(col_3_value)):
                                                            #Adicionando os detalhes ao buffer
                                                            pe_emails_oc.append({
                                                                "E-mail Pregoeiro": value_email_preg,
                                                                "E-mail Aut. Pregão": value_email_autopreg})

                                                            email += 1

                                        browser_driver.close()
                                        time.sleep(2)
                                        iterator+=1
                                        #Voltando para a aba principal        
                                        browser_driver.switch_to.window(browser_driver.window_handles[0])
                                        


    n +=1
    browser_driver.quit() 
    
    
                                                ######################################################################################
                                                ####Criando uma tabela (Excel) para visualizar os valores coletados da lista do PE#### 
                                                ######################################################################################
            
    #Utilizando pandas para criar e visualizar uma tabela formatada com os valores coletados da lista do Pregão Eletrônico
    
    #Tabela de OCs
    df_table_oc = pd.DataFrame(pe_infos_result)
    
    #Valores coletados dentro de cada OC (detalhes)
    df_oc_details_data = pd.DataFrame(pe_details_oc) #{"NÚMERO OC":number_ocs, "CÓDIGO":col_4_values, "DESCRIÇÃO":col_5_values, "QTDE.":col_6_values, "U.F.":col_7_values}
    
    #Valores de telefone
    df_oc_tels_data = pd.DataFrame(pe_tels_oc)
    
    #Valores dos e-mails
    df_oc_emails_data = pd.DataFrame(pe_emails_oc)

    #Criando um Pandas Excel Writer para usar o Openpyxl como engine e salvar os detalhes das OCs selecionadas.
    writer = pd.ExcelWriter(path_with_filename, engine='openpyxl')
    
    #Criando um Pandas Excel Writer para salvar os dados atuais da tabela de OCs
    writer2 = pd.ExcelWriter(path_with_filenameallocs, engine='openpyxl')
    
    #Unindo as duas tabelas em uma única
    df_final_data = pd.concat([df_oc_details_data, df_oc_tels_data, df_oc_emails_data], axis=1) #axis 1 é para colocar depois da coluna da primeira tabela, enquanto 0 é para colocar depois da última linha

    #Criando um arquivo .xls para utilizar os dados dos detalhes de OCs no Excel
    df_final_data.to_excel(writer, sheet_name=DateSheet, header=True, index=False)
    
    #Criando arquivo .xls para ver os dados gerais da tabela de OCs
    df_table_oc.to_excel(writer2, sheet_name=DateSheet, header=True, index=False)
    
    print(df_oc_details_data.dtypes)
    
    #Fechando o Pandas Excel Writer e fazendo o output do arquivo .xls
    writer.close()
    writer2.close() 

    #print(df_final_data)
    print('DataFrame is written to Excel File successfully!!!')
    