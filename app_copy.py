import streamlit as st
# import os
# from datetime import datetime
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import os
# from bs4 import BeautifulSoup
# import re
# from collections import OrderedDict
from io import BytesIO
from selenium.common.exceptions import NoSuchElementException
# import getpass
import pyshorteners

# Função para realizar login
def realizar_login(url, login1, password1, orgao1):
    """
    Função para realizar login no sistema SEI.

    Args:
        url (str): URL do sistema SEI.
        login1 (str): Nome de usuário para login.
        password1 (str): Senha para login.
        orgao1 (str): Nome do órgão para acesso.

    Returns:
        webdriver: Instância do WebDriver com o usuário autenticado.
    """
    try:
        # Inicializa o navegador
        driver = webdriver.Chrome()
        driver.implicitly_wait(0.5)
        driver.get(url)

        # Localiza os elementos de login
        login = driver.find_element(By.XPATH, '//*[@id="txtUsuario"]')
        password = driver.find_element(By.XPATH, '//*[@id="pwdSenha"]')
        orgao = driver.find_element(By.XPATH, '//*[@id="selOrgao"]')
        submit_button = driver.find_element(By.XPATH, '//*[@id="Acessar"]')

        # Preenche as credenciais
        login.send_keys(login1)
        password.send_keys(password1)
        orgao.send_keys(orgao1)

        # Realiza o login
        submit_button.click()
        time.sleep(0.5)  # Aguarda carregamento da página após o login

        print("Login realizado com sucesso!")
        
        # except Exception as e:
        # print(f"Erro ao realizar o login: {e}")
        # return None
        
        # Acessa a área de busca
        searching = driver.find_element(By.XPATH, '//*[@id="infraMenu"]/li[14]/a/span')
        searching.click()
        time.sleep(3)

        # Restringe busca ao órgão específico
        sel_orgao = driver.find_element(By.XPATH, '//*[@id="divSinRestringirOrgao"]/div')
        sel_orgao.click()
        time.sleep(4)

        # Especifica os termos de pesquisa
        espec_pesq = driver.find_element(By.XPATH, '//*[@id="txtDescricaoPesquisa"]')
        espec_pesq.send_keys('"Projeto Lei" ou "PL" ou "RIC" ou "Projeto de Lei" ou "Requisição de Informação" ou "PLP" ou "PLN"')
        time.sleep(3)
        # colocar datas
        time.sleep(3)
        data_inicio = driver.find_element('xpath', '//*[@id="txtDataInicio"]')
        data_inicio.send_keys("01/09/2024")

        # Realiza a pesquisa
        time.sleep(3)
        b_pesq = driver.find_element(By.XPATH, '//*[@id="sbmPesquisar"]')
        b_pesq.click()
        time.sleep(5)
        

        print("Busca realizada com sucesso.\nRestringindo em PL e dentro do MGI.\n\nOs Externo entram como MGI.")
        return driver
    except Exception as e:
        print(f"Erro durante a busca: {e}")
        return None
    
        
def extrair_dados(driver):     
    """
    Função para realizar web scraping no SEI e retornar os dados em um DataFrame.

    Args:
        driver (webdriver): Instância do Selenium WebDriver.

    Returns:
        pd.DataFrame: DataFrame contendo os dados extraídos.
    """
    def remove_items(lista, item): 
        """Remove todos os itens iguais a `item` de uma lista."""
        return [i for i in lista if i != item]

    # Extraindo os elementos da pesquisa
    tree_elements = driver.find_elements("xpath", '//*[@class="pesquisaTituloEsquerda"]/a')
    list_tree = [element.text for element in tree_elements]
    trees = remove_items(list_tree, '')  # Remover elementos vazios

    abts = driver.find_elements("xpath", '//*[@class="pesquisaSnippet"]')
    list_abts = [element.text for element in abts]

    unidades = driver.find_elements("xpath", '//*[@class="pesquisaMetatag"]')
    list_uni = [element.text.split(':') for element in unidades]
    info = [sublist_uni[1] for sublist_uni in list_uni if len(sublist_uni) > 1]  # Removendo listas vazias

    rows = driver.find_elements("xpath", '//*[@id="conteudo"]/table/tbody/tr')
    links = []
    for i in range(1, len(rows), 3):
        try:
            a = driver.find_element("xpath", f'//*[@id="conteudo"]/table/tbody/tr[{i}]/td[1]/a[1]')
            time.sleep(0.5)
            link = a.get_attribute('href')
            time.sleep(0.5)
            links.append(link)
        except Exception as e:
            print(f"Erro ao processar a linha {i}: {e}")

    shortener = pyshorteners.Shortener()
    links_curtos = []
    for link in links:
        try:
            link_curto = shortener.tinyurl.short(link)
            links_curtos.append(link_curto)
            time.sleep(3)
        except Exception as e:
            print(f"Erro ao encurtar o link {link}: {e}")

    dados = {
        "Número do Processo": trees[::2],
        "Documento": trees[1::2],
        "Resumo": list_abts,
        "Unidade": info[::3],
        "Usuário": info[1::3],
        "Data de Inclusão": info[2::3],
        "Links": links_curtos
    }

    df = pd.DataFrame(dados)
    return df

def navegar_paginas(driver):
    """
    Loop para navegar por todas as páginas até que não haja mais um botão 'Próxima'.
    Retorna um DataFrame consolidado com os dados de todas as páginas.
    """
    dados_consolidados = pd.DataFrame()  # DataFrame vazio para acumular os dados

    while True:
        try:
            # Extrair os dados da página atual e adicionar ao DataFrame consolidado
            df_pagina = extrair_dados(driver)
            dados_consolidados = pd.concat([dados_consolidados, df_pagina], ignore_index=True)

            # Procurar o botão "Próxima"
            next_page = driver.find_element("xpath", '//*[@id="conteudo"]/div[2]/div[3]/a')

            # Verificar se o botão "Próxima" tem o atributo 'href'
            proxima_href = next_page.get_attribute('href')
            time.sleep(3)
            if not proxima_href:
                print("Não há mais páginas. Encerrando navegação.")
                break  # Sai do loop se não houver link para a próxima página

            # Clicar no botão "Próxima"
            next_page.click()
            time.sleep(10)  # Aguarda o carregamento da próxima página

        except NoSuchElementException:
            print("Botão 'Próxima' não encontrado. Encerrando navegação.")
            break  # Sai do loop se o botão "Próxima" não existir

        except Exception as e:
            print(f"Erro inesperado: {e}")
            break  # Sai do loop em caso de erro inesperado

    # Fechar o navegador após o término
    # driver.close()
    # driver.quit()

    return dados_consolidados

def gerar_excel(dados_consolidados):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        dados_consolidados.to_excel(writer, index=False, sheet_name="Dados")
    processed_data = output.getvalue()
    return processed_data

# Interface Streamlit
st.title("Automação SEI")

st.header("Configurações de Login")
url = st.text_input("URL do SEI", value="https://sei.economia.gov.br/")
login1 = st.text_input("Login")
password1 = st.text_input("Senha", type="password")
orgao1 = st.text_input("Órgão")

# TO DO: fazer o __main__

if st.button("Executar Busca"):
    with st.spinner("Realizando login e extração..."):
        driver = realizar_login(url, login1, password1, orgao1)
        # buscar_arquivos(driver)
        if driver:
            df = navegar_paginas(driver)
            st.dataframe(df)
            
            excel_data = gerar_excel(df)
            st.download_button(
                label="Baixar dados em Excel",
                data=excel_data,
                file_name="dados_sei.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            
