import os
import csv
import time
import base64
import logging
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


# ----------------------------------------------------------------------
# CONFIGURAÇÕES E UTILITÁRIOS
# ----------------------------------------------------------------------

def configurar_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler("automacao_sei.log", encoding="utf-8"),
            logging.StreamHandler()
        ]
    )


def configurar_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--window-size=1920,1080")
    options.add_experimental_option("prefs", {
        "printing.print_preview_sticky_settings.appState": '{"recentDestinations": [{"id": "Save as PDF","origin": "local"}],"selectedDestinationId": "Save as PDF","version": 2}',
        "savefile.default_directory": os.getcwd()
    })
    options.add_argument("--kiosk-printing")

    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    return driver


# ----------------------------------------------------------------------
# FUNÇÕES DE INTERAÇÃO COM O SISTEMA
# ----------------------------------------------------------------------

def realizar_login(driver, url_login, usuario, senha):
    driver.get(url_login)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "txtUsuario")))

    driver.find_element(By.NAME, "txtUsuario").send_keys(usuario)
    driver.find_element(By.NAME, "pwdSenha").send_keys(senha)
    driver.find_element(By.NAME, "btnLogin").click()

    logging.info("Login realizado com sucesso.")


def executar_busca(driver):
    """
    Navega até a tela de pesquisa e executa a busca com os critérios definidos.

    Args:
        driver: Instância do WebDriver.
    
    Returns:
        bool: True se a busca for executada, False caso contrário.
    """
    try:
        logging.info("Iniciando o processo de busca de documentos.")
        wait = WebDriverWait(driver, 5)

        # Acessa a área de busca
        searching = driver.find_element(By.XPATH, '//*[@id="infraMenu"]/li[14]/a/span')
        searching.click()
        time.sleep(1)

        # Preenche o formulário de busca
        logging.info("Preenchendo os critérios de busca.")

        # Restringe busca ao órgão específico
        sel_orgao = driver.find_element(By.XPATH, '//*[@id="divSinRestringirOrgao"]/div')
        sel_orgao.click()
        time.sleep(1)

        # Especifica os termos de pesquisa
        espec_pesq = driver.find_element(By.XPATH, '//*[@id="q"]')
        espec_pesq.send_keys('INSS não "SIC" não "Ficha" não "Nota Fiscal" não "Capacitação" não "Avaliação de Reação" não "Controle de Acesso" não "Solicitação de Cessão" não "TERMO ANUÊNCIA"')
        
        #colocar como tramitação dentro do orgão
        chktram = driver.find_element(By.XPATH, '//*[@id="divSinTramitacao"]/div')
        chktram.click()
        time.sleep(1)
        # Clica no botão de pesquisar
        b_pesq = driver.find_element(By.XPATH, '//*[@id="sbmPesquisar"]')
        b_pesq.click()
        
        # Aguarda a tabela de resultados aparecer
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo"]/table')))
        logging.info("Busca realizada com sucesso.")
        return True

    except TimeoutException:
        logging.error("Falha na busca: A página demorou muito para carregar ou um elemento não foi encontrado.")
        # return False
    except Exception as e:
        logging.error(f"Erro inesperado durante a busca: {e}")
        # return False

def extrair_dados(driver):
    """
    Extrai dados da página atual do SEI.
    """
    def remove_items(lista, item): 
        return [i for i in lista if i != item]

    tree_elements = driver.find_elements(By.XPATH, '//*[@class="pesquisaTituloEsquerda"]/a')
    list_tree = [element.text for element in tree_elements]
    trees = remove_items(list_tree, '')

    abts = driver.find_elements(By.XPATH, '//*[@class="pesquisaSnippet"]')
    list_abts = [element.text for element in abts]

    unidades = driver.find_elements(By.XPATH, '//*[@class="pesquisaMetatag"]')
    list_uni = [element.text.split(':') for element in unidades]
    info = [sub[1].strip() for sub in list_uni if len(sub) > 1]

    rows = driver.find_elements(By.XPATH, '//*[@id="conteudo"]/table/tbody/tr')
    links = []
    for i in range(1, len(rows), 3):
        try:
            a = driver.find_element(By.XPATH, f'//*[@id="conteudo"]/table/tbody/tr[{i}]/td[1]/a[1]')
            link = a.get_attribute('href')
            links.append(link)
        except Exception as e:
            logging.warning(f"Erro ao extrair link da linha {i}: {e}")

    dados = {
        "Número do Processo": trees[::2],
        "Documento": trees[1::2],
        "Resumo": list_abts,
        "Unidade": info[::3],
        "Usuário": info[1::3],
        "Data de Inclusão": info[2::3],
        "Link Completo": links
    }

    return pd.DataFrame(dados)


def criar_pastas(diretorio_base):
    if not os.path.exists(diretorio_base):
        os.makedirs(diretorio_base)
        logging.info(f"Pasta criada: {diretorio_base}")


def salvar_documento_como_pdf(driver, url, nome_arquivo_pdf, pagina, indice):
    try:
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[-1])
        driver.get(url)

        WebDriverWait(driver, 10).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        time.sleep(2)

        result = driver.execute_cdp_cmd("Page.printToPDF", {
            "printBackground": True,
            "paperWidth": 8.27,
            "paperHeight": 11.69,
            "scale": 1
        })

        with open(nome_arquivo_pdf, "wb") as f:
            f.write(base64.b64decode(result['data']))

        logging.info(f"[OK] Página {pagina} - Documento {indice}: PDF salvo em {nome_arquivo_pdf}")

    except Exception:
        logging.error(
            f"[ERRO] Página {pagina} - Documento {indice}: Falha ao salvar PDF de {url}\n{traceback.format_exc()}"
        )
    finally:
        try:
            driver.close()
        except:
            pass
        driver.switch_to.window(driver.window_handles[0])
        logging.info("Continuando com o próximo documento...")


def salvar_documentos_da_pagina(driver, diretorio_base, pagina):
    documentos = driver.find_elements(By.CSS_SELECTOR, "a[id^='lnkDocumento']")
    dados = []

    for i, doc in enumerate(documentos, 1):
        try:
            link = doc.get_attribute("href")
            nome_arquivo = os.path.join(diretorio_base, f"pagina{pagina}_doc{i}.pdf")
            salvar_documento_como_pdf(driver, link, nome_arquivo, pagina, i)
            dados.append({"pagina": pagina, "documento": i, "link": link, "arquivo": nome_arquivo})
        except Exception as e:
            logging.error(f"[ERRO] Falha ao processar documento {i} da página {pagina}: {e}")
    return dados


def navegar_paginas(driver, diretorio_base):
    pagina = 1
    dados_gerais = []

    while True:
        logging.info(f"Navegando pela página {pagina}")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "divInfraAreaTabela")))

        dados_pagina = salvar_documentos_da_pagina(driver, diretorio_base, pagina)
        dados_gerais.extend(dados_pagina)

        try:
            proximo = driver.find_element(By.LINK_TEXT, "Próximo")
            if "disabled" in proximo.get_attribute("class"):
                break
            proximo.click()
            pagina += 1
            time.sleep(2)
        except Exception:
            logging.info("Fim das páginas.")
            break

    return dados_gerais


def salvar_csv(dados, caminho_arquivo):
    with open(caminho_arquivo, "w", newline="", encoding="utf-8") as csvfile:
        campos = ["pagina", "documento", "link", "arquivo"]
        writer = csv.DictWriter(csvfile, fieldnames=campos)
        writer.writeheader()
        writer.writerows(dados)
    logging.info(f"CSV salvo em {caminho_arquivo}")


# ----------------------------------------------------------------------
# FUNÇÃO PRINCIPAL
# ----------------------------------------------------------------------

def main(url_login, usuario, senha, pasta_destino, caminho_csv):
    configurar_logging()
    driver = configurar_driver()

    try:
        realizar_login(driver, url_login, usuario, senha)
        executar_busca(driver)
        criar_pastas(pasta_destino)

        dados = navegar_paginas(driver, pasta_destino)
        salvar_csv(dados, caminho_csv)

    except Exception as e:
        logging.error(f"Erro inesperado na execução: {e}")
    finally:
        driver.quit()


# Exemplo de chamada
if __name__ == "__main__":
    main(
        url_login="https://colaboragov.sei.gov.br/sip/modulos/MF/login_especial/login_especial.php?sigla_orgao_sistema=MGI&sigla_sistema=SEI",
        usuario="",
        senha="",
        pasta_destino="documentos_INSS",
        caminho_csv="documentos_INSS.csv"
    )