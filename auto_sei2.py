import os
import time
import csv
import base64
import logging
import traceback
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ------------------------------------
# CONFIGURAÇÕES INICIAIS
# ------------------------------------
def configurar_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler("execucao.log", mode='w', encoding='utf-8'),
            logging.StreamHandler()
        ]
    )

# ------------------------------------
# ACESSO E LOGIN
# ------------------------------------
def realizar_login(driver, url_login, usuario, senha):
    driver.get(url_login)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "txtUsuario")))
    driver.find_element(By.NAME, "txtUsuario").send_keys(usuario)
    driver.find_element(By.NAME, "pwdSenha").send_keys(senha)
    driver.find_element(By.NAME, "sbmLogin" ).click()
    logging.info("Login realizado com sucesso.")

# ------------------------------------
# BUSCA
# ------------------------------------
def executar_busca(driver, url_busca):
    driver.get(url_busca)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "resultado")))
    logging.info("Busca executada e resultados carregados.")

# ------------------------------------
# CRIAÇÃO DE PASTAS
# ------------------------------------
def criar_pastas(diretorio):
    os.makedirs(diretorio, exist_ok=True)
    logging.info(f"Pasta criada: {diretorio}")

# ------------------------------------
# EXTRAÇÃO DE DADOS DA PÁGINA
# ------------------------------------
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

 

# ------------------------------------
# SALVAMENTO COMO PDF
# ------------------------------------
def salvar_documento_como_pdf(driver, url, nome_arquivo_pdf, pagina, indice):
    try:
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[-1])

        driver.get(url)
        try:
            WebDriverWait(driver, 15).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            time.sleep(2)
        except TimeoutException:
            logging.warning(f"[AVISO] Timeout ao carregar {url} na Página {pagina} - Documento {indice}")

        result = driver.execute_cdp_cmd("Page.printToPDF", {
            "printBackground": True,
            "paperWidth": 8.27,
            "paperHeight": 11.69,
            "scale": 1
        })

        with open(nome_arquivo_pdf, "wb") as f:
            f.write(base64.b64decode(result['data']))
        logging.info(f"[OK] Página {pagina} - Documento {indice}: PDF salvo em {nome_arquivo_pdf}")

        driver.close()
        driver.switch_to.window(driver.window_handles[0])

    except Exception:
        logging.error(f"[ERRO] Página {pagina} - Documento {indice}: Falha ao salvar PDF de {url}\n{traceback.format_exc()}")
        try:
            driver.close()
        except Exception as fe:
            logging.warning(f"Falha ao fechar aba: {fe}")
        finally:
            if len(driver.window_handles) > 0:
                driver.switch_to.window(driver.window_handles[0])
            logging.info("Continuando com o próximo documento...")

# ------------------------------------
# NAVEGAÇÃO ENTRE PÁGINAS E DOWNLOADS
# ------------------------------------
def navegar_paginas(driver, pasta_destino, max_paginas=10):
    todos_dados = []

    for pagina in range(1, max_paginas + 1):
        logging.info(f"Processando página {pagina}...")

        dados = extrair_dados(driver)
        for idx, (titulo, link) in enumerate(dados):
            nome_base = f"pagina{pagina:02d}_doc{idx + 1:03d}.pdf"
            caminho_pdf = os.path.join(pasta_destino, nome_base)
            salvar_documento_como_pdf(driver, link, caminho_pdf, pagina, idx + 1)
            todos_dados.append({"pagina": pagina, "titulo": titulo, "link": link, "arquivo": nome_base})

        # Tentar ir para a próxima página
        try:
            proximo = driver.find_element(By.LINK_TEXT, str(pagina + 1))
            proximo.click()
            WebDriverWait(driver, 10).until(EC.staleness_of(proximo))
        except Exception:
            logging.info("Última página alcançada ou botão de próxima página não encontrado.")
            break

    return todos_dados

# ------------------------------------
# SALVAR CSV DOS DOCUMENTOS
# ------------------------------------
def salvar_csv(dados, caminho_csv):
    with open(caminho_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["pagina", "titulo", "link", "arquivo"])
        writer.writeheader()
        writer.writerows(dados)
    logging.info(f"CSV salvo em: {caminho_csv}")

# ------------------------------------
# MAIN
# ------------------------------------
def main():
    configurar_logging()

    # CONFIGURAÇÕES DO USUÁRIO
    url_login = "https://colaboragov.sei.gov.br/sip/modulos/MF/login_especial/login_especial.php?sigla_orgao_sistema=MGI&sigla_sistema=SEI"
    usuario = "seu_usuario"
    senha = "sua_senha"
    pasta_destino = "documentos_pdf"
    caminho_csv = "documentos_extraidos.csv"

    criar_pastas(pasta_destino)

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.set_capability("goog:loggingPrefs", {"performance": "ALL"})

    driver = webdriver.Chrome(options=chrome_options)

    try:
        realizar_login(driver, url_login, usuario, senha)
        executar_busca(driver, url_busca)
        dados = navegar_paginas(driver, pasta_destino, max_paginas=24)
        salvar_csv(dados, caminho_csv)
    finally:
        driver.quit()
        logging.info("Execução finalizada.")

if __name__ == "__main__":
    main()