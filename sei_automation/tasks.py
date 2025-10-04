import os
import time
import logging
from celery import shared_task
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from .models import BuscaSEI, DocumentoSEI
import pandas as pd
from datetime import datetime


logger = logging.getLogger(__name__)


def configurar_driver():
    """Configura o Chrome driver em modo headless"""
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    driver = webdriver.Chrome(options=chrome_options)
    driver.implicitly_wait(5)
    return driver


def realizar_login(driver, url, usuario, senha, orgao):
    """Realiza login no SEI"""
    try:
        driver.get(url)
        wait = WebDriverWait(driver, 10)
        
        wait.until(EC.presence_of_element_located((By.ID, 'txtUsuario'))).send_keys(usuario)
        wait.until(EC.presence_of_element_located((By.ID, 'pwdSenha'))).send_keys(senha)
        wait.until(EC.presence_of_element_located((By.ID, 'selOrgao'))).send_keys(orgao)
        
        submit_button = driver.find_element(By.ID, 'Acessar')
        submit_button.click()
        time.sleep(3)
        
        logger.info(f"Login realizado com sucesso para {usuario}")
        return True
    except Exception as e:
        logger.error(f"Erro no login: {e}")
        return False


def executar_busca(driver, termos, data_inicio=None, data_fim=None):
    """Executa a busca no SEI"""
    try:
        # Acessa a área de busca
        searching = driver.find_element(By.XPATH, '//*[@id="infraMenu"]/li[14]/a/span')
        searching.click()
        time.sleep(2)
        
        # Restringe busca ao órgão
        sel_orgao = driver.find_element(By.XPATH, '//*[@id="divSinRestringirOrgao"]/div')
        sel_orgao.click()
        time.sleep(1)
        
        # Preenche os termos de pesquisa
        espec_pesq = driver.find_element(By.XPATH, '//*[@id="txtDescricaoPesquisa"]')
        espec_pesq.send_keys(termos)
        time.sleep(1)
        
        # Adiciona datas se fornecidas
        if data_inicio:
            data_inicio_input = driver.find_element(By.XPATH, '//*[@id="txtDataInicio"]')
            data_inicio_input.send_keys(data_inicio.strftime('%d/%m/%Y'))
        
        if data_fim:
            data_fim_input = driver.find_element(By.XPATH, '//*[@id="txtDataFim"]')
            data_fim_input.send_keys(data_fim.strftime('%d/%m/%Y'))
        
        # Realiza a pesquisa
        b_pesq = driver.find_element(By.XPATH, '//*[@id="sbmPesquisar"]')
        b_pesq.click()
        time.sleep(3)
        
        logger.info("Busca executada com sucesso")
        return True
    except Exception as e:
        logger.error(f"Erro ao executar busca: {e}")
        return False


def extrair_dados_pagina(driver):
    """Extrai dados de uma página de resultados"""
    documentos = []
    
    try:
        # Extrai elementos da página
        tree_elements = driver.find_elements(By.XPATH, '//*[@class="pesquisaTituloEsquerda"]/a')
        trees = [element.text for element in tree_elements if element.text]
        
        abts = driver.find_elements(By.XPATH, '//*[@class="pesquisaSnippet"]')
        resumos = [element.text for element in abts]
        
        unidades = driver.find_elements(By.XPATH, '//*[@class="pesquisaMetatag"]')
        metadados = []
        for element in unidades:
            parts = element.text.split(':')
            if len(parts) > 1:
                metadados.append(parts[1].strip())
        
        # Extrai links
        rows = driver.find_elements(By.XPATH, '//*[@id="conteudo"]/table/tbody/tr')
        links = []
        for i in range(1, len(rows), 3):
            try:
                a = driver.find_element(By.XPATH, f'//*[@id="conteudo"]/table/tbody/tr[{i}]/td[1]/a[1]')
                link = a.get_attribute('href')
                links.append(link)
            except:
                links.append(None)
        
        # Organiza os dados
        for i in range(0, len(trees), 2):
            if i+1 < len(trees):
                doc = {
                    'numero_processo': trees[i] if i < len(trees) else '',
                    'nome_documento': trees[i+1] if i+1 < len(trees) else '',
                    'resumo': resumos[i//2] if i//2 < len(resumos) else '',
                    'unidade': metadados[i*3//2] if i*3//2 < len(metadados) else '',
                    'usuario_inclusao': metadados[i*3//2 + 1] if i*3//2 + 1 < len(metadados) else '',
                    'data_inclusao': metadados[i*3//2 + 2] if i*3//2 + 2 < len(metadados) else '',
                    'link': links[i//2] if i//2 < len(links) else '',
                }
                documentos.append(doc)
        
        return documentos
    except Exception as e:
        logger.error(f"Erro ao extrair dados: {e}")
        return []


def navegar_paginas(driver, busca_obj):
    """Navega por todas as páginas e salva os documentos"""
    todos_documentos = []
    pagina = 1
    
    while True:
        logger.info(f"Processando página {pagina}")
        
        # Extrai dados da página atual
        documentos = extrair_dados_pagina(driver)
        
        # Salva documentos no banco
        for doc_data in documentos:
            try:
                # Tenta parsear a data
                data_inclusao = None
                if doc_data.get('data_inclusao'):
                    try:
                        data_inclusao = datetime.strptime(doc_data['data_inclusao'], '%d/%m/%Y %H:%M')
                    except:
                        pass
                
                documento = DocumentoSEI.objects.create(
                    busca=busca_obj,
                    numero_processo=doc_data.get('numero_processo', ''),
                    nome_documento=doc_data.get('nome_documento', ''),
                    resumo=doc_data.get('resumo', ''),
                    unidade=doc_data.get('unidade', ''),
                    usuario_inclusao=doc_data.get('usuario_inclusao', ''),
                    data_inclusao=data_inclusao,
                    link_original=doc_data.get('link', ''),
                )
                todos_documentos.append(documento)
            except Exception as e:
                logger.error(f"Erro ao salvar documento: {e}")
        
        # Verifica se há próxima página
        try:
            next_page = driver.find_element(By.XPATH, '//*[@id="conteudo"]/div[2]/div[3]/a')
            if not next_page.get_attribute('href'):
                break
            next_page.click()
            time.sleep(3)
            pagina += 1
        except:
            break
    
    return todos_documentos


@shared_task
def executar_busca_sei(busca_id, senha):
    """Tarefa Celery para executar a busca SEI"""
    try:
        busca = BuscaSEI.objects.get(id=busca_id)
        busca.status = 'em_andamento'
        busca.save()
        
        logger.info(f"Iniciando busca {busca_id}")
        
        # Configura o driver
        driver = configurar_driver()
        
        try:
            # Realiza login
            if not realizar_login(driver, busca.url_sei, busca.usuario_sei, senha, busca.orgao):
                raise Exception("Falha no login")
            
            # Executa a busca
            if not executar_busca(driver, busca.termos_pesquisa, busca.data_inicio, busca.data_fim):
                raise Exception("Falha ao executar busca")
            
            # Navega pelas páginas e coleta documentos
            documentos = navegar_paginas(driver, busca)
            
            # Atualiza status
            busca.status = 'concluida'
            busca.save()
            
            logger.info(f"Busca {busca_id} concluída. {len(documentos)} documentos encontrados.")
            
        finally:
            driver.quit()
        
        return f"Busca concluída: {len(documentos)} documentos"
        
    except Exception as e:
        logger.error(f"Erro na busca {busca_id}: {e}")
        busca = BuscaSEI.objects.get(id=busca_id)
        busca.status = 'erro'
        busca.mensagem_erro = str(e)
        busca.save()
        raise
