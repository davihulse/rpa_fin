# -*- coding: utf-8 -*-
"""
Created on Sat Jan 31 11:52:43 2026

@author: davi.hulse
"""

from time import sleep
import re
import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
#from datetime import datetime
import os
#import ctypes
#import win32com.client as win32
#import gspread
#import csv

def criar_driver(pasta_downloads):
    options = Options()
    #options.add_argument("--headless")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-gcm-registration")
    #options.add_argument("--user-data-dir=" + os.path.abspath("chrome_profile"))  # perfil persistente
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_argument("--log-level=3")  # reduz nível de log do Chrome
    
    options.add_experimental_option("prefs", {
        "download.default_directory": pasta_downloads,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    
    #options.add_experimental_option("excludeSwitches", ["enable-logging"])
    service = Service(log_path="NUL")
    driver = Chrome(service=service, options=options)
    return driver

def login_microsoft(driver):
    davpass = open(os.path.join(os.path.dirname(os.getcwd()), '.cpass'), 'r').read()    

    print("Realizando Login...")

    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, "i0116"))
    ).send_keys('davi.hulse@sc.senai.br' + Keys.ENTER)

    # Etapa 2 – Senha com loop de tentativa
    for tentativa in range(5):
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "i0118"))
            )
            password_input = driver.find_element(By.ID, "i0118")
            password_input.clear()
            password_input.send_keys(davpass)
            botao_entrar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "idSIButton9"))
            )
            botao_entrar.click()
            break
        except Exception:
            print(f"⏳ Tentativa {tentativa+1}/5 falhou ao localizar campo de senha. Retentando...")
            sleep(1)
    else:
        print("❌ Não foi possível enviar a senha após múltiplas tentativas.")

    print("✅ Login realizado com sucesso.")

def baixar_xls():
    if not hasattr(baixar_xls, "limpeza_executada"):
        # Limpar a pasta de downloads antes de iniciar
        for f in os.listdir(pasta_downloads):
            caminho_arquivo = os.path.join(pasta_downloads, f)
            try:
                os.remove(caminho_arquivo)
            except Exception as e:
                print(f"⚠️ Erro ao remover {f}: {e}")
        baixar_xls.limpeza_executada = True  # marca que já limpou
    
    WebDriverWait(driver, 100).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    sleep(2)
        
    for tentativa in range(3):
        try:
            # botão seta
            botao_seta = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@aria-label, 'Opções de plano')]"))
            )
            botao_seta.click()
    
            # botão "Exportar para Excel"
            botao_exportar = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Exportar plano para o Excel']"))
            )
            botao_exportar.click()
    
            print("Baixando arquivo XLS...")
            break  # sucesso: sai do loop
    
        except Exception as e:
            print(f"⚠️ Tentativa {tentativa + 1}/3 falhou: {e}")
            sleep(2)
    else:
        print("❌ Não foi possível clicar no botão de exportação após 3 tentativas.")
        return
    
    sleep(1)
    
    print("Aguardando download...")

    arquivos_antes = set(os.listdir(pasta_downloads))

    inicio = time.time()
    timeout_inicio = 60

    while time.time() - inicio < timeout_inicio:
        arquivos_depois = set(os.listdir(pasta_downloads))
        novos = arquivos_depois - arquivos_antes
        if novos:
            break
        time.sleep(1)
    else:
        raise TimeoutError("Download não iniciou dentro do tempo esperado.")
    
    inicio = time.time()
    timeout = 600
    
    while time.time() - inicio < timeout:
        arquivos_em_download = [
            f for f in os.listdir(pasta_downloads)
            if f.endswith(".crdownload")
        ]
    
        if not arquivos_em_download:
            print("Download concluído.")
            break

        time.sleep(1)
    else:
        raise TimeoutError("Download não terminou dentro do tempo esperado.")

def exportar_planners(driver):
    for idx, url in enumerate(planners_urls, 1):
        print(f"🔗 Acessando Planner {idx}...")
        for tentativa in range(3):
            try:
                driver.get(url)
                WebDriverWait(driver, 60).until(
                    lambda d: d.execute_script("return document.readyState") == "complete"
                )
                sleep(1)
                baixar_xls()
                break
            except Exception as e:
                print(f"⚠️ Erro ao acessar o Planner {idx} na tentativa {tentativa+1}/3: {e}")
                sleep(2)
        else:
            print(f"❌ Falha ao exportar Planner {idx} após 3 tentativas.")

def consolidar_planilhas(pasta_downloads):
    nomes_esperados = [
        "Aquisições ISI Manufatura.xlsx",
        "Aquisições ISI Laser.xlsx",
        "Aquisições ISI Embarcados 2.xlsx"
    ]

    dfs = []
    for nome_arquivo in nomes_esperados:
        caminho = os.path.join(pasta_downloads, nome_arquivo)
        if not os.path.exists(caminho):
            print(f"⚠️ Arquivo não encontrado: {nome_arquivo}")
            continue
        try:
            df = pd.read_excel(caminho)
            #df["__arquivo_origem__"] = nome_arquivo
            dfs.append(df)
            print(f"📥 Planilha carregada: {nome_arquivo} ({df.shape[0]} linhas, {df.shape[1]} colunas)")
        except Exception as e:
            print(f"❌ Erro ao ler {nome_arquivo}: {e}")

    if not dfs:
        print("⚠️ Nenhuma planilha válida foi carregada.")
        return pd.DataFrame()

    df_final = pd.concat(dfs, ignore_index=True, sort=False)
    print(f"✅ Consolidação concluída: {df_final.shape[0]} linhas totais, {df_final.shape[1]} colunas distintas")
    return df_final

#######################
    
planners_urls = [
    "https://planner.cloud.microsoft/webui/plan/QXrbRoU7UEGdjE_bhw-QY2QAFn9X/view/board?tid=2cf7d4d5-bd1b-4956-acf8-2995399b2168",
    "https://planner.cloud.microsoft/webui/plan/vIOkh-y5EEuwwAlkWsRQRmQAER1C/view/board?tid=2cf7d4d5-bd1b-4956-acf8-2995399b2168",
    "https://planner.cloud.microsoft/webui/plan/By2-rKiP6EWT0TfgDNLG12QAGgc8/view/board?tid=2cf7d4d5-bd1b-4956-acf8-2995399b2168"
]

pasta_downloads = r"C:\RPA\rpa_fin\Downloads"
driver = criar_driver(pasta_downloads)
#driver.get("https://planner.cloud.microsoft/webui/plan/QXrbRoU7UEGdjE_bhw-QY2QAFn9X/view/board?tid=2cf7d4d5-bd1b-4956-acf8-2995399b2168")
#login_microsoft(driver)
#exportar_planners(driver)

df = consolidar_planilhas(pasta_downloads)

df["Nome do Bucket"] = df["Nome do Bucket"].astype(str).str.strip()

df = df[~df["Nome do Bucket"].isin(["Brementur", "Pc de Viagem", "PC de Viagem"])]


#%%

def extrair_numero_tarefa(texto):
    if pd.isna(texto):
        return None

    # 1) Número de tarefa com 5, 6 ou 7 dígitos
    match_num = re.search(r"(?:Tarefa|Chamado)[^0-9]{0,10}(\d{5,7})", texto, flags=re.IGNORECASE)
    if match_num:
        return match_num.group(1).zfill(6)

    # 2) CT apenas nos formatos CT082/25 ou CT 082/25
    match_ct = re.search(r"(?:Tarefa|Chamado)[^C]{0,15}(CT ?\d{3}/\d{2})", texto, flags=re.IGNORECASE)
    if match_ct:
        return match_ct.group(1).replace(" ", "")

    return None


# Aplicar nova função ao DataFrame
df["Numero Tarefa"] = df["Nome da tarefa"].apply(extrair_numero_tarefa)


def extrair_fin(texto):
    if pd.isna(texto):
        return None
    match = re.search(r"FIN[:.\s\-]*?(\d{4,6}/\d{2})", texto, flags=re.IGNORECASE)
    if match:
        return f"FIN.{match.group(1)}"
    return None

df["FIN"] = df["Itens da lista de verificação"].apply(extrair_fin)


#### Bloco apenas para organizar colunas, pode ser removido depois:
colunas = df.columns.tolist()
idx = colunas.index("Nome da tarefa")
colunas.remove("Numero Tarefa")
colunas.insert(idx + 1, "Numero Tarefa")
df = df[colunas]


#df.to_excel("df.xlsx", index=False)

def login_sesuite():
    
    davpass = open(os.path.join(os.path.dirname(os.getcwd()), '.cpass'), 'r').read()    
    
    print("Acessando SE Suite...")
    
    driver.get(r'https://sesuite.fiesc.com.br/softexpert/workspace?page=home')
    
    #driver.maximize_window()
    
    username_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH,'//*[@id="userNameInput"]'))
    )
    username_input.send_keys('davi.hulse@sc.senai.br')
    
    password_input = driver.find_element(By.XPATH, '//*[@id="passwordInput"]')
    password_input.send_keys(davpass + Keys.ENTER)
    
    print("Login no SE Suite realizado.")


def extrai_fin(numfin):
    sleep(1)
    
    driver.get(r'https://sesuite.fiesc.com.br/softexpert/workspace?page=home')
    
    janela_principal = driver.window_handles[0]
 
    xpaths_input = [
        '//*[@id="st-container"]/div/div/div/div[1]/ul[3]/div/div/div[1]/input',
        '//*[@id="st-container"]/div/div[1]/div/div[1]/ul[3]/div/div/div[1]/input',
        '//*[@id="st-container"]/div/div/div/div[1]/ul[3]/div/div/div[2]/input'
    ]
    
    inserir_fin = None

    for xpath_input in xpaths_input:
        try:
            inserir_fin = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, xpath_input))
            )
            break  # encontrou, sai do loop
        except:
            continue
    
    if not inserir_fin:
        print(f"❌ Não foi possível localizar o campo de busca do {numfin}. Pulando.")
        return None
    
    # for xpath_input in xpaths_input:
    #     try:
    #         inserir_compra = WebDriverWait(driver, 3).until(
    #             EC.element_to_be_clickable((By.XPATH, xpath_input))
    #         )
    #         break
    #     except:
    #         continue
    
    inserir_fin.clear()
    sleep(1)
    inserir_fin.send_keys(str(numfin))
    sleep(1)
    inserir_fin.send_keys(Keys.ENTER)
    
    print("Aguardando SE Suite...")
        
    try:
        primeiro_item = WebDriverWait(driver, 200).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="st-container"]/div/div/div/div[4]/div/div[2]/div/div/div[2]/div/div[2]/div[1]/span'))
        )
        print("FIN localizado. Extraindo dados...")
    except TimeoutException:
        print("❌ Nenhum FIN encontrado. Pulando.")
        return None
        
    for tentativa in range(5):
        handles_antes = set(driver.window_handles)
        try:
            primeiro_item.click()
            WebDriverWait(driver, 10).until(lambda d: len(set(d.window_handles) - handles_antes) > 0)
            nova_janela = list(set(driver.window_handles) - handles_antes)[0]
            driver.switch_to.window(nova_janela)
            break
        except:
            print("❌ Erro ao abrir nova janela para o chamado. Tentando novamente...")
            sleep(2)
    else:
        print("❌ Todas as tentativas falharam. Pulando chamado.")
        return None
    
    dados_dos_chamados = {}
    
    # titulo_element = WebDriverWait(driver, 10).until(
    # EC.presence_of_element_located((By.XPATH, '//*[@id="headerTitle"]'))
    # )
    # titulo_completo = titulo_element.text.strip()
    # titulo_limpo = titulo_completo.split(" - ", 1)[1] if " - " in titulo_completo else ""
    
    try:
        titulo_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="headerTitle"]'))
        )
        titulo_completo = titulo_element.text.strip()
        #titulo_limpo = titulo_completo.split(" - ", 1)[1] if " - " in titulo_completo else ""
    except TimeoutException:
        print("❌ Timeout ao tentar localizar o título do chamado. Pulando.")
        driver.close()
        driver.switch_to.window(janela_principal)
        return None
    
    # Status do chamado
    status_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="statusTextSpan"]'))
    )
    status_texto = status_element.text.strip()
    
    # frames = driver.find_elements(By.TAG_NAME, "iframe")
    # print(len(frames))
    # for f in frames:
    #     print(f.get_attribute("name"), f.get_attribute("id"))
        
    ## Troca para o frame
    try:
        WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.NAME, "ribbonFrame"))
        )

    except TimeoutException:
        print(f"❌ Timeout ao carregar frame 'ribbonFrame' no {numfin}. Pulando FIN.")
        driver.close()
        driver.switch_to.window(janela_principal)
        return None

    #############Debug para achar o frame correto
    # iframes = driver.find_elements(By.TAG_NAME, "iframe")
    # print(f"🔎 {len(iframes)} iframe(s) encontrados:")
    # for i, iframe in enumerate(iframes):
    #     nome = iframe.get_attribute("name")
    #     id_ = iframe.get_attribute("id")
    #     src = iframe.get_attribute("src")
    #     print(f"[{i}] name={nome}, id={id_}, src={src}")


    try:
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, 'iframe[name^="frame_form_"]'))
        )
    except TimeoutException:
        print("❌ Frame não carregou. Pulando chamado.")
        return None


        
    #Campos a extrair
    campos = [
        ("Número do FIN", '//*[@id="field_8a3449076f9f6db3016fe74952d0181b"]'),
        ("CNPJ Fornecedor", '//*[@id="field_8a3449076f9f6db3016fe747b2fe17cf"]'),
        ("Número do documento", '//*[@id="field_8a3449077918207d017980762ad719ba"]'),
        ("Ordem de compra", '//*[@id="field_8a3449076f9f6db301701adda4b73de1"]'),
        ("Nr. do documento (CAP)", '//*[@id="field_8a344907739c40c101743078077162ed"]'),
        ("Número AP", '//*[@id="field_8a3449076f9f6db3016fe735db491529"]'),
        ("Data Agendada para Pagamento", '//*[@id="field_8a3449076f9f6db301701a9b6d9c3533"]'),
        ("Valor Bruto a Pagar (R$)", '//*[@id="field_8a34490770c96a380170e426eed86216"]'),
        ("Valor a deduzir (R$)", '//*[@id="field_8a34490770c96a380170e427cc6e6266"]'),
        ("Valor Líquido a Pagar (R$)", '//*[@id="field_8a34490770c96a380170e4277bfb623e"]')
    ]
               
    for nome, xpath in campos:
        element = WebDriverWait(driver, 100).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        dados_dos_chamados[nome] = element.get_attribute("value")
    
    dados_dos_chamados["Descrição"] = titulo_completo
                 
    for janela in driver.window_handles:
        if janela != janela_principal:
            driver.switch_to.window(janela)
            driver.close()

    driver.switch_to.window(janela_principal)
    
    dados_dos_chamados["Status"] = status_texto
    
    print("Descrição: ", dados_dos_chamados["Descrição"])
    print("Número do FIN: ", dados_dos_chamados["Número do FIN"])
    print("Valor Líquido: ", dados_dos_chamados["Valor Líquido"])
    print("Status: ", dados_dos_chamados["Status"])

    print("Dados do ", numfin, " extraídos.")
    
    return dados_dos_chamados


login_sesuite()

lista_fin = ["FIN.778678/26", "FIN.764097/25"]

for idx, numero in enumerate(lista_fin):
    print(f"[{idx+1}/{len(lista_fin)}] Acessando FIN {numero}")
    extrai_fin(numero)

    # if dados_dos_chamados:
    #     registrar_chamado(
    #         dados_dos_chamados,
    #         atividade=atividadehabilitada[idx],
    #         descricao=objetos_compra[idx],
    #         identificador=str(numero),
    #         hoje=hoje,
    #         remover_manual=False
    #     )

print("✅ Encerrada a extração dos FIN.")


print("Finalizando.....")

driver.quit()
