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
from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException
#from datetime import datetime
import os
#import ctypes
#import win32com.client as win32
import gspread
#import csv

#%%

def criar_driver(pasta_downloads):
    options = Options()
    options.add_argument("--headless")
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

    print("Realizando Login Microsoft...")

    # WebDriverWait(driver, 20).until(
    #     EC.presence_of_element_located((By.ID, "i0116"))
    # ).send_keys('davi.hulse@sc.senai.br' + Keys.ENTER)
    
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "i0116"))
        ).send_keys('davi.hulse@sc.senai.br' + Keys.ENTER)
    except TimeoutException:
        print("✅ Login já autenticado. Pulando etapa de usuário.")

    # Etapa 2 – Senha com loop de tentativa
    for tentativa in range(3):
        try:
            WebDriverWait(driver, 5).until(
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
            print(f"⏳ Tentativa {tentativa+1}/3 falhou ao localizar campo de senha. Retentando...")
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

    instituto_map = {
        "Aquisições ISI Manufatura.xlsx": "ISI SM",
        "Aquisições ISI Laser.xlsx": "ISI PL",
        "Aquisições ISI Embarcados 2.xlsx": "ISI SE"
    }    
    
    dfs = []
    for nome_arquivo in nomes_esperados:
        caminho = os.path.join(pasta_downloads, nome_arquivo)
        if not os.path.exists(caminho):
            print(f"⚠️ Arquivo não encontrado: {nome_arquivo}")
            continue
        try:
            df = pd.read_excel(caminho)
            df["Instituto"] = instituto_map.get(nome_arquivo, "")
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

# Registrar alerta tendo ID_CARD como chave primária
def registrar_alerta(fin, identificador, tipo, mensagem, titulo_card="", id_card="", instituto=""):
    from datetime import datetime
    valores = worksheet_alertas.get_all_values()
    cabecalho = ["ID_CARD", "Instituto", "Número do FIN", "Título do Card",
                 "Identificador", "Tipo", "Mensagem", "Data"]

    if not valores:
        worksheet_alertas.append_row(cabecalho, value_input_option="USER_ENTERED")
        linhas_dados = []
    else:
        linhas_dados = valores[1:]

    data_atual = datetime.now().strftime("%d/%m/%Y %H:%M")
    nova_linha = [id_card, instituto, fin, titulo_card, identificador, tipo, mensagem, data_atual]

    # Busca por ID_CARD + Tipo
    for i, linha in enumerate(linhas_dados):
        if str(linha[0]).strip() == id_card and (len(linha) > 5 and str(linha[5]).strip() == tipo):
            worksheet_alertas.update(
                values=[nova_linha],
                range_name=f"A{i + 2}",
                value_input_option="USER_ENTERED"
            )
            print(f"🔁 Alerta '{tipo}' do card {id_card} atualizado.")
            return

    worksheet_alertas.append_row(nova_linha, value_input_option="USER_ENTERED")
    print(f"➕ Alerta '{tipo}' do card {id_card} registrado.")

#def remover_alerta(id_card, tipo):
def remover_alerta(id_card, tipo, numero_fin=None):
    valores = worksheet_alertas.get_all_values()
    if not valores or len(valores) <= 1:
        return
    
    # for i, linha in enumerate(valores[1:]):
    #     if str(linha[0]).strip() == id_card and (len(linha) > 5 and str(linha[5]).strip() == tipo):
    #         worksheet_alertas.delete_rows(i + 2)
    #         print(f"✅ Alerta '{tipo}' do card {id_card} removido.")
    #         return

    for i, linha in enumerate(valores[1:]):
        # Verifica por ID_CARD + Tipo
        match_id = str(linha[0]).strip() == id_card
        # OU verifica por Número do FIN + Tipo (fallback para alertas órfãos)
        match_fin = numero_fin and len(linha) > 2 and str(linha[2]).strip() == numero_fin
        match_tipo = len(linha) > 5 and str(linha[5]).strip() == tipo
        
        if (match_id or match_fin) and match_tipo:
            worksheet_alertas.delete_rows(i + 2)
            print(f"✅ Alerta '{tipo}' removido (ID_CARD={id_card}, FIN={numero_fin}).")
            return


#%% #######################

#Acessar Dados do RPA no Google Sheets
gc = gspread.service_account(filename=os.path.join(os.path.dirname(os.getcwd()), 'crested-century-386316-01c90985d6e4.json'))

#Dados Aquisições RPA
spreadsheet_rpa = gc.open("Acompanhamento_Aquisições_RPA")
worksheet_rpa = spreadsheet_rpa.worksheet("Dados")
dados_rpa = worksheet_rpa.get_all_values()

worksheet_rpa_eproc = spreadsheet_rpa.worksheet("EPROC")
dados_rpa_eproc = worksheet_rpa_eproc.get_all_values()

df_dados_rpa = pd.DataFrame(dados_rpa[1:], columns=dados_rpa[0])

df_dados_rpa_eproc = pd.DataFrame(dados_rpa_eproc[1:], columns=dados_rpa_eproc[0])

df_dados_rpa = pd.concat([df_dados_rpa, df_dados_rpa_eproc], ignore_index=True, sort=False)
df_dados_rpa['Valor R$'] = df_dados_rpa['Valor R$'].str.replace('.', '', regex=False)


#Dados FIN RPA - planilha destino
spreadsheet_fin = gc.open("Acompanhamento_FIN_RPA")
worksheet_fin = spreadsheet_fin.worksheet("Dados")
worksheet_manuais = spreadsheet_fin.worksheet("Manuais")
worksheet_ignorar = spreadsheet_fin.worksheet("Ignorar")
worksheet_alertas = spreadsheet_fin.worksheet("Alertas")


planners_urls = [
    "https://planner.cloud.microsoft/webui/plan/QXrbRoU7UEGdjE_bhw-QY2QAFn9X/view/board?tid=2cf7d4d5-bd1b-4956-acf8-2995399b2168",
    "https://planner.cloud.microsoft/webui/plan/vIOkh-y5EEuwwAlkWsRQRmQAER1C/view/board?tid=2cf7d4d5-bd1b-4956-acf8-2995399b2168",
    "https://planner.cloud.microsoft/webui/plan/By2-rKiP6EWT0TfgDNLG12QAGgc8/view/board?tid=2cf7d4d5-bd1b-4956-acf8-2995399b2168"
]

pasta_downloads = r"C:\RPA\rpa_fin\Downloads"
driver = criar_driver(pasta_downloads)

#Comentar as 3 linhas abaixo para pular o Download dos Planners em Excel
driver.get("https://planner.cloud.microsoft/webui/plan/QXrbRoU7UEGdjE_bhw-QY2QAFn9X/view/board?tid=2cf7d4d5-bd1b-4956-acf8-2995399b2168")
login_microsoft(driver)
exportar_planners(driver)

df = consolidar_planilhas(pasta_downloads)

df["Nome do Bucket"] = df["Nome do Bucket"].astype(str).str.strip()

df = df[~df["Nome do Bucket"].isin(["Brementur", "Pc de Viagem", "PC de Viagem"])]


#%%

def extrair_numero_tarefa(texto):
    if pd.isna(texto):
        return None
    
    # 1) E-PROC no formato E-PROC.00154.25, EPROC.00154.25, E PROC.00154.25
    match_eproc = re.search(r"E[\s\-]?PROC\.(\d{5})\.(\d{2})", texto, flags=re.IGNORECASE)
    if match_eproc:
        return f"E-PROC.{match_eproc.group(1)}.{match_eproc.group(2)}"

    # 2) Número de tarefa com 5, 6 ou 7 dígitos
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

#%%

def extrair_numero_documento_card(titulo_card):
    """Extrai o número do documento do título do card (após 'NF nº:')"""
    if pd.isna(titulo_card):
        return None
    match = re.search(r"NF nº:\s*([0-9\-/]+)", titulo_card, flags=re.IGNORECASE)
    if match:
        # Remove pontuação, mantém só números
        return re.sub(r"[^0-9]", "", match.group(1))
    return None


def extrair_numero_fin(texto):
    if pd.isna(texto):
        return None
    match = re.search(r"FIN[:.\s\-]*?(\d{4,6}/\d{2})", texto, flags=re.IGNORECASE)
    if match:
        return f"FIN.{match.group(1)}"
    return None

df["FIN"] = df["Itens da lista de verificação"].apply(extrair_numero_fin)


#%%

# Exportar DF para Excel:
df.to_excel("df.xlsx", index=False)

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
    
    try:
        driver.get(r'https://sesuite.fiesc.com.br/softexpert/workspace?page=home')
        WebDriverWait(driver, 30).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    except UnexpectedAlertPresentException:
        try:
            driver.switch_to.alert.accept()
        except:
            pass

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
    
    inserir_fin.clear()
    WebDriverWait(driver, 30).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    sleep(1)
    inserir_fin.send_keys(str(numfin))
    WebDriverWait(driver, 30).until(lambda d: d.execute_script('return document.readyState') == 'complete')
    sleep(1)
    inserir_fin.send_keys(Keys.ENTER)
    
    print("Aguardando SE Suite...")
        
    try:
        primeiro_item = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="st-container"]/div/div/div/div[4]/div/div[2]/div/div/div[2]/div/div[2]/div[1]/span'))
        )
        print("FIN localizado.")
    except TimeoutException:
        print("❌ Nenhum FIN encontrado. Pulando.")
        return None

    # Extrai o texto do link antes de clicar para validar DOC FISCAL
    texto_link = primeiro_item.text.strip()
    match_doc_fiscal = re.search(r"DOC FISCAL:\s*([0-9\-/]+)", texto_link, flags=re.IGNORECASE)
    doc_fiscal_sesuite = re.sub(r"[^0-9]", "", match_doc_fiscal.group(1)) if match_doc_fiscal else None
    print("Extraindo dados...")

        
    for tentativa in range(3):
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

    tipo_documento_map = {
    "cd5234b365d0e0be2f9f0f35675e4ad7": "Adiantamento de Mídias Sociais - Facebook",
    "89806d075fdecb2a7f5ef09b419f49cf": "AL (Aviso de lançamento)",
    "fbcd6551dabee395975520eac85d26ce": "Aluguel",
    "aff075fad54d65227c6f4d11a3594675": "Conhecimento de transporte",
    "8e20bb220b9b6682c00c3e6f90ad6d95": "Faturas (Energia/Água/Internet/Telefonia)",
    "b21bb8d92ca5c9d55bc72fcc4eb38263": "Insumos de Alimentação e Farmácia",
    "a8cf7668aab4f099f2c840e689008368": "Licenciamento/DPVAT",
    "39a7ac201ee522c9d13821ddfe7ce445": "Mídias Sociais - Facebook",
    "2f9be0442387599aadd5e99acc6c738d": "NF Comunicação",
    "4e67ce2c5f6e2356def6de66ff09fc56": "NF Serviço - Entre Casas (SESI/SENAI)",
    "c259d830eee9864f1d5492f937aceb42": "Notas de Débito - Entre Casas",
    "35823390368b62a8b9aee21a25ee1eca": "Pagamentos Rejeitados",
    "b6e50aa9f31672ed0f62d37f91ad8207": "Produto",
    "55b1fe3b58b61c2c46ebcc7b37fcc502": "Reembolso/Ressarcimento",
    "0de84361ee0aa6c3f785370c00ead580": "Serviço",
    "b7d11f1bff8a92db84b81adea62d6ca9": "Taxas"
    }
    
    #Preencher as demais especificações
    especificacao_map = {
    "683772e630d835dbc855afe622b9ec35": "Produto",    
    "9fd571ca945c53601f45ce940e69bcc4": "Taxas",
    "d7c5f59fec1d99fed96b0769fc2427bb": "Serviço",
    "d2fe1c2a4a00ab7cad978ae2fa913bac": "Conhecimento de transporte"
    }

    adiantamento_map = {
    "308a83de517f730f3de47ae53c296af9": "Não",
    "2e6296ec66975b8b87e0d94a73a5f391": "Sim"
    }

    tipo_compra_map = {
    "677e3c3209b3c0cc63670b3b3333546f": "Chave de compra (Sem Ordem de Compra)",
    "3c6575bae11225e7378448ea26b5f455": "Com Contrato de Compra",
    "45b792a9ff6c51b8aa198243848dcf15": "Com ordem de compra"
    }
        
    #Campos a extrair
    campos = [
        ("Número do FIN", '//*[@id="field_8a3449076f9f6db3016fe74952d0181b"]'),
        ("Data da Abertura do FIN", '//*[@id="field_8a3449076f9f6db3016fa46bce563614"]'),
        ("Tipo de Documento", '//*[@id="oidzoom_8a34490772df4a7a0172eb5952b56c38"]'),
        ("Especificação", '//*[@id="oidzoom_8a3449077843843601785b0a8d400c5c"]'),
        ("Valor pago por Adiantamento?", '//*[@id="oidzoom_8a34490770c96a380170cfe876536a31"]'),
        ("Filial Faturada", '//*[@id="field_8a3449076f9f6db3016fe80ac15f31aa"]'),
        ("CNPJ Fornecedor", '//*[@id="field_8a3449076f9f6db3016fe747b2fe17cf"]'),
        ("Número do documento", '//*[@id="field_8a3449077918207d017980762ad719ba"]'),
        ("Tipo de Compra", '//*[@id="oidzoom_8a3449076f9f6db301701a5907032a88"]'),
        ("Ordem de compra (FIN)", '//*[@id="field_8a3449076f9f6db301701adda4b73de1"]'),
        ("Contrato (FIN)", '//*[@id="field_8a3449076f9f6db301701adcaaf43da7"]'),
        ("Registro Gerado (Apontamento)", '//*[@id="field_8a34490770c96a380170cfe7a19969c2"]'),
        ("RNs", '//*[@id="field_8a3449076f9f6db3016fe7454ab71792"]'),
        ("Observações", '//*[@id="field_8a344907739c40c10174300c129a4832"]'),
        ("Número AP", '//*[@id="field_8a3449076f9f6db3016fe735db491529"]'),
        ("Data Agendada para Pagamento", '//*[@id="field_8a3449076f9f6db301701a9b6d9c3533"]'),
        ("Competência", '//*[@id="field_8a3449076f9f6db301701a6bdc3e2e27"]'),
        ("Valor Bruto a Pagar (R$)", '//*[@id="field_8a34490770c96a380170e426eed86216"]'),
        ("Valor a deduzir (R$)", '//*[@id="field_8a34490770c96a380170e427cc6e6266"]'),
        ("Valor Líquido a Pagar (R$)", '//*[@id="field_8a34490770c96a380170e4277bfb623e"]'),
        ("Nr. do documento (CAP)", '//*[@id="field_8a344907739c40c101743078077162ed"]')
    ]
               
    for nome, xpath in campos:
        try:
            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            dados_dos_chamados[nome] = element.get_attribute("value")
        except TimeoutException:
            print(f"⚠️ Campo '{nome}' não encontrado. Registrando como vazio.")
            dados_dos_chamados[nome] = ""
    
    dados_dos_chamados["Descrição"] = titulo_completo
    
    # Campos que precisam de Mapeamento (Início)
    cod_tipo_documento = dados_dos_chamados.get("Tipo de Documento")
    dados_dos_chamados["Tipo de Documento"] = tipo_documento_map.get(cod_tipo_documento, cod_tipo_documento)
    
    cod_especificacao = dados_dos_chamados.get("Especificação")
    dados_dos_chamados["Especificação"] = especificacao_map.get(cod_especificacao, cod_especificacao)
    
    cod_adiantamento = dados_dos_chamados.get("Valor pago por Adiantamento?")
    dados_dos_chamados["Valor pago por Adiantamento?"] = adiantamento_map.get(cod_adiantamento, cod_adiantamento)
    
    cod_tipo_compra = dados_dos_chamados.get("Tipo de Compra")
    dados_dos_chamados["Tipo de Compra"] = tipo_compra_map.get(cod_tipo_compra, cod_tipo_compra)
    # /Campos que precisam de Mapeamento (Fim)
             
    for janela in driver.window_handles:
        if janela != janela_principal:
            driver.switch_to.window(janela)
            driver.close()

    driver.switch_to.window(janela_principal)
    
    dados_dos_chamados["Status"] = status_texto
    
    dados_dos_chamados["_doc_fiscal_validacao"] = doc_fiscal_sesuite

    print("Dados do ", numfin, " extraídos.")
    
    return dados_dos_chamados

#%%

def registrar_fin_google_sheets(dados_fin, dados_aquisicao, worksheet_fin):

    colunas_esperadas = [
        "ID_CARD", "Título do Card",
        # --- Dados da Aquisição ---
        "Código Unidade", "Identificador", "Apelido Projeto", "Descrição", "Fonte",
        "Rubrica", "Valor Aquisição R$", "Ordem de Compra (Aquisição)",
        # --- Dados do FIN ---
        "Número do FIN", "Descrição FIN", "Status FIN", "Data da Abertura do FIN",
        "Tipo de Documento", "Especificação", "Valor pago por Adiantamento?",
        "Filial Faturada", "CNPJ Fornecedor", "Número do documento", "Tipo de Compra",
        "Ordem de compra (FIN)", "Registro Gerado (Apontamento)", "RNs", "Observações",
        "Número AP", "Data Agendada para Pagamento", "Competência",
        "Valor Bruto a Pagar (R$)", "Valor a deduzir (R$)", "Valor Líquido a Pagar (R$)",
        "Nr. do documento (CAP)"
    ]

    # Limpeza de strings
    for origem in (dados_fin, dados_aquisicao):
        for k, v in origem.items():
            if isinstance(v, str):
                origem[k] = v.replace("\n", " ").strip()

    linha = {
        "ID_CARD": dados_aquisicao.get("ID_CARD", ""),
        "Título do Card": dados_aquisicao.get("Título do Card", ""),
        # --- Aquisição ---
        "Código Unidade": dados_aquisicao.get("Código Unidade", ""),
        "Identificador": dados_aquisicao.get("Identificador", ""),
        "Apelido Projeto": dados_aquisicao.get("Apelido Projeto", ""),
        "Descrição": dados_aquisicao.get("Descrição", ""),
        "Fonte": dados_aquisicao.get("Fonte", ""),
        "Rubrica": dados_aquisicao.get("Rubrica", ""),
        "Valor Aquisição R$": dados_aquisicao.get("Valor R$", ""),
        "Ordem de Compra (Aquisição)": dados_aquisicao.get("Ordem de Compra", ""),
        # --- FIN ---
        "Número do FIN": dados_fin.get("Número do FIN", ""),
        "Descrição FIN": dados_fin.get("Descrição", ""),
        "Status FIN": dados_fin.get("Status", ""),
        "Data da Abertura do FIN": dados_fin.get("Data da Abertura do FIN", ""),
        "Tipo de Documento": dados_fin.get("Tipo de Documento", ""),
        "Especificação": dados_fin.get("Especificação", ""),
        "Valor pago por Adiantamento?": dados_fin.get("Valor pago por Adiantamento?", ""),
        "Filial Faturada": dados_fin.get("Filial Faturada", ""),
        "CNPJ Fornecedor": dados_fin.get("CNPJ Fornecedor", ""),
        "Número do documento": dados_fin.get("Número do documento", ""),
        "Tipo de Compra": dados_fin.get("Tipo de Compra", ""),
        "Ordem de compra (FIN)": dados_fin.get("Ordem de compra (FIN)", "") or dados_fin.get("Contrato (FIN)", ""),
        #"Ordem de compra (FIN)": dados_fin.get("Ordem de compra (FIN)", ""),
        "Contrato (FIN)": dados_fin.get("Contrato (FIN)", ""),
        "Registro Gerado (Apontamento)": dados_fin.get("Registro Gerado (Apontamento)", ""), 
        "RNs": dados_fin.get("RNs", ""),
        "Observações": dados_fin.get("Observações", ""),
        "Número AP": dados_fin.get("Número AP", ""),
        "Data Agendada para Pagamento": dados_fin.get("Data Agendada para Pagamento", ""),
        "Competência": dados_fin.get("Competência", ""),
        "Valor Bruto a Pagar (R$)": dados_fin.get("Valor Bruto a Pagar (R$)", ""),
        "Valor a deduzir (R$)": dados_fin.get("Valor a deduzir (R$)", ""),
        "Valor Líquido a Pagar (R$)": dados_fin.get("Valor Líquido a Pagar (R$)", ""),
        "Nr. do documento (CAP)": dados_fin.get("Nr. do documento (CAP)", "")
    }

    linha_ordenada = [linha.get(col, "") for col in colunas_esperadas]

    # Leitura existente
    valores_existentes = worksheet_fin.get_all_records()
    df_existente = pd.DataFrame(valores_existentes)

    if not df_existente.empty and "ID_CARD" in df_existente.columns:
        ids_existentes = df_existente["ID_CARD"].astype(str).tolist()
    else:
        ids_existentes = []

    id_card = str(linha["ID_CARD"])
    numero_fin = str(linha["Número do FIN"])
    identificador = linha.get("Identificador", "")


    if id_card in ids_existentes:
        idx = ids_existentes.index(id_card)
        linha_planilha = idx + 2  # header + base 1
        worksheet_fin.update(
            values=[linha_ordenada],
            range_name=f"A{linha_planilha}",
            value_input_option="USER_ENTERED"
        )
        print(f"🔁 Card {id_card} ({numero_fin}) atualizado na linha {linha_planilha}.")
    else:
        worksheet_fin.append_row(linha_ordenada, value_input_option="USER_ENTERED")
        print(f"➕ Card {id_card} ({numero_fin}) inserido como nova linha.")

    # (Re)carrega planilha com cabeçalhos
    valores_existentes = worksheet_fin.get_all_values()
    df_existente = pd.DataFrame(valores_existentes[1:], columns=valores_existentes[0])


    # Verifica e atualiza linha de saldo
    if identificador:
        registros_mesmo_chamado = df_existente[
            (df_existente["Identificador"] == identificador)
            & (df_existente["Número do FIN"] != "Saldo")
        ]

        soma_fins = (
                    registros_mesmo_chamado["Valor Líquido a Pagar (R$)"]
                    .astype(str)
                    .str.strip()
                    .str.replace(r"\.", "", regex=True)
                    .str.replace(",", ".", regex=False)
                    .pipe(pd.to_numeric, errors="coerce")
                    .fillna(0)
                    .sum()
                )

        valor_oc = dados_aquisicao.get("Valor R$", "")
        try:
            valor_oc_float = float(str(valor_oc).replace(".", "").replace(",", "."))
        except:
            valor_oc_float = 0.0
    
        saldo = valor_oc_float - soma_fins
        if saldo < -9.99:
            msg = f"Saldo negativo na soma de FINs: {saldo:,.2f}."
            print(f"⚠️ ATENÇÃO: {msg} Identificador: {identificador}.")
            registrar_alerta(numero_fin, identificador, "SALDO_NEGATIVO", msg, id_card=id_card)
        else:
            remover_alerta(id_card, "SALDO_NEGATIVO")
    
        idx_saldo = df_existente[
            (df_existente["Identificador"] == identificador)
            & (df_existente["Número do FIN"] == "Saldo")
        ].index
    
        if saldo != 0:
            linha_saldo = {col: "" for col in colunas_esperadas}
            linha_saldo["Código Unidade"] = linha.get("Código Unidade", "")
            linha_saldo["Identificador"] = identificador
            linha_saldo["Apelido Projeto"] = linha.get("Apelido Projeto", "")
            linha_saldo["Descrição"] = linha.get("Descrição", "")
            linha_saldo["Fonte"] = linha.get("Fonte", "")
            linha_saldo["Rubrica"] = linha.get("Rubrica", "")
            linha_saldo["Valor Aquisição R$"] = linha.get("Valor Aquisição R$", "")
            linha_saldo["Número do FIN"] = "Saldo"
            linha_saldo["Valor Bruto a Pagar (R$)"] = f"{saldo:,.2f}".replace(".", "X").replace(",", ".").replace("X", ",")
    
            if len(idx_saldo) > 0:
                worksheet_fin.update(values=[[linha_saldo[col] for col in df_existente.columns]],
                                     range_name=f"A{idx_saldo[0]+2}", value_input_option="USER_ENTERED")
                #worksheet_fin.update(f"A{idx_saldo[0]+2}", [linha_saldo[col] for col in df_existente.columns])
            else:
                worksheet_fin.insert_rows([[linha_saldo.get(col, "") for col in df_existente.columns]],
                                          row=len(df_existente)+2, value_input_option="USER_ENTERED")
    
        elif len(idx_saldo) > 0:
            worksheet_fin.delete_rows(int(idx_saldo[0]) + 2)

#%% #########################

login_sesuite()

fins_em_dados = worksheet_fin.col_values(
    worksheet_fin.row_values(1).index("ID_CARD") + 1
)

fins_em_manuais = [v.strip() for v in worksheet_manuais.col_values(1) if v.strip()]
fins_ignorados = set(v.strip() for v in worksheet_ignorar.col_values(1) if v.strip())

# Remove FINs ignorados que estão na lista manual
for fin in fins_em_manuais[:]:  # cópia da lista para iterar
    if fin in fins_ignorados:
        print(f"⚠️ {fin} está na lista de ignorados. Removendo da aba Manuais.")
        todas_linhas = worksheet_manuais.col_values(1)
        for i, valor in enumerate(todas_linhas):
            if valor.strip() == fin:
                worksheet_manuais.delete_rows(i + 1)
                break
# Atualiza a lista após remoções
fins_em_manuais = [v.strip() for v in worksheet_manuais.col_values(1) if v.strip()]


# Remove alertas de FINs ignorados

# Remove alertas de FINs ignorados
print("🧹 Removendo alertas de FINs ignorados...")
print(f"FINs ignorados: {fins_ignorados}")
for fin_ignorado in fins_ignorados:
    print(f"Processando FIN ignorado: {fin_ignorado}")
    linha_ignorada = df[df["FIN"] == fin_ignorado]
    if not linha_ignorada.empty:
        id_card_ignorado = linha_ignorada.iloc[0].get("Identificação da tarefa", "")
        print(f"  -> ID_CARD: {id_card_ignorado}")
        if id_card_ignorado:
            remover_alerta(id_card_ignorado, "FALHA_EXTRACAO", fin_ignorado)
            sleep(1)
            remover_alerta(id_card_ignorado, "DOC_DIVERGENTE", fin_ignorado)
            sleep(1)
            remover_alerta(id_card_ignorado, "SEM_NF_CARD", fin_ignorado)
            sleep(1)
            remover_alerta(id_card_ignorado, "AQUISICAO_NAO_ENCONTRADA", fin_ignorado)
            sleep(1)


#%% --- Primeiro: manuais ---
print("📌 Iniciando extração de FINs manuais...")
for idx, fin in enumerate(fins_em_manuais):
    print(f"[MANUAL {idx+1}/{len(fins_em_manuais)}] Processando {fin}")
        
    linha_com_fin = df[df["FIN"] == fin]
    titulo_card = linha_com_fin.iloc[0]["Nome da tarefa"] if not linha_com_fin.empty else ""
    numero_tarefa = linha_com_fin.iloc[0]["Numero Tarefa"] if not linha_com_fin.empty else ""
    id_card_atual = linha_com_fin.iloc[0].get("Identificação da tarefa", "") if not linha_com_fin.empty else ""
    instituto = linha_com_fin.iloc[0].get("Instituto", "") if not linha_com_fin.empty else ""
    
    dados_fin = extrai_fin(fin)
    
    if not dados_fin:
        msg = "Falha ao extrair dados do FIN."
        print(f"⚠️ {msg}")
        registrar_alerta(fin, str(numero_tarefa) if numero_tarefa else "", "FALHA_EXTRACAO",
                         msg, titulo_card, id_card=id_card_atual, instituto=instituto)
        continue

    if linha_com_fin.empty:
        print(f"⚠️ {fin} não encontrado na planilha consolidada. Pulando.")
        continue
    
    linha_aquisicao = df_dados_rpa[
        df_dados_rpa["Identificador"].astype(str).str.zfill(6) == str(numero_tarefa).zfill(6)
    ]
    
    if linha_aquisicao.empty:
        msg = f"Nenhuma aquisição encontrada para a Tarefa {numero_tarefa}."
        print(f"⚠️ {msg}")
        registrar_alerta(fin, str(numero_tarefa) if numero_tarefa else "", "AQUISICAO_NAO_ENCONTRADA",
                         msg, titulo_card, id_card=id_card_atual, instituto=instituto)
        continue    

    dados_aquisicao = linha_aquisicao.iloc[0].to_dict()
    #dados_aquisicao["ID_CARD"] = linha_com_fin.iloc[0].get("Identificação da tarefa", "")
    
    dados_aquisicao["Título do Card"] = titulo_card
    
    dados_aquisicao["ID_CARD"] = id_card_atual
   
    
    registrar_fin_google_sheets(dados_fin, dados_aquisicao, worksheet_fin)
    
    # Remove alerta, se existir (FIN foi processado com sucesso)
    remover_alerta(id_card_atual, "FALHA_EXTRACAO")
    sleep(1)
    remover_alerta(id_card_atual, "DOC_DIVERGENTE")
    sleep(1)
    remover_alerta(id_card_atual, "SEM_NF_CARD")
    sleep(1)
    remover_alerta(id_card_atual, "AQUISICAO_NAO_ENCONTRADA")
    sleep(1)
    
    # /Remove alerta, se existir (FIN foi processado com sucesso)
    

    todas_linhas = worksheet_manuais.col_values(1)
    for i, valor in enumerate(todas_linhas):
        if valor.strip() == fin:
            worksheet_manuais.delete_rows(i + 1)
            print(f"🗑️ {fin} removido da aba Manuais.")
            break

print("✅ Encerrada a extração de FINs manuais. Continuando para os demais...")

#%% --- Depois: df ---
print("📌 Iniciando extração dos FINs do Planner...")
lista_fins = df[df["FIN"].notna()]["FIN"].unique().tolist()

# Carrega dados existentes para verificar status
valores_dados = worksheet_fin.get_all_values()
df_completo = pd.DataFrame(valores_dados[1:], columns=valores_dados[0])

for idx, fin in enumerate(lista_fins):
    print(f"[{idx+1}/{len(lista_fins)}] Processando {fin}")

    if fin in fins_ignorados:
        print(f"⏭️ {fin} está na lista de ignorados. Pulando.")
        continue

    linha_com_fin = df[df["FIN"] == fin]
    if linha_com_fin.empty:
        print(f"⚠️ {fin} não encontrado na planilha consolidada. Pulando.")
        continue

    id_card_atual = linha_com_fin.iloc[0].get("Identificação da tarefa", "")
    
    if id_card_atual in fins_em_dados:
        # Verifica se o FIN já está encerrado
        linha_existente = df_completo[df_completo["ID_CARD"] == id_card_atual]
        if not linha_existente.empty:
            status_atual = str(linha_existente.iloc[0].get("Status FIN", "")).strip()
            if status_atual.lower() == "encerrado":
                continue
            else:
                continue

    # id_card_atual = linha_com_fin.iloc[0].get("Identificação da tarefa", "")
    
    # if id_card_atual in fins_em_dados:
    #     print(f"⏭️ Card {id_card_atual} ({fin}) já existe em Dados. Pulando.")
    #     continue

    titulo_card = linha_com_fin.iloc[0]["Nome da tarefa"]
    numero_doc_card = extrair_numero_documento_card(titulo_card)
    numero_tarefa = linha_com_fin.iloc[0]["Numero Tarefa"]
    instituto = linha_com_fin.iloc[0].get("Instituto", "")
    
    dados_fin = extrai_fin(fin)
    if not dados_fin:
        msg = "Falha ao extrair dados do FIN."
        print(f"❌ {msg}")
        registrar_alerta(fin, str(numero_tarefa) if numero_tarefa else "", "FALHA_EXTRACAO",
                         msg, titulo_card, id_card=id_card_atual, instituto=instituto)
        continue    
    
    # Validação: número do documento
    if numero_doc_card:
        doc_fiscal_fin = dados_fin.get("_doc_fiscal_validacao", "")
        # Remove zeros à esquerda de ambos para comparação
        numero_doc_card_sem_zeros = numero_doc_card.lstrip('0')
        doc_fiscal_fin_sem_zeros = doc_fiscal_fin.lstrip('0') if doc_fiscal_fin else ''
        if doc_fiscal_fin and numero_doc_card_sem_zeros != doc_fiscal_fin_sem_zeros:

    # if numero_doc_card:
    #     doc_fiscal_fin = dados_fin.get("_doc_fiscal_validacao", "")
    #     if doc_fiscal_fin and numero_doc_card != doc_fiscal_fin:
            msg = f"Número do documento divergente. Planner: {numero_doc_card} | FIN: {doc_fiscal_fin}"
            print(f"⚠️ ATENÇÃO: {msg}")
            registrar_alerta(fin, str(numero_tarefa) if numero_tarefa else "", "DOC_DIVERGENTE",
                             msg, titulo_card, id_card=id_card_atual, instituto=instituto)
            continue
    else:
        msg = "Título do card sem 'NF nº:' para validação."
        print(f"⚠️ ATENÇÃO: {msg}")
        registrar_alerta(fin, str(numero_tarefa) if numero_tarefa else "", "SEM_NF_CARD",
                         msg, titulo_card, id_card=id_card_atual, instituto=instituto)


    linha_aquisicao = df_dados_rpa[
        df_dados_rpa["Identificador"].astype(str).str.zfill(6) == str(numero_tarefa).zfill(6)
    ]
    if linha_aquisicao.empty:
        msg = f"Nenhuma aquisição encontrada para a Tarefa {numero_tarefa}."
        print(f"⚠️ {msg}")
        registrar_alerta(fin, str(numero_tarefa) if numero_tarefa else "", "AQUISICAO_NAO_ENCONTRADA",
                         msg, titulo_card, id_card=id_card_atual, instituto=instituto)
        continue

    dados_aquisicao = linha_aquisicao.iloc[0].to_dict()
    dados_aquisicao["ID_CARD"] = id_card_atual
    dados_aquisicao["Título do Card"] = titulo_card

    registrar_fin_google_sheets(dados_fin, dados_aquisicao, worksheet_fin)
    
    # Remove alerta, se existir (FIN foi processado com sucesso)
    remover_alerta(id_card_atual, "FALHA_EXTRACAO")
    sleep(1)
    remover_alerta(id_card_atual, "DOC_DIVERGENTE")
    sleep(1)
    remover_alerta(id_card_atual, "SEM_NF_CARD")
    sleep(1)
    remover_alerta(id_card_atual, "AQUISICAO_NAO_ENCONTRADA")
    sleep(1)    
    # /Remove alerta, se existir (FIN foi processado com sucesso)

    # Se estava em Manuais, remove de lá após inserir com sucesso em Dados
    if fin in fins_em_manuais:
        todas_linhas = worksheet_manuais.col_values(1)
        for i, valor in enumerate(todas_linhas):
            if valor.strip() == fin:
                worksheet_manuais.delete_rows(i + 1)
                print(f"🗑️ {fin} removido da aba Manuais.")
                break

#%% Recalculando saldos de todos os identificadores:
### Removido pois demorava demais, calculava todos os milhares de saldos.

# print("📊 Recalculando saldos de todos os identificadores...")

# # Recarrega planilha completa
# valores_existentes = worksheet_fin.get_all_values()
# df_completo = pd.DataFrame(valores_existentes[1:], columns=valores_existentes[0])

# # Lista de identificadores únicos (excluindo Saldo)
# identificadores_unicos = df_completo[
#     (df_completo["Identificador"] != "") & 
#     (df_completo["Número do FIN"] != "Saldo")
# ]["Identificador"].unique()

# for identificador in identificadores_unicos:
#     # Filtra FINs deste identificador
#     registros = df_completo[
#         (df_completo["Identificador"] == identificador) &
#         (df_completo["Número do FIN"] != "Saldo")
#     ]
    
#     if registros.empty:
#         continue
    
#     # Soma valores líquidos
#     soma_fins = (
#         registros["Valor Líquido a Pagar (R$)"]
#         .astype(str)
#         .str.strip()
#         .str.replace(r"\.", "", regex=True)
#         .str.replace(",", ".", regex=False)
#         .pipe(pd.to_numeric, errors="coerce")
#         .fillna(0)
#         .sum()
#     )
    
#     # Pega valor da aquisição (primeira linha)
#     valor_aquisicao_str = registros.iloc[0]["Valor Aquisição R$"]
#     try:
#         valor_aquisicao = float(str(valor_aquisicao_str).replace(".", "").replace(",", "."))
#     except:
#         valor_aquisicao = 0.0
    
#     saldo = valor_aquisicao - soma_fins
    
#     # Pega ID_CARD e FIN da primeira linha para o alerta
#     id_card_ref = registros.iloc[0]["ID_CARD"]
#     numero_fin_ref = registros.iloc[0]["Número do FIN"]
    
#     # Atualiza ou remove alerta
#     if saldo < -9.99:
#         msg = f"Saldo negativo na aquisição: {saldo:,.2f}."
#         instituto_ref = registros.iloc[0]["Instituto"] if "Instituto" in registros.columns else ""
#         registrar_alerta(numero_fin_ref, identificador, "SALDO_NEGATIVO", msg, id_card=id_card_ref, instituto=instituto_ref)
#         #registrar_alerta(numero_fin_ref, identificador, "SALDO_NEGATIVO", msg, id_card=id_card_ref)
#     else:
#         remover_alerta(id_card_ref, "SALDO_NEGATIVO")
    
#     sleep(1.5)

# print("✅ Recálculo de saldos concluído.")

#%%

print("📊 Verificando alertas de saldo existentes...")

# Carrega apenas os alertas de SALDO_NEGATIVO
valores_alertas = worksheet_alertas.get_all_values()

df_alertas = pd.DataFrame(valores_alertas[1:], columns=valores_alertas[0]) if len(valores_alertas) > 1 else pd.DataFrame()

alertas_saldo = df_alertas[df_alertas["Tipo"] == "SALDO_NEGATIVO"][["ID_CARD", "Identificador"]].to_dict('records') if not df_alertas.empty else []


if not alertas_saldo:
    print("✅ Nenhum alerta de saldo para verificar.")
else:
    print(f"Encontrados {len(alertas_saldo)} alertas de saldo para verificar.")
    
    # Carrega planilha de dados
    valores_existentes = worksheet_fin.get_all_values()
    df_completo = pd.DataFrame(valores_existentes[1:], columns=valores_existentes[0])
    
    for alerta in alertas_saldo:
        identificador = alerta["Identificador"]
        id_card_ref = alerta["ID_CARD"]    
    
        # Filtra FINs deste identificador
        registros = df_completo[
            (df_completo["Identificador"] == identificador) &
            (df_completo["Número do FIN"] != "Saldo")
        ]
        
        if registros.empty:
            print(f"⚠️ Identificador {identificador} não encontrado na planilha. Removendo alerta.")
            remover_alerta(id_card_ref, "SALDO_NEGATIVO")
            sleep(1)
            continue
        
        # Soma valores líquidos
        soma_fins = (
            registros["Valor Líquido a Pagar (R$)"]
            .astype(str).str.strip()
            .str.replace(r"\.", "", regex=True)
            .str.replace(",", ".", regex=False)
            .pipe(pd.to_numeric, errors="coerce")
            .fillna(0).sum()
        )
        
        # Pega valor da aquisição
        valor_aquisicao_str = registros.iloc[0]["Valor Aquisição R$"]
        try:
            valor_aquisicao = float(str(valor_aquisicao_str).replace(".", "").replace(",", "."))
        except:
            valor_aquisicao = 0.0
        
        saldo = valor_aquisicao - soma_fins
        
        numero_fin_ref = registros.iloc[0]["Número do FIN"]
        instituto_ref = registros.iloc[0].get("Instituto", "")
        
        # Atualiza ou remove alerta
        if saldo < -9.99:
            msg = f"Saldo negativo na Aquisição: {saldo:,.2f}."
            registrar_alerta(numero_fin_ref, identificador, "SALDO_NEGATIVO", msg, id_card=id_card_ref, instituto=instituto_ref)
            print(f"🔁 Alerta de saldo atualizado: Identificador {identificador}, saldo {saldo:,.2f}")
        else:
            remover_alerta(id_card_ref, "SALDO_NEGATIVO")
            print(f"✅ Alerta de saldo removido: Identificador {identificador}, saldo agora {saldo:,.2f}")
        
        sleep(1)
    
    print("✅ Verificação de alertas de saldo concluída.")
    
    #%%

print("Finalizando.....")

driver.quit()

sleep(3)