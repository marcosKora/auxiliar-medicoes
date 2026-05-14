import time
import csv
import os
import re
import requests
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import sys
import json
import threading
import eel

# Inicializa o Eel
eel.init('web')

def resource_path(relative_path):
    """ Retorna o caminho absoluto para o recurso, funciona em dev e no PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- CONFIGURAÇÕES DE ACESSO ---
URL_CONTROLE = "https://robomedicoesaccesscontrol-default-rtdb.firebaseio.com/.json"

def verificar_acesso_remoto():
    try:
        try:
            nome_usuario = os.getlogin().upper().replace(".", "_")
        except:
            nome_usuario = os.environ.get('USERNAME', 'DESCONHECIDO').upper().replace(".", "_")

        response = requests.get(URL_CONTROLE, timeout=10)
        dados = response.json()
        
        # Pega a lista de autorizados
        autorizados = dados.get("autorizados", {})
        
        if autorizados and autorizados.get(nome_usuario) is True:
            return True
        else:
            return False
    except Exception as e:
        return False

def verificar_versao():
    try:
        response = requests.get(URL_CONTROLE, timeout=10)
        dados = response.json()
        return dados.get("versao_minima", "0.0.0")
    except:
        return "0.0.0"


# --- CONFIGURAÇÕES DE ARQUIVOS ---
CONFIG_FILE = "credenciais.txt"
BACKUP_FILE = "historico_pedidos.txt"
ERROR_FILE = "erros_pedidos.txt"

def carregar_credenciais():
    creds = {"V360_USER": "", "V360_PASS": "", "SAP_USER": "", "SAP_PASS": "", "KORA_MED_PASS": ""}
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            for linha in f:
                if "=" in linha:
                    k, v = linha.strip().split("=", 1)
                    if k in creds: creds[k] = v
    return creds

def salvar_credencial(chave, valor):
    creds = carregar_credenciais()
    creds[chave] = valor
    with open(CONFIG_FILE, "w") as f:
        for k, v in creds.items():
            f.write(f"{k}={v}\n")

def salvar_backup(id_v360, resultado):
    horario = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    with open(BACKUP_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{horario}] ID: {id_v360} - {resultado}\n")

def salvar_erro_txt(id_v360, erro):
    horario = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    with open(ERROR_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{horario}] ID: {id_v360} - {erro}\n")

# 999999: salva metrica no csv para historico
def salvar_metrica(id_v, status):
    agora = datetime.now()
    with open("metricas.csv", "a", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([agora.strftime("%d/%m/%Y"), agora.strftime("%H:%M:%S"), id_v, status])        

# 999999: carrega metricas do csv com filtro de data
# MODIFIQUE a função carregar_metricas para incluir 'solicitante':
def carregar_metricas(data_inicio=None, data_fim=None):
    if not os.path.exists("metricas.csv"):
        return {"sucesso": 0, "erro": 0, "solicitante": 0, "total": 0}
    
    sucesso = 0
    erro = 0
    solicitante = 0  # NOVO
    
    with open("metricas.csv", "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        for row in reader:
            if len(row) < 4:
                continue
            data_linha, hora, id_v, status = row
            if data_inicio and data_fim:
                if data_inicio <= data_linha <= data_fim:
                    if status == "sucesso": sucesso += 1
                    elif status == "solicitante": solicitante += 1  # NOVO
                    else: erro += 1
            elif data_inicio:
                if data_linha == data_inicio:
                    if status == "sucesso": sucesso += 1
                    elif status == "solicitante": solicitante += 1  # NOVO
                    else: erro += 1
    
    return {
        "sucesso": sucesso, 
        "erro": erro, 
        "solicitante": solicitante,  # NOVO
        "total": sucesso + erro + solicitante
    }

# --- EXPOSED FUNCTIONS PARA EEL ---

@eel.expose
def verificar_acesso():
    """Verifica acesso remoto e retorna status"""
    return verificar_acesso_remoto()

@eel.expose
def get_versao():
    """Retorna versão atual e mínima"""
    VERSAO_ATUAL = "3.0.0"
    versao_minima = verificar_versao()
    return {"atual": VERSAO_ATUAL, "minima": versao_minima}

@eel.expose
def get_credenciais():
    """Retorna credenciais salvas"""
    return carregar_credenciais()

@eel.expose
def save_credencial(chave, valor):
    """Salva uma credencial"""
    salvar_credencial(chave, valor)
    return True
@eel.expose
def get_metricas(opcao="sessao", data_inicio=None, data_fim=None):
    """Retorna métricas baseado no filtro"""
    hoje = datetime.now().strftime("%d/%m/%Y")
    ontem = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
    
    if opcao == "hoje":
        return carregar_metricas(data_inicio=hoje)
    elif opcao == "ontem":
        return carregar_metricas(data_inicio=ontem)
    elif opcao == "data":
        return carregar_metricas(data_inicio=data_inicio)
    elif opcao == "periodo":
        return carregar_metricas(data_inicio=data_inicio, data_fim=data_fim)
    else:
        return carregar_metricas()

@eel.expose
def processar_ids(ids_lista):
    """Inicia o processamento dos IDs"""
    # Inicia thread de automação
    threading.Thread(target=executar_automacao, args=(ids_lista,), daemon=True).start()
    return True

@eel.expose
def limpar_logs():
    """Limpa arquivos de log"""
    try:
        if os.path.exists("historico_pedidos.txt"):
            os.remove("historico_pedidos.txt")
        if os.path.exists("erros_pedidos.txt"):
            os.remove("erros_pedidos.txt")
        return True
    except:
        return False

# Funções de callback para o frontend
def atualizar_log_frontend(mensagem, tipo="info"):
    """Envia log para o frontend"""
    timestamp = datetime.now().strftime('%H:%M:%S')
    try:
        eel.addLog(timestamp, mensagem, tipo)()
    except:
        pass

def atualizar_sucesso_frontend(id_v, num_pedido):
    """Envia pedido pronto para a tab específica"""
    timestamp = datetime.now().strftime('%H:%M:%S')
    try:
        eel.addPedidoPronto(timestamp, id_v, num_pedido)()
    except:
        pass

def atualizar_erro_frontend(id_v, mensagem):
    """Envia erro para a tab específica"""
    timestamp = datetime.now().strftime('%H:%M:%S')
    try:
        eel.addErro(timestamp, id_v, mensagem)()
    except:
        pass

def atualizar_solicitante_frontend(id_v, mensagem):
    """Envia solicitante para a tab específica"""
    timestamp = datetime.now().strftime('%H:%M:%S')
    try:
        eel.addSolicitante(timestamp, id_v, mensagem)()
    except:
        pass    

def atualizar_progresso_frontend(atual, total):
    """Atualiza barra de progresso no frontend"""
    try:
        eel.updateProgress(atual, total)()
    except:
        pass

def atualizar_metricas_frontend(sucesso, erro, solicitante, total):
    """Atualiza métricas no frontend"""
    try:
        eel.updateMetrics(sucesso, erro, solicitante, total)()
    except:
        pass

def executar_automacao(ids_processar):
    """Função principal de automação (MANTIDA IGUAL)"""
    # Usando listas mutáveis para contadores (evita nonlocal)
    contadores = [0, 0, 0]  # [sucesso, erro, solicitante]
    cont_total = len(ids_processar)
    
    caminho_driver = os.path.join(os.getcwd(), "chromedriver.exe")
    servico = Service(caminho_driver)
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--window-size=1920,1080")
    
    creds = carregar_credenciais()
    
    driver = webdriver.Chrome(service=servico, options=chrome_options)
    wait = WebDriverWait(driver, 45)

    def esperar_elemento(by, selector, timeout=30):
        """Espera DINÂMICA - rápido em PCs bons, paciente em PCs lentos"""
        try:
            return WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((by, selector))
            )
        except:
            return None

    def esperar_clicavel(by, selector, timeout=30):
        """Espera elemento ficar clicável"""
        try:
            return WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((by, selector))
            )
        except:
            return None    

    def focar_sap():
        driver.switch_to.default_content()
        try:
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            if iframes: driver.switch_to.frame(0)
        except: pass

    def clicar_com_retry(seletor, tipo_seletor, nome_campo, seletor_espera=None):
        for tentativa in range(3):
            try:
                elemento = wait.until(EC.element_to_be_clickable((tipo_seletor, seletor)))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
                time.sleep(0.8)
                elemento.click()
                if seletor_espera:
                    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, seletor_espera)))
                return True
            except:
                pass
                time.sleep(1)
        return False

    def enviar_ao_solicitante(id_v, tipo_erro, org_atual=None):
        try:
            driver.switch_to.window(v360_handle)
            driver.get(f"https://kora.virtual360.io/nf/acceptance_terms/{id_v}")
            time.sleep(2)

            # bom dia/tarde/noite dependendo do horário
            hora = datetime.now().hour
            if hora < 12: saudacao = "Bom dia!"
            elif hora < 18: saudacao = "Boa tarde!"
            else: saudacao = "Boa noite!"

            msg = ""

            # carregar mensagens do txt mensagensSolicitante.txt, na pasta do projeto, para permitir personalização da mensagem sem alterar o código
            mensagens_dict = {}
            if os.path.exists("mensagensSolicitante.txt"):
                with open("mensagensSolicitante.txt", "r", encoding="utf-8") as f:
                    for linha in f:
                        if "::" in linha:
                            k, v = linha.strip().split("::", 1)
                            mensagens_dict[k] = v

            if tipo_erro == "inexistente_sap":
                msg = mensagens_dict.get("inexistente_sap", f"{saudacao} A solicitação não foi encontrada na base SAP, favor liberar a solicitação caso não esteja liberada ou disponibilizar uma nova.")
            elif tipo_erro == "de_para_errado":
                if org_atual in ["1418", "2001", "2901"]:
                    msg = mensagens_dict.get("de_para_errado_especifico", f"{saudacao} A solicitação não refletiu corretamente, favor disponibilizar uma nova.")
                else:
                    msg = mensagens_dict.get("de_para_errado_geral", f"{saudacao} A solicitação MV disponibilizada não está refletindo no SAP, por gentileza poderia gerar uma nova? Favor verificar também o material usado, de acordo com a unidade da medição, na planilha: https://docs.google.com/spreadsheets/d/1ho46nlMd3eDM2Axce8XBx2RNsQ2h_r06/edit?gid=8752513#gid=8752513")
            elif tipo_erro == "cnpj_sem_cadastro":
                msg = mensagens_dict.get("cnpj_sem_cadastro", f"{saudacao} Informo que o fornecedor não possui cadastro na base SAP, por gentileza solicitar o cadastro do mesmo *para a unidade da medição*. Segue link para realizar a solicitação: https://www.appsheet.com/start/44e01724-465a-42f5-9680-9884977096ad#view=Nova%20Solicita%C3%A7%C3%A3o")
            elif tipo_erro == "cnpj_sem_expansao":
                msg = mensagens_dict.get("cnpj_sem_expansao", f"{saudacao} Informo que o fornecedor não está expandido na base SAP para a unidade, por gentileza solicitar a expansão do mesmo *para a unidade da medição*. Segue link para realizar a solicitação: https://www.appsheet.com/start/44e01724-465a-42f5-9680-9884977096ad#view=Nova%20Solicita%C3%A7%C3%A3o")

            # botão "ir para ações pendentes" do v360
            btn_pendentes = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.processable-pending-actions-btn")))
            time.sleep(0.5)
            btn_pendentes.click()
            time.sleep(0.5)

            # botão de enviar para o solicitante
            btn_env_sol = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Enviar para solicitante')]")))
            btn_env_sol.click()
            time.sleep(1)
            
            # confirmar pop-up de confirmação que abre, caso abra
            try:
                alert = driver.switch_to.alert
                alert.accept()
            except: pass
            
            time.sleep(1)

            # digita mensagem no campo de texto para enviar ao solicitante
            editor = wait.until(EC.presence_of_element_located((By.ID, "note_text_show")))
            driver.execute_script(f"arguments[0].innerHTML = '<div>{msg}</div>';", editor)
            time.sleep(1)

            # confirma envio da mensagem para o solicitante
            btn_confirmar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and contains(text(), 'Enviar para solicitante')]")))
            btn_confirmar.click()
            
            # confirmar pop-up de confirmação que abre, caso abra
            time.sleep(3)
            try:
                driver.switch_to.alert.accept()
            except: pass
            
            atualizar_log_frontend(f"ID {id_v}: Enviado para o solicitante", "warning")
            atualizar_solicitante_frontend(id_v, f"Enviado ao solicitante: {tipo_erro}")
            # Registrar como solicitante na sidebar e métricas
            contadores[2] += 1  # solicitante
            salvar_backup(id_v, f"ENVIADO SOLICITANTE: {tipo_erro}")
            salvar_metrica(id_v, "solicitante")
            atualizar_metricas_frontend(contadores[0], contadores[1], contadores[2], cont_total)
            
            # --- marcar como enviado ao solicitante no kora-medicoes.web.app ---
            driver.switch_to.window(kora_handle)
            campo_pesquisa = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='Pesquisar...']")))
            campo_pesquisa.clear()
            campo_pesquisa.send_keys(id_v)
            time.sleep(0.5)
            
            btn_obs = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@title='Observação']")))
            btn_obs.click()
            time.sleep(1)
            
            btn_enviado_sol = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Enviado para o solicitante')]")))
            time.sleep(0.5)
            btn_enviado_sol.click()
            time.sleep(0.5)
        except Exception as e_envio:
            atualizar_log_frontend(f"Erro ao enviar para solicitante: {e_envio}", "error")

    def processar_guarda_chuva(id_v360, kora_handle, v360_handle):
        # caso seja pedido guarda-chuva, precisa pegar o número do pedido e o tipo do pedido no Kora para preencher no V360, então aqui tem um processo específico para isso
        driver.switch_to.window(kora_handle)
        campo_pesquisa = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='Pesquisar...']")))
        campo_pesquisa.clear()
        campo_pesquisa.send_keys(id_v360)
        time.sleep(1.5)
        
        # pegar o número do pedido
        try:
            num_pedido_el = driver.find_element(By.XPATH, "//button[contains(text(), '45') and contains(@style, 'rgb(19, 62, 81)')]")
            num_pedido = num_pedido_el.text.strip()
        except:
            num_pedido = driver.find_element(By.CSS_SELECTOR, "button[style*='background-color: rgb(19, 62, 81)']").text.strip()
        
        # pegar o tipo do pedido (fixo ou variavel)
        tipo_info = driver.find_element(By.CSS_SELECTOR, "span.bg-slate-100").text
        
        atualizar_log_frontend(f"Possui contrato GED e pedido pronto ({tipo_info}). Liberando medição.")
        time.sleep(0.5)
        # preencher no v360
        driver.switch_to.window(v360_handle)
        driver.get(f"https://kora.virtual360.io/nf/acceptance_terms/{id_v360}")
        
        # ver se está na etapa certa do setor de contratos...
        try:
            titulo_etapa = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "checkout-bar-item-title"))).text
        except:
            contadores[1] += 1  # erro
            salvar_erro_txt(id_v360, "Página do V360 não carregou a tempo")
            atualizar_erro_frontend(id_v360, "Página do V360 não carregou a tempo")
            salvar_metrica(id_v360, "erro")
            atualizar_log_frontend(f"Aviso: ID {id_v360} - Página não carregou.", "warning")
            atualizar_metricas_frontend(contadores[0], contadores[1], contadores[2], cont_total)
            return False
        
        btn_editar = wait.until(EC.element_to_be_clickable((By.ID, "nav-edit-tab")))
        time.sleep(1)
        btn_editar.click()
        time.sleep(1.5)
        
        # colocar o pedido na medição via js 
        script_final = """
        var pedido = arguments[0];
        var campos = document.querySelectorAll('#acceptance_term_purchase_order');
        var campoAtivo = null;
        for (var i = 0; i < campos.length; i++) {
            if (campos[i].offsetWidth > 0 || campos[i].offsetHeight > 0) {
                campoAtivo = campos[i]; break;
            }
        }
        if (campoAtivo) {
            campoAtivo.focus(); campoAtivo.value = pedido;
            campoAtivo.dispatchEvent(new Event('input', { bubbles: true }));
            campoAtivo.dispatchEvent(new Event('change', { bubbles: true }));
        }
        """
        driver.execute_script(script_final, num_pedido)
        time.sleep(0.5); time.sleep(0.5); time.sleep(0.5)
        
        # selecionar o tipo de pedido correto no v360 (guarda-chuva variável ou fixo)
        span_select2 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "[aria-labelledby='select2-acceptance_term_items_attributes_0_cf_tipo_de_pedido-container']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", span_select2)
        time.sleep(1)
        span_select2.click()
        
        opcao_texto = "Pedido Guarda-Chuva Variavel" if "Variavel" in tipo_info else "Pedido Guarda-Chuva Fixo"
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[contains(text(), '{opcao_texto}')]"))).click()
        time.sleep(0.5)
        # salvar medição
        btn_salvar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@name='status_id' and contains(text(), 'Salvar')]")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_salvar)
        time.sleep(0.5)
        btn_salvar.click()
        
        # botão "ir para ações pendentes" do v360
        btn_pendentes = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.processable-pending-actions-btn")))
        time.sleep(0.5)
        btn_pendentes.click()
        
        # botão de "tentar novamente" (libera medição)
        btn_tentar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Tentar Novamente__IServ')]")))
        time.sleep(0.5)
        btn_tentar.click()
        
        # verificar se foi para a alçada do solicitante (seguiu) ou se ficou na mesma etapa, com logica de atualizar a pagina algumas vezes para evitar falhas de atualização do status do v360, que é bem frequente.
        status_correto = False
        esperar_elemento(By.CLASS_NAME, "checkout-bar-item-title", 15)
        
        for i in range(5):
            try:
                status_txt = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "checkout-bar-item-title"))).text.strip().lower()
                if "analisar - ciclo de alçada solicitante" in status_txt:
                    status_correto = True
                    break
            except: pass
            time.sleep(1)
        
        if not status_correto:
            driver.refresh()
            time.sleep(1.5)
            for i in range(5):
                try:
                    status_txt = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "checkout-bar-item-title"))).text.strip().lower()
                    if "analisar - ciclo de alçada solicitante" in status_txt:
                        status_correto = True
                        break
                except: pass
                time.sleep(1)
        
        if not status_correto:
            atualizar_log_frontend(f"[FALHA] Verificar pedido e medição do id #{id_v360}", "error")
            contadores[1] += 1  # erro
            salvar_erro_txt(id_v360, "Status não atualizou")
            atualizar_erro_frontend(id_v360, "Status não atualizou")
            salvar_metrica(id_v360, "erro")
            atualizar_metricas_frontend(contadores[0], contadores[1], contadores[2], cont_total)
            return False
        
        # marcar como feito no kora-medicoes.web.app
        driver.switch_to.window(kora_handle)
        campo_pesquisa = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='Pesquisar...']")))
        campo_pesquisa.clear()
        campo_pesquisa.send_keys(id_v360)
        time.sleep(1)
        
        btn_feito = wait.until(EC.visibility_of_element_located((By.XPATH, "(//button[text()='FEITO'])[1]")))
        time.sleep(1)
        btn_feito.click()
        time.sleep(0.5)
        
        atualizar_log_frontend(f"✅ ID {id_v360} LIBERADO!", "success")
        atualizar_sucesso_frontend(id_v360, num_pedido)
        contadores[0] += 1  # sucesso
        salvar_backup(id_v360, num_pedido)
        salvar_metrica(id_v360, "sucesso")        
        atualizar_metricas_frontend(contadores[0], contadores[1], contadores[2], cont_total)
        return True

    try:
        atualizar_log_frontend("Iniciando Logins...")
        
        # abrir site do kora-medicoes e logar
        driver.get("https://kora-medicoes.web.app/")
        driver.execute_script("document.body.style.zoom='90%'")
        
        senha_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='password']")))
        senha_input.click()
        senha_input.clear()
        time.sleep(0.5)
        senha_input.send_keys(creds["KORA_MED_PASS"])
        time.sleep(1.5)        
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Entrar no Painel')]"))).click()
        time.sleep(1)
        kora_handle = driver.current_window_handle
        
        # abrir site do v360 em outra aba e logar
        driver.execute_script("window.open('https://kora.virtual360.io/', '_blank');")
        time.sleep(1)
        for handle in driver.window_handles:
            if handle != kora_handle:
                v360_handle = handle
                break
        driver.switch_to.window(v360_handle)
        wait.until(EC.presence_of_element_located((By.ID, "user_login"))).send_keys(creds["V360_USER"])
        wait.until(EC.presence_of_element_located((By.ID, "user_password"))).send_keys(creds["V360_PASS"])
        
        clicar_com_retry("button.v-btn.submit-button", By.CSS_SELECTOR, "Login V360")
        

        for idx, id_v360 in enumerate(ids_processar, 1):
            # Atualiza progresso
            atualizar_progresso_frontend(idx, len(ids_processar))
            
            sap_handle = None
            erro_ja_registrado = False  # 999999: controle de erro duplicado
            try:
                atualizar_log_frontend(f"Processando ID: {id_v360}...")
                
                # verificar se a medição tem pedido pronto no kora-medicoes.web.app, caso tenha, vai liberar no v360, caso não tenha, vai criar o pedido no SAP (avulso) e depois liberar no v360
                driver.switch_to.window(kora_handle)
                campo_pesquisa = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='Pesquisar...']")))
                campo_pesquisa.clear()
                campo_pesquisa.send_keys(id_v360)
                time.sleep(1)
                
                # tenta encontrar pedido no kora-medicoes.web.app (formato 45xxxxxx)
                try:
                    num_pedido_el = driver.find_element(By.XPATH, "//button[contains(text(), '45') and contains(@style, 'rgb(19, 62, 81)')]")
                    num_pedido = num_pedido_el.text.strip()
                    is_guarda_chuva = num_pedido.startswith("45")  # verifica se começa com 45 para confirmar que é um pedido, e não outra informação que esteja no botão por acaso
                except:
                    try:
                        num_pedido = driver.find_element(By.CSS_SELECTOR, "button[style*='background-color: rgb(19, 62, 81)']").text.strip()
                        is_guarda_chuva = num_pedido.startswith("45")  # verifica se começa com 45 para confirmar que é um pedido, e não outra informação que esteja no botão por acaso
                    except:
                        is_guarda_chuva = False
                if is_guarda_chuva:
                    atualizar_log_frontend(f"ID {id_v360} → GUARDA-CHUVA (Pedido: {num_pedido})")
                    processar_guarda_chuva(id_v360, kora_handle, v360_handle)
                    continue
                
                # ---- LÓGICA NORMAL (SAP + V360) ----
                atualizar_log_frontend(f"Não possui contrato no GED. Fazendo avulso.")
                
                driver.execute_script(f"window.open('https://prd.sap.korasaude.app.br/sap/bc/ui2/flp?sap-client=400&sap-language=PT#PurchaseOrder-create?sap-ui-tech-hint=GUI&uitype=advanced', '_blank');")
                time.sleep(1)
                for handle in driver.window_handles:
                    if handle != v360_handle and handle != kora_handle: sap_handle = handle
                driver.switch_to.window(sap_handle)
                
                try:
                    user_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "USERNAME_FIELD-inner")))
                    user_field.send_keys(creds["SAP_USER"])
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "PASSWORD_FIELD-inner"))).send_keys(creds["SAP_PASS"])
                    clicar_com_retry("LOGIN_LINK", By.ID, "Login SAP")
                except: pass
                
                driver.switch_to.window(v360_handle)
                driver.get(f"https://kora.virtual360.io/nf/acceptance_terms/{id_v360}")
                
                # ver se está na etapa certa do setor de contratos, caso esteja, faça, caso contrário, registre erro e vá para o próximo id
                try:
                    titulo_etapa = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "checkout-bar-item-title"))).text
                except:
                    contadores[1] += 1  # erro
                    salvar_erro_txt(id_v360, "Página do V360 não carregou a tempo")
                    atualizar_erro_frontend(id_v360, "Página do V360 não carregou a tempo")
                    salvar_metrica(id_v360, "erro")
                    atualizar_log_frontend(f"Aviso: ID {id_v360} - Página não carregou.", "warning")
                    atualizar_metricas_frontend(contadores[0], contadores[1], contadores[2], cont_total)
                    continue
                
                if "Analisar - Divergência Entre Pedido de Compras e Medição" not in titulo_etapa:
                    contadores[1] += 1  # erro
                    salvar_erro_txt(id_v360, "ID não está na etapa de criar pedido do zero.")
                    atualizar_erro_frontend(id_v360, "ID não está na etapa de criar pedido do zero.")
                    salvar_metrica(id_v360, "erro")
                    atualizar_log_frontend(f"Aviso: ID {id_v360} não está na etapa de criar pedido do zero no v360.", "warning")
                    atualizar_metricas_frontend(contadores[0], contadores[1], contadores[2], cont_total)
                    continue

                v_solicitacao = wait.until(EC.presence_of_element_located((By.ID, "acceptance_term_purchase_order"))).get_attribute("value")
                                    # Remove .0 do final da solicitação (ex: 123456.0 → 123456)
                if v_solicitacao.endswith(".0"):
                    v_solicitacao = v_solicitacao[:-2]
                v_cnpj = driver.find_element(By.ID, "acceptance_term_supplier_identification_number").get_attribute("value")
                v_valor = driver.find_element(By.ID, "acceptance_term_total_value").get_attribute("value")
                v_org_cod = driver.find_element(By.ID, "acceptance_term_erp_purchasing_organization").get_attribute("value")[:4]
                
                # verifica se o campo de solicitação tem realmente uma solicitação e não um pedido aleatório preenchido pelo solicitante, caso tenha um numero de 7 ou mais caracteres, classifica como pedido e avisa no painel que tem um pedido no lugar da solicitação, para que o usuário verifique a situação do pedido.
                if len(v_solicitacao) >= 7:
                    atualizar_log_frontend(f"Aviso: {id_v360} possui pedido no lugar da solicitação. Verifique a situação do pedido.", "warning")
                    contadores[1] += 1  # erro
                    salvar_erro_txt(id_v360, f"Possui pedido ({v_solicitacao}) no lugar da solicitação. Verifique manualmente.")
                    atualizar_erro_frontend(id_v360, f"Possui pedido ({v_solicitacao}) no lugar da solicitação.")
                    salvar_metrica(id_v360, "erro")
                    atualizar_metricas_frontend(contadores[0], contadores[1], contadores[2], cont_total)
                    continue
                    
                # se a medição for da Kora (1400), avisa que é da Kora pois o processo para fazer pedido da Kora é diferente e no momento deve ser feito manual.
                if v_org_cod == "1400":
                    atualizar_log_frontend(f"Aviso: {id_v360} é da Kora, fazer manual.", "warning")
                    continue
                    
                v_iva = "ZZ" 
                csv_path = resource_path(f"{v_org_cod}.csv")
                if os.path.exists(csv_path):
                    with open(csv_path, mode='r', encoding='utf-8') as f_csv:
                        reader = csv.reader(f_csv)
                        for row in reader:
                            if row and row[0].strip() == v_cnpj.strip():
                                v_iva = row[1].strip()
                                break
                
                driver.switch_to.window(sap_handle)
                time.sleep(1)
                focar_sap()
                
                # Aguarda elemento do SAP carregar (por ID)
                try:
                    WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "span[id*='78-text']"))
                    )
                except:
                    pass
                time.sleep(1)
                
                # ativar sintese sap com retry
                if not clicar_com_retry("div[lsdata*='Ativar síntese de documentos']", By.CSS_SELECTOR, "Ativar síntese"):
                    atualizar_log_frontend("Botão Ativar síntese não funcionou", "warning")
                    continue
                
                # Verifica se o dropdown apareceu
                try:
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "[title*='Variante de seleção']")))
                except:
                    atualizar_log_frontend("Dropdown da síntese não apareceu", "warning")
                    continue

                # abrir dropdown
                try:
                    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "[title*='Variante de seleção']")))
                    time.sleep(0.5)
                except:
                    pass
                if not clicar_com_retry("[title*='Variante de seleção']", By.CSS_SELECTOR, "Variante de seleção"): continue
                time.sleep(0.5)
                
                # clicar em requisições de compra
                try:
                    wait.until(EC.visibility_of_element_located((By.XPATH, "//tr[contains(@aria-label, 'Requisições de compra')]")))
                    time.sleep(0.5)
                except:
                    pass
                if not clicar_com_retry("//tr[contains(@aria-label, 'Requisições de compra')]", By.XPATH, "Requisições de compra"): continue
                
                # selecionar campo de solicitação e inserir número
                campo_acomp = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[title*='acompanhamento']")))
                if not clicar_com_retry("input[title*='acompanhamento']", By.CSS_SELECTOR, "Campo acompanhamento"): continue
                
                campo_acomp.clear()
                campo_acomp.send_keys(v_solicitacao)
                campo_acomp.send_keys(Keys.F8)

                # Espera INTELIGENTE: ou aparece o erro OU aparece a caixinha
                try:
                    # Espera até 15s por QUALQUER um dos dois
                    WebDriverWait(driver, 15).until(
                        lambda d: (
                            d.find_elements(By.ID, "M1:46:::0:5-text") or 
                            d.find_elements(By.CLASS_NAME, "urST5SCMetricInner")
                        )
                    )
                    
                    # Verifica se foi erro
                    try:
                        msg_inexistente = driver.find_element(By.ID, "M1:46:::0:5-text").text
                        if "Não existem dados para os critérios de seleção" in msg_inexistente:
                            atualizar_log_frontend(f"Aviso: ID {id_v360} não foi encontrado na base SAP", "warning")
                            enviar_ao_solicitante(id_v360, "inexistente_sap")
                            continue
                    except:
                        pass  # Não era erro, era a caixinha!
                    
                except:
                    atualizar_log_frontend(f"Timeout: SAP não respondeu para ID {id_v360}", "warning")
                    continue

                # Se chegou aqui, a caixinha apareceu - continua normalmente
                wait.until(EC.presence_of_element_located((By.CLASS_NAME, "urST5SCMetricInner")))
                driver.execute_script("""
                    const metric = document.querySelector('.urST5SCMetricInner');
                    if (metric) {
                        const cell = metric.closest('td');
                        ['mousedown', 'mouseup', 'click'].forEach(evtType => {
                            cell.dispatchEvent(new MouseEvent(evtType, { bubbles: true }));
                        });
                    }
                """)
                time.sleep(1.5)
                
                
                # clicar para espelhar a requisição
                if not clicar_com_retry("[title='Transferir']", By.CSS_SELECTOR, "Botão Transferir"): continue

                # Espera INTELIGENTE: sai assim que o campo do hospital aparecer COM VALOR
                try:
                    WebDriverWait(driver, 15).until(
                        lambda d: (
                            d.find_elements(By.ID, "M0:46:1:3:2:1:1[1,16]_c") and
                            json.loads(d.find_element(By.ID, "M0:46:1:3:2:1:1[1,16]_c").get_attribute("lsdata")).get("21", {}).get("value", "") != ""
                        )
                    )
                except:
                    atualizar_log_frontend("Timeout: campo hospital não carregou após Transferir", "error")
                    continue
                
                 # tratamento de erro para verificar se a solicitação espelhou corretamente (hospital correto e como serviço), caso tenha espelhado corretamente, segue o processo normalmente, caso contrário, manda automaticamente para o solicitante no v360.
                # parte especifica de verificar se o hospital está correto, de acordo com a unidade da medição
                try:
                    # PRIMEIRO: Verifica Unidade (hospital)
                    wait.until(EC.presence_of_element_located((By.ID, "M0:46:1:3:2:1:1[1,16]_c")))
                    time.sleep(0.5)
                    json_data = driver.execute_script('return document.getElementById("M0:46:1:3:2:1:1[1,16]_c").getAttribute("lsdata");')
                    nome_campo = json.loads(json_data).get("21", {}).get("value", "")
                    
                    achou_de_para = False
                    valor_esperado = ""
                    if os.path.exists(resource_path("deParaUnidades.csv")):
                        with open(resource_path("deParaUnidades.csv"), mode='r', encoding='utf-8') as f:
                            for linha in f:
                                partes = linha.strip().split(',')
                                if len(partes) >= 2 and partes[0].strip() == v_org_cod:
                                    valor_esperado = partes[1].strip()
                                    achou_de_para = True
                                    break
                                    
                    if achou_de_para and valor_esperado.upper() == nome_campo.upper():
                        atualizar_log_frontend(f"Unidade correta? Sim ({nome_campo})")
                    else:
                        atualizar_log_frontend(f"Unidade correta? Não ({nome_campo})")
                        enviar_ao_solicitante(id_v360, "de_para_errado", v_org_cod)
                        continue
                except Exception as e:
                    atualizar_log_frontend(f"Erro na leitura da Unidade: {e}", "error")
                    continue

                # parte especifica de verificar se é serviço e não material
                try:
                    wait.until(EC.presence_of_element_located((By.ID, "M0:46:1:3:2:1:1[1,15]_c")))
                    time.sleep(0.5)
                    json_data_cat = driver.execute_script('return document.getElementById("M0:46:1:3:2:1:1[1,15]_c").getAttribute("lsdata");')
                    val_cat = json.loads(json_data_cat).get("21", {}).get("value", "")
                    
                    prefixos_validos = ("REPASSE", "SERV.", "PLANO DE SAUDE", "DESPESAS COM SOFTW")
                    if val_cat.startswith(prefixos_validos):
                        atualizar_log_frontend(f"Serviço correto? Sim ({val_cat})")
                    else:
                        atualizar_log_frontend(f"Serviço correto? Não ({val_cat})")
                        enviar_ao_solicitante(id_v360, "de_para_errado", v_org_cod)
                        continue
                except Exception as e:
                    atualizar_log_frontend(f"Erro na leitura do Serviço: {e}", "error")
                    continue
                
                # clicar no campo do fornecedor
                if not clicar_com_retry("input[title*='Fornecedor/centro fornecedor']", By.CSS_SELECTOR, "Campo fornecedor"): continue
                
                # clicar no botão para abrir a telinha de digitar o cnpj
                try:
                    wait.until(EC.visibility_of_element_located((By.ID, "ls-inputfieldhelpbutton")))
                    time.sleep(0.5)
                except:
                    pass
                if not clicar_com_retry("ls-inputfieldhelpbutton", By.ID, "Botão Help Fornecedor"): continue
                
                campo_token = wait.until(EC.element_to_be_clickable((By.NAME, "TokenizerComboBox")))
                campo_token.send_keys(v_cnpj)
                campo_token.send_keys(Keys.ENTER)
                
                # clicar no botão inicio para pesquisar o fornecedor
                try:
                    wait.until(EC.visibility_of_element_located((By.ID, "btnGO2")))
                    time.sleep(0.5)
                except:
                    pass
                if not clicar_com_retry("btnGO2", By.ID, "Botão GO"): continue
                time.sleep(1.5)
                
                # tratamento de erro para verificar se o fornecedor está cadastrado e expandido para a unidade da medição. caso esteja, segue normalmente, caso não esteja, envia para o solicitante no v360.
                try:
                    msg_erro_fornecedor = driver.find_element(By.ID, "wnd[0]/sbar_msg-txt")
                    texto_erro = msg_erro_fornecedor.get_attribute("title") or msg_erro_fornecedor.text
                    if "Nenhum valor para esta seleção" in texto_erro:
                        enviar_ao_solicitante(id_v360, "cnpj_sem_cadastro")
                        continue
                    if "não foi criado para organização de compras" in texto_erro or "não foi criado para a organização de compras" in texto_erro:
                        enviar_ao_solicitante(id_v360, "cnpj_sem_expansao")
                        continue
                except:
                    texto_body = driver.find_element(By.TAG_NAME, "body").text
                    if "Nenhum valor para esta seleção" in texto_body:
                        enviar_ao_solicitante(id_v360, "cnpj_sem_cadastro")
                        continue
                    if "não foi criado para a organização de compras" in texto_body:
                        enviar_ao_solicitante(id_v360, "cnpj_sem_expansao")
                        continue
                
                # clicar no botão ir para confirmar a seleção do fornecedor
                try:
                    wait.until(EC.visibility_of_element_located((By.ID, "NSH2_copy")))
                    time.sleep(0.5)
                except:
                    pass
                if not clicar_com_retry("NSH2_copy", By.ID, "Botão Copy Fornecedor"): continue
                time.sleep(1.5)
                
                # ultimo enter para confirmar a seleção do fornecedor depois da tela ser fechada
                try:
                    wait.until(EC.invisibility_of_element_located((By.ID, "NSH2_copy")))
                    time.sleep(0.5)
                except:
                    time.sleep(0.5)
                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ENTER)
                time.sleep(2)

                # tratamento de erro adicional para se certificar de que o fornecedor está expandido corretamente para a unidade da medição
                fornecedor_com_erro = driver.execute_script("""
                    const element = document.querySelector('.lsMessageBar__text');
                    if (element) {
                        const mensagemOriginal = element.textContent || element.innerText;
                        const padrao = /O fornecedor \\d+ não foi criado para organização de compras \\d+/;
                        return padrao.test(mensagemOriginal);
                    }
                    return false;
                """)
                if fornecedor_com_erro:
                    enviar_ao_solicitante(id_v360, "cnpj_sem_expansao")
                    continue

                # caso não esteja aberta, abrir a primeira aba dentro do SAP onde tem remessa/fatura para colocar a condição de pagamento, usando CTRL+F2 
                campo_moeda_element = driver.find_elements(By.CSS_SELECTOR, "input[title*='Código da moeda']")
                if not campo_moeda_element:
                    ActionChains(driver).key_down(Keys.CONTROL).send_keys(Keys.F2).key_up(Keys.CONTROL).perform()
                    time.sleep(1)
                    
                    # clicar em remessa/fatura com retry até aparecer o ZTERM
                    for tentativa_remessa in range(5):
                        clicar_com_retry("//span[contains(text(), 'Remessa/fatura')]", By.XPATH, "Aba Remessa/fatura")
                        time.sleep(0.5)
                        # Verifica se ZTERM apareceu
                        try:
                            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[lsdata*='MEPO1226-ZTERM']")), timeout=10)
                            break
                        except:
                            if tentativa_remessa == 4:
                                atualizar_log_frontend("Aba Remessa/fatura não abriu corretamente", "warning")
                                continue
                
                # aguardar o campo de moeda e verificar se está preenchido BRL
                if campo_moeda_element:
                    campo_moeda = campo_moeda_element[0]
                else:
                    campo_moeda = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[title*='Código da moeda']")))
                time.sleep(0.5)
                moeda_valor = campo_moeda.get_attribute("value")
                if not moeda_valor or moeda_valor.strip() == "":
                    campo_moeda.clear()
                    campo_moeda.send_keys("BRL")
                    time.sleep(0.5)
                
                # colocar z030 no campo condição de pagamento (para pedidos avulsos)
                z_field = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[lsdata*='MEPO1226-ZTERM']")))
                if not clicar_com_retry("input[lsdata*='MEPO1226-ZTERM']", By.CSS_SELECTOR, "Campo ZTERM"): continue
                z_field.send_keys(Keys.CONTROL + "a", Keys.BACKSPACE)
                z_field.send_keys("Z030", Keys.ENTER)
                time.sleep(1.5)
                
                # abrir a "ultima" aba dentro do SAP para colocar a o valor do pedido, quantidade 1 e o codigo IVA, usando CTRL+4
                for tentativa_condicoes in range(10):
                    ActionChains(driver).key_down(Keys.CONTROL).send_keys("4").key_up(Keys.CONTROL).perform()
                    time.sleep(0.5)
                    
                    # clicar no botão de condições para abrir a tela de condições de pagamento e inserir o valor do pedido
                    if clicar_com_retry("//span[contains(text(), 'Condições')]", By.XPATH, "Aba Condições"):
                        break  # Achou → sai do loop
                else:
                    # Não achou após 10 tentativas de CTRL+4
                    atualizar_log_frontend("Aba Condições não apareceu após 10 tentativas", "warning")
                    continue  # Pula pro próximo ID
                time.sleep(1)
                
                # colocar o valor
                try:
                    wait.until(EC.visibility_of_element_located((By.ID, "M0:46:1:4:2:1:2:1:2B264:1:3[1,4]_c")))
                    time.sleep(0.5)
                except:
                    pass
                if not clicar_com_retry("M0:46:1:4:2:1:2:1:2B264:1:3[1,4]_c", By.ID, "Valor Condições", "M0:46:1:4:2:1:2:1:2B264:1:3[1,4]_c"): continue
                driver.execute_script(f'document.getElementById("M0:46:1:4:2:1:2:1:2B264:1:3[1,4]_c").focus(); document.execCommand("insertText", false, "{v_valor}");')
                time.sleep(1)
                
                # clicar no botão de quantidade para inserir a quantidade 1
                try:
                    wait.until(EC.visibility_of_element_located((By.ID, "M0:46:1:4:2:1:2:1::0:4-text")))
                    time.sleep(0.5)
                except:
                    pass
                if not clicar_com_retry("M0:46:1:4:2:1:2:1::0:4-text", By.ID, "Quantidade", "M0:46:1:4:2:1:2:1:2B260:1::0:16"): continue
                time.sleep(1)
                driver.execute_script("document.getElementById('M0:46:1:4:2:1:2:1:2B260:1::0:16').value = '1';")
                time.sleep(0.8)
                
                # clicar no botão de fatura para inserir o IVA
                try:
                    wait.until(EC.visibility_of_element_located((By.ID, "M0:46:1:4:2:1:2:1::0:7-text")))
                    time.sleep(0.5)
                except:
                    pass
                if not clicar_com_retry("M0:46:1:4:2:1:2:1::0:7-text", By.ID, "Fatura/IVA", "M0:46:1:4:2:1:2:1:2B263::0:68"): continue
                time.sleep(0.5)
                
                # colocar o código do iva com base nos csv salvos localmente na mesma pasta do robô
                try:
                    wait.until(EC.visibility_of_element_located((By.ID, "M0:46:1:4:2:1:2:1:2B263::0:68")))
                    time.sleep(0.5)
                except:
                    pass
                iva_field = wait.until(EC.element_to_be_clickable((By.ID, "M0:46:1:4:2:1:2:1:2B263::0:68")))
                if not clicar_com_retry("M0:46:1:4:2:1:2:1:2B263::0:68", By.ID, "Campo IVA"): continue
                iva_field.send_keys(Keys.CONTROL + "a", Keys.BACKSPACE)
                time.sleep(0.5) 
                driver.execute_script(f'arguments[0].focus(); document.execCommand("insertText", false, "{v_iva}");', iva_field)
                time.sleep(0.5)
                iva_field.send_keys(Keys.ENTER)

                # Espera INTELIGENTE: sai assim que o campo de data início estiver disponível
                try:
                    WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.ID, "M0:46:1:3:2:1:1[1,42]_c"))
                    )
                except:
                    atualizar_log_frontend("Timeout: campo de data não carregou após ENTER", "warning")
                time.sleep(1)
                # colocar as datas de inicio e fim no pedido (dia da criação do pedido e 30 dias depois)
                hoje = datetime.now().strftime("%d.%m.%Y")
                prox = (datetime.now() + timedelta(days=30)).strftime("%d.%m.%Y")

                # Preenche data início (campo já carregado)
                driver.execute_script(f"""
                    let d1 = document.getElementById("M0:46:1:3:2:1:1[1,42]_c");
                    if(d1) {{ d1.focus(); d1.value = ''; document.execCommand('insertText', false, '{hoje}'); d1.dispatchEvent(new Event('change', {{bubbles:true}})); }}
                """)
                time.sleep(0.5)

                # Preenche data fim (campo já carregado - mesma tela)
                driver.execute_script(f"""
                    let d2 = document.getElementById("M0:46:1:3:2:1:1[1,43]_c");
                    if(d2) {{ d2.focus(); d2.value = ''; document.execCommand('insertText', false, '{prox}'); d2.dispatchEvent(new Event('change', {{bubbles:true}})); }}
                """)
                time.sleep(1)
                
                # salvar pedido com CTRL+S
                ActionChains(driver).key_down(Keys.CONTROL).send_keys("s").key_up(Keys.CONTROL).perform()
                time.sleep(1)
                focar_sap()
                
                # Aguarda o número do pedido OU erro de gravação (10 tentativas de 1s)
                pedido_gerado = None
                for tentativa in range(10):
                    # Verifica popup de erro PRIMEIRO
                    try:
                        erro_gravacao = driver.find_element(By.XPATH, "//span[contains(text(), 'Gravar documento incorreto')]")
                        if erro_gravacao.is_displayed():
                            atualizar_log_frontend(f"ID {id_v360} com erro ao criar pedido. Verificar manual", "error")
                            contadores[1] += 1  # erro
                            salvar_erro_txt(id_v360, "Erro ao gravar documento no SAP")
                            atualizar_erro_frontend(id_v360, "Erro ao gravar documento no SAP")
                            salvar_metrica(id_v360, "erro")
                            erro_ja_registrado = True  # 999999: evita duplicar
                            atualizar_metricas_frontend(contadores[0], contadores[1], contadores[2], cont_total)
                            break
                    except:
                        pass
                    
                    # Verifica número do pedido
                    try:
                        status_el = driver.find_element(By.CSS_SELECTOR, "[id$='sbar_msg-txt']")
                        msg_status = status_el.get_attribute("title") or status_el.text
                        match = re.search(r"45\d+", msg_status)
                        if match:
                            pedido_gerado = match.group()
                            break
                    except:
                        pass
                    
                    time.sleep(1)
                
                # Se encontrou pedido, continua
                if pedido_gerado:
                    contadores[0] += 1  # sucesso
                    salvar_backup(id_v360, pedido_gerado)
                    salvar_metrica(id_v360, "sucesso")
                    atualizar_log_frontend(f"Sucesso: Pedido criado. Número: {pedido_gerado}", "success")
                    atualizar_sucesso_frontend(id_v360, pedido_gerado)
                    atualizar_metricas_frontend(contadores[0], contadores[1], contadores[2], cont_total)

                    atualizar_log_frontend("DEBUG: Voltando para V360...")

                    # voltar para o v360 para colocar o numero do pedido, categoria avulso e liberar a medição
                    driver.switch_to.window(v360_handle)
                    
                    btn_editar = wait.until(EC.element_to_be_clickable((By.ID, "nav-edit-tab")))
                    time.sleep(0.5)
                    btn_editar.click()
                    time.sleep(0.5) 

                    # preencher o pedido
                    script_final = """
                    var pedido = arguments[0];
                    var campos = document.querySelectorAll('#acceptance_term_purchase_order');
                    var campoAtivo = null;
                    for (var i = 0; i < campos.length; i++) {
                        if (campos[i].offsetWidth > 0 || campos[i].offsetHeight > 0) {
                            campoAtivo = campos[i]; break;
                        }
                    }
                    if (campoAtivo) {
                        campoAtivo.focus(); campoAtivo.value = pedido;
                        campoAtivo.dispatchEvent(new Event('input', { bubbles: true }));
                        campoAtivo.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                    """
                    driver.execute_script(script_final, pedido_gerado)
                    time.sleep(0.5); time.sleep(0.5); time.sleep(0.5)

                    # selecionar o tipo de pedido como avulso
                    try:
                        span_select2 = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "[aria-labelledby='select2-acceptance_term_items_attributes_0_cf_tipo_de_pedido-container']")))
                        time.sleep(0.5)
                    except:
                        pass
                    span_select2 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "[aria-labelledby='select2-acceptance_term_items_attributes_0_cf_tipo_de_pedido-container']")))
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", span_select2)
                    time.sleep(0.5)
                    span_select2.click()
                    time.sleep(0.5)
                    
                    opcao_texto = "Pedido Avulso"
                    wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[contains(text(), '{opcao_texto}')]"))).click()

                    # salvar medição
                    btn_salvar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@name='status_id' and contains(text(), 'Salvar')]")))
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_salvar)
                    time.sleep(0.5)
                    btn_salvar.click()

                    # botão "ir para ações pendentes" do v360
                    btn_pendentes = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.processable-pending-actions-btn")))
                    time.sleep(0.5)
                    btn_pendentes.click()

                    # botão de "tentar novamente" (libera medição)
                    btn_tentar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Tentar Novamente__IServ')]")))
                    time.sleep(0.5)
                    btn_tentar.click()

                    # verificar se foi para a alçada do solicitante (seguiu) ou se ficou na mesma etapa, com logica de atualizar a pagina algumas vezes para evitar falhas de atualização do status do v360, que é bem frequente.
                    status_correto = False
                    esperar_elemento(By.CLASS_NAME, "checkout-bar-item-title", 15)

                    for i in range(5):
                        try:
                            status_txt = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "checkout-bar-item-title"))).text.strip().lower()
                            if "analisar - ciclo de alçada solicitante" in status_txt:
                                status_correto = True
                                break
                        except: pass
                        time.sleep(1)
                    
                    if not status_correto:
                        driver.refresh()
                        time.sleep(1.5)
                        for i in range(5):
                            try:
                                status_txt = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "checkout-bar-item-title"))).text.strip().lower()
                                if "analisar - ciclo de alçada solicitante" in status_txt:
                                    status_correto = True
                                    break
                            except: pass
                            time.sleep(1)

                    if not status_correto:
                        atualizar_log_frontend(f"[FALHA] Verificar pedido e medição do id #{id_v360}", "error")
                        contadores[1] += 1  # erro
                        salvar_erro_txt(id_v360, "Status não atualizou após tentar novamente")
                        atualizar_erro_frontend(id_v360, "Status não atualizou após tentar novamente")
                        salvar_metrica(id_v360, "erro")
                        atualizar_metricas_frontend(contadores[0], contadores[1], contadores[2], cont_total)

                    # --- marcar AV e FEITO no kora-medicoes.web.app ---
                    driver.switch_to.window(kora_handle)
                    campo_pesquisa = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='Pesquisar...']")))
                    campo_pesquisa.clear()
                    campo_pesquisa.send_keys(id_v360)
                    time.sleep(0.5)
                    
                    btn_av = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'AV')]")))
                    btn_av.click()
                    time.sleep(0.5)
                    
                    btn_feito = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'FEITO')]")))
                    btn_feito.click()
                    atualizar_log_frontend(f"✅ ID {id_v360} LIBERADO!")

                else:
                    if not erro_ja_registrado:  # 999999: so registra se nao foi antes
                        atualizar_log_frontend(f"Erro não identificado: {id_v360}: Verificar manual", "error")
                        contadores[1] += 1  # erro
                        salvar_erro_txt(id_v360, "ALERTA: Pedido não localizado após tentar salvar. Verificar manualmente.")
                        atualizar_erro_frontend(id_v360, "ALERTA: Pedido não localizado após tentar salvar. Verificar manualmente.")
                        salvar_metrica(id_v360, "erro")
                        atualizar_metricas_frontend(contadores[0], contadores[1], contadores[2], cont_total)

            except Exception as e:
                import traceback
                atualizar_log_frontend(f"Falha ID {id_v360}: {str(e)[:200]}", "error")
                contadores[1] += 1  # erro
                salvar_erro_txt(id_v360, str(e)[:200])
                atualizar_erro_frontend(id_v360, str(e)[:200])
                salvar_metrica(id_v360, "erro")
                atualizar_metricas_frontend(contadores[0], contadores[1], contadores[2], cont_total)
            finally:
                if sap_handle:
                    try:
                        driver.switch_to.window(sap_handle)
                        driver.close()
                    except: pass
                driver.switch_to.window(v360_handle)
                
        atualizar_log_frontend("🎉 PROCESSO FINALIZADO!", "success")
        atualizar_progresso_frontend(len(ids_processar), len(ids_processar))
    except Exception as e:
        erro_msg = f"ERRO CRÍTICO: {str(e)}"
        atualizar_log_frontend(erro_msg, "error")


        
        # Salvar em arquivo pra não perder
        with open("erro_critico.txt", "w", encoding="utf-8") as f:
            import traceback
            f.write(traceback.format_exc())
    finally:
        driver.quit()

# --- INICIALIZAÇÃO ---
if __name__ == "__main__":
    import random
    port = random.randint(8000, 8999)
    # Configurações da janela Eel
    eel.start('index.html', 
              mode='chrome',
              size=(1400, 900),
              port=port,
              cmdline_args=['--start-maximized'])
