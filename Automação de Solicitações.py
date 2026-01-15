"""
Automação de Solicitações - Sistemon / WTool

Funcionalidades:
- Login automatizado
- Cancelamento em lote
- Modificação de solicitações:
    - Mudar classificação
    - Mudar grupo
    - Mudar local
    - Mudar data
    - Delegar solicitação
- Controle de status via planilha Excel

Autor: ---
Data: ---
"""

import time
import os
import json
import pandas as pd
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from webdriver_manager.chrome import ChromeDriverManager


# ==========================================================
# CONFIGURAÇÕES E UTILITÁRIOS
# ==========================================================

def carregar_configuracoes():
    """Carrega o arquivo config.json com credenciais e parâmetros do script."""
    if not os.path.exists("config.json"):
        print("❌ ERRO: Arquivo config.json não encontrado.")
        return None

    with open("config.json", "r", encoding="utf-8") as f:
        print("✅ Configurações carregadas com sucesso.")
        return json.load(f)


def clique_robusto(driver, elemento):
    """Clique via JavaScript para evitar interceptação de elementos."""
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elemento)
    time.sleep(0.3)
    driver.execute_script("arguments[0].click();", elemento)


def clique_real(driver, elemento):
    """Clique humano usando ActionChains."""
    ActionChains(driver).move_to_element(elemento).click().perform()


def aguardar_carregamento(driver, timeout):
    """Aguarda o spinner de carregamento desaparecer."""
    spinner = (By.CSS_SELECTOR, "div.spinner-overlay")
    try:
        WebDriverWait(driver, 3).until(EC.visibility_of_element_located(spinner))
        WebDriverWait(driver, timeout).until(EC.invisibility_of_element_located(spinner))
    except TimeoutException:
        pass


def obter_valor_mapeado(valor, tipo, mapeamento):
    """Aplica mapeamento de sinônimos definido no config.json."""
    if not valor or pd.isna(valor):
        return ""
    return mapeamento.get(tipo, {}).get(valor.lower().strip(), valor)


# ==========================================================
# LOGIN
# ==========================================================

def fazer_login(driver, login_cfg, timeouts):
    """Realiza login e seleção de contrato."""
    print("🔐 Iniciando login...")

    driver.get(login_cfg["URL_LOGIN"])

    WebDriverWait(driver, timeouts["curto"]).until(
        EC.presence_of_element_located((By.NAME, "email"))
    ).send_keys(login_cfg["EMAIL"])

    driver.find_element(By.NAME, "senha").send_keys(login_cfg["SENHA"])
    driver.find_element(By.CSS_SELECTOR, "input[type='submit']").click()

    # Seleção de contrato
    contrato_btn = WebDriverWait(driver, timeouts["longo"]).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.dropdown-toggle"))
    )
    clique_robusto(driver, contrato_btn)

    contrato = WebDriverWait(driver, timeouts["curto"]).until(
        EC.element_to_be_clickable((
            By.XPATH,
            f"//span[normalize-space()='{login_cfg['CONTRATO']}']"
        ))
    )
    clique_robusto(driver, contrato)

    driver.find_element(By.CSS_SELECTOR, "input[type='submit']").click()

    WebDriverWait(driver, timeouts["longo"]).until(
        EC.presence_of_element_located((By.ID, "informacoesUsuario"))
    )

    print("✅ Login realizado com sucesso.")


# ==========================================================
# DROPDOWN BOOTSTRAP
# ==========================================================

def selecionar_opcao_bootstrap(driver, seletor_botao, texto):
    """Seleciona opções em dropdown Bootstrap usando busca + ENTER."""
    if not texto:
        return

    botao = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(seletor_botao)
    )
    clique_robusto(driver, botao)

    busca = WebDriverWait(driver, 5).until(
        EC.visibility_of_element_located((
            By.XPATH,
            "//div[contains(@class,'open')]//input"
        ))
    )
    busca.clear()
    busca.send_keys(texto)
    time.sleep(0.3)
    busca.send_keys(Keys.ENTER)
    time.sleep(1)


# ==========================================================
# EXECUÇÃO PRINCIPAL
# ==========================================================

if __name__ == "__main__":

    config = carregar_configuracoes()
    if not config:
        exit()

    cfg_login = config["LOGIN_CREDENCIAS"]
    cfg_script = config["CONFIGURACAO_SCRIPT"]
    cfg_mapeamento = config.get("MAPEAMENTO_VALORES", {})
    cfg_gerais = config.get("CONFIGURACOES_GERAIS", {})

    timeouts = {
        "curto": cfg_gerais.get("TIMEOUT_CURTO", 15),
        "longo": cfg_gerais.get("TIMEOUT_LONGO", 40)
    }

    df = pd.read_excel(cfg_script["CAMINHO_PLANILHA_COMANDOS"], dtype=str).fillna("")
    if "Status" not in df.columns:
        df["Status"] = ""

    pendentes = df[
        (~df["Status"].str.lower().isin(["concluída", "concluída (valor já correto)"])) &
        (df["ID"].str.strip() != "")
    ].copy()

    print(f"📊 Total de solicitações a processar: {len(pendentes)}")

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--window-size=1920,1080")

    driver = webdriver.Chrome(
        service=ChromeService(ChromeDriverManager().install()),
        options=chrome_options
    )

    try:
        fazer_login(driver, cfg_login, timeouts)

        # ==================================================
        # LOOP DE PROCESSAMENTO
        # ==================================================
        for idx, linha in pendentes.iterrows():
            solicitacao_id = linha["ID"].split(".")[0]
            acao = linha["Ação"].lower().strip()

            print(f"\n➡️ Processando ID {solicitacao_id} | Ação: {acao}")

            driver.get(
                f"https://www.wtool.eng.br/sistemonWeb/solicitacao/verSolicitacao/{solicitacao_id}"
            )

            try:
                botao_editar = WebDriverWait(driver, timeouts["curto"]).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//button[contains(.,'Editar/Delegar')]")
                    )
                )
                clique_robusto(driver, botao_editar)

                modal = WebDriverWait(driver, timeouts["curto"]).until(
                    EC.visibility_of_element_located((By.ID, "editarSolicitacao"))
                )

                # =============================
                # AÇÕES
                # =============================

                if acao == "delegar":
                    agente = obter_valor_mapeado(linha["Mod1"], "AGENTE", cfg_mapeamento)
                    selecionar_opcao_bootstrap(
                        driver,
                        (By.CSS_SELECTOR, "button[data-id='agent_nameEditar']"),
                        agente
                    )

                elif acao == "mudar class":
                    valor = obter_valor_mapeado(linha["Mod1"], "CLASSIFICACAO", cfg_mapeamento)
                    selecionar_opcao_bootstrap(
                        driver,
                        (By.CSS_SELECTOR, "button[data-id^='editar_preenchimentoPadrao']"),
                        valor
                    )

                # =============================
                # SALVAR
                # =============================

                botao_salvar = modal.find_element(
                    By.XPATH,
                    ".//button[normalize-space()='Salvar']"
                )
                clique_real(driver, botao_salvar)

                WebDriverWait(driver, timeouts["curto"]).until(
                    EC.alert_is_present()
                ).accept()

                WebDriverWait(driver, timeouts["longo"]).until(
                    EC.visibility_of_element_located(
                        (By.XPATH, "//div[contains(.,'sucesso')]")
                    )
                )

                df.loc[idx, "Status"] = "Concluída"
                print("✅ Ação concluída com sucesso.")

            except Exception as e:
                df.loc[idx, "Status"] = "Falha"
                print(f"❌ Erro ao processar ID {solicitacao_id}: {e}")

    finally:
        df.to_excel(cfg_script["CAMINHO_PLANILHA_COMANDOS"], index=False)
        driver.quit()
        print("\n📁 Planilha atualizada e navegador fechado.")
