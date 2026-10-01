import json
import os
import re
import time
import ctypes

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from servicos_prestados import (
    URL_LOGIN,
    atualizar_excel_status,
    click_element,
    extrair_numeros_imagem,
    get_resource_path,
    preencher_campo,
    digitar_captcha,
    processar_login,
)


CAMINHO_EXCEL = get_resource_path("Senha Municipio Itapira.xlsx")
MES_COMPETENCIA = "Dezembro"
ANO_COMPETENCIA = "2025"
MAX_TENTATIVAS = 8


def normalizar_cnpj(valor):
    return re.sub(r"\D", "", str(valor or ""))


def obter_monitores_windows():
    if os.name != "nt":
        return []

    monitores = []
    callback_type = ctypes.WINFUNCTYPE(
        ctypes.c_int,
        ctypes.c_ulong,
        ctypes.c_ulong,
        ctypes.POINTER(ctypes.c_long * 4),
        ctypes.c_double,
    )

    def callback(_monitor, _hdc, rect, _data):
        left, top, right, bottom = rect.contents
        monitores.append(
            {
                "left": left,
                "top": top,
                "width": right - left,
                "height": bottom - top,
            }
        )
        return 1

    ctypes.windll.user32.EnumDisplayMonitors(0, 0, callback_type(callback), 0)
    monitores.sort(key=lambda item: (item["left"], item["top"]))
    return monitores


def posicionar_janela_no_monitor_secundario(driver):
    monitores = obter_monitores_windows()
    if len(monitores) < 2:
        driver.maximize_window()
        print("Monitor secundário não encontrado. Chrome maximizado no monitor padrão.")
        return

    monitor = monitores[1]
    driver.set_window_position(monitor["left"], monitor["top"])
    driver.set_window_size(monitor["width"], monitor["height"])
    print(
        "Chrome posicionado no monitor secundário "
        f"({monitor['left']}, {monitor['top']}) "
        f"com tamanho {monitor['width']}x{monitor['height']}."
    )


def construir_pasta_download():
    pasta = os.path.join(os.getcwd(), "livroFiscalAnual", ANO_COMPETENCIA, "12")
    os.makedirs(pasta, exist_ok=True)
    return pasta


def listar_arquivos_planilha(pasta):
    arquivos = set()
    for nome in os.listdir(pasta):
        if nome.lower().endswith((".xls", ".xlsx")):
            arquivos.add(nome)
    return arquivos


def esperar_download_planilha(pasta, arquivos_antes, timeout=90):
    limite = time.time() + timeout
    while time.time() < limite:
        temporarios = [
            nome
            for nome in os.listdir(pasta)
            if nome.lower().endswith((".crdownload", ".tmp", ".part"))
        ]
        candidatos = []
        for nome in os.listdir(pasta):
            if not nome.lower().endswith((".xls", ".xlsx")):
                continue
            if nome in arquivos_antes:
                continue
            caminho = os.path.join(pasta, nome)
            candidatos.append(caminho)

        if candidatos and not temporarios:
            candidatos.sort(key=os.path.getmtime, reverse=True)
            return candidatos[0]
        time.sleep(1)
    return None


def preparar_chrome(pasta_download):
    chrome_options = Options()
    chrome_options.add_argument("--kiosk-printing")

    app_state = {
        "recentDestinations": [
            {"id": "Save as PDF", "origin": "local", "account": ""}
        ],
        "selectedDestinationId": "Save as PDF",
        "version": 2,
    }

    prefs = {
        "download.default_directory": pasta_download,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
    }
    prefs["printing.print_preview_sticky_settings.appState"] = json.dumps(app_state)
    chrome_options.add_experimental_option("prefs", prefs)
    return chrome_options


def abrir_login(driver, wait):
    driver.get(URL_LOGIN)
    time.sleep(2)
    try:
        btn_ciente = wait.until(EC.element_to_be_clickable((By.ID, "btnCiente")))
        btn_ciente.click()
        print("Botão 'Estou Ciente' clicado com sucesso!")
    except Exception:
        print("Botão 'Estou Ciente' não encontrado ou já foi fechado.")


def preencher_competencia(driver, wait, mes, ano):
    campo_modificar = wait.until(EC.element_to_be_clickable((By.ID, "btnAlterar")))
    campo_modificar.click()

    campo_mes = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="panelFiltro"]/table/tbody/tr/td[3]/select')
        )
    )
    campo_mes.send_keys(mes)
    print(f"Mês '{mes}' digitado com sucesso!")

    campo_ano = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="panelFiltro"]/table/tbody/tr/td[7]/input')
        )
    )
    campo_ano.clear()
    campo_ano.send_keys(ano)
    print(f"Ano '{ano}' digitado com sucesso!")

    botao_ok = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "btn-success")))
    botao_ok.click()
    print("Botão OK clicado com sucesso!")
    time.sleep(2)

    driver.refresh()
    print("Página recarregada após aplicar a competência.")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(3)


def abrir_menu_acessorios(wait):
    click_element(
        wait,
        (By.ID, "dropdownMenu2"),
        "Botão 'Acessórios'",
    )
    time.sleep(1)


def acessar_painel_controle(wait):
    click_element(
        wait,
        (By.XPATH, "//a[contains(@onclick, \"abre_arquivo('dmm/_menu.php');\")]"),
        "Link 'Painel de Controle'",
    )
    time.sleep(3)


def entrar_iframe_painel(driver, wait):
    driver.switch_to.default_content()
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "main")))
    print("Entrou no iframe 'main'.")
    time.sleep(2)


def abrir_livro_fiscal(wait):
    click_element(
        wait,
        (By.XPATH, "//td[contains(@onclick, \"display('tableLivro_p');\")]"),
        "Menu 'Livro Fiscal'",
    )
    time.sleep(1)


def clicar_anual_excel(wait):
    click_element(
        wait,
        (By.XPATH, "//a[contains(@onclick, 'livroAnualP_xls()')]"),
        "Link 'Anual Excel'",
    )
    print("Download do livro anual em Excel solicitado com sucesso.")


def executar_download_anual(driver, wait):
    abrir_menu_acessorios(wait)
    acessar_painel_controle(wait)
    entrar_iframe_painel(driver, wait)
    abrir_livro_fiscal(wait)
    clicar_anual_excel(wait)
    driver.switch_to.default_content()


def renomear_planilha(caminho_arquivo, cnpj):
    _, extensao = os.path.splitext(caminho_arquivo)
    destino = os.path.join(os.path.dirname(caminho_arquivo), f"{cnpj}{extensao.lower()}")
    os.replace(caminho_arquivo, destino)
    print(f"Arquivo renomeado para: {destino}")
    return destino


def processar_empresa(row, index):
    empresa = str(row.get("Empresa", "")).strip()
    usuario = str(row.get("Usuário", "")).strip()
    senha = str(row.get("Senha", "")).strip()
    cnpj = normalizar_cnpj(usuario)

    if not cnpj or not senha:
        print(f"Linha {index + 2} ignorada: CNPJ ou senha ausentes.")
        atualizar_excel_status(index, "CNPJ ou senha ausentes.")
        return

    pasta_download = construir_pasta_download()
    tentativas = 0
    login_falhou_credenciais = False

    while tentativas < MAX_TENTATIVAS:
        driver = None
        try:
            chrome_options = preparar_chrome(pasta_download)
            driver = webdriver.Chrome(options=chrome_options)
            posicionar_janela_no_monitor_secundario(driver)
            wait = WebDriverWait(driver, 20)

            abrir_login(driver, wait)

            if tentativas == 0:
                print(f"Processando linha {index + 1}: {empresa} | CNPJ: {cnpj}")
            else:
                print(f"Tentativa {tentativas + 1} para {empresa} | CNPJ: {cnpj}")

            preencher_campo(driver, "cnpj", usuario, wait)
            preencher_campo(driver, "senha", senha, wait)

            numeros = extrair_numeros_imagem(driver, wait)
            if not numeros:
                tentativas += 1
                print("Nenhum número foi detectado no captcha.")
                continue

            print(f"Números extraídos: {numeros}")
            digitar_captcha(driver, numeros, wait)
            time.sleep(10)

            if not processar_login(driver, wait):
                current_url = driver.current_url
                if "msg=C%F3digo+de+Confirma%E7%E3o+Inv%E1lido" in current_url:
                    tentativas += 1
                    print(f"Captcha inválido para {empresa}, tentando novamente...")
                    continue
                if "msg=Contribuinte+Inexistente+ou+Senha+Inv%E1lida" in current_url:
                    print(f"Login falhou por credenciais incorretas para {empresa}.")
                    atualizar_excel_status(index, "Não foi possivel realizar o login.")
                    login_falhou_credenciais = True
                    break

                print(f"Login falhou por outro motivo para {empresa}.")
                atualizar_excel_status(index, "Falha no login por outro motivo.")
                break

            preencher_competencia(driver, wait, MES_COMPETENCIA, ANO_COMPETENCIA)
            arquivos_antes = listar_arquivos_planilha(pasta_download)
            executar_download_anual(driver, wait)

            arquivo_baixado = esperar_download_planilha(pasta_download, arquivos_antes)
            if not arquivo_baixado:
                print(f"Nenhum arquivo XLS/XLSX novo foi detectado para {empresa}.")
                atualizar_excel_status(index, "Download anual nao encontrado.")
                return

            renomear_planilha(arquivo_baixado, cnpj)
            atualizar_excel_status(index, "Livro anual XLS baixado.")
            return
        except Exception as exc:
            print(f"Erro ao processar {empresa}: {type(exc).__name__}: {exc}")
            if tentativas + 1 >= MAX_TENTATIVAS:
                atualizar_excel_status(index, "Erro durante a execucao da automacao.")
            tentativas += 1
        finally:
            if driver is not None:
                try:
                    driver.quit()
                except Exception:
                    pass

    if not login_falhou_credenciais:
        print(f"Excedido o número máximo de tentativas para {empresa}.")


def main():
    try:
        df = pd.read_excel(CAMINHO_EXCEL, engine="openpyxl")
    except Exception as exc:
        print(f"Erro ao ler a planilha: {exc}")
        return

    for index, row in df.iterrows():
        processar_empresa(row, index)


if __name__ == "__main__":
    main()
