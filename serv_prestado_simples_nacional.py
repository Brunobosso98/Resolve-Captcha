import time
import os

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

import exportar_notas_prestadas as export_mod
import encerrar_mes as encerrar_mod


def definir_competencia(driver, wait, mes, ano):
    """Define competência uma única vez após o login."""
    btn_alterar = wait.until(EC.element_to_be_clickable((By.ID, "btnAlterar")))
    btn_alterar.click()

    campo_mes = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="panelFiltro"]/table/tbody/tr/td[3]/select')
        )
    )
    campo_mes.send_keys(mes)
    print(f"Mês '{mes}' digitado com sucesso!")

    campo_ano = wait.until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="panelFiltro"]/table/tbody/tr/td[7]/input'))
    )
    campo_ano.clear()
    campo_ano.send_keys(ano)
    print(f"Ano '{ano}' digitado com sucesso!")

    botao_ok = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "btn-success")))
    botao_ok.click()
    print("Botão OK clicado com sucesso!")


def executar_fluxo_simples(driver, wait, mes, ano, empresa, linha_index=None):
    """Define competência, exporta notas e encerra o mês no mesmo login."""
    definir_competencia(driver, wait, mes, ano)
    export_mod.exportar_notas_prestadas(driver, wait, mes, ano, empresa, linha_index)
    encerrar_mod.clicar_encerramento_fiscal_basico(
        driver, wait, mes, ano, empresa, linha_index
    )


def main():
    try:
        df = pd.read_excel(export_mod.CAMINHO_EXCEL, engine="openpyxl")
        for index, row in df.iterrows():
            tentativas = 0
            max_tentativas = 8
            login_bem_sucedido = False
            login_falhou_credenciais = False

            while tentativas < max_tentativas and not login_bem_sucedido:
                chrome_options = Options()
                pasta_notas = export_mod.construir_pasta_notas_prestados(
                    row["Ano"], row["Mês"], row["Empresa"]
                )
                prefs = {
                    "download.default_directory": pasta_notas,
                    "download.prompt_for_download": False,
                    "download.directory_upgrade": True,
                    "plugins.always_open_pdf_externally": True,
                    "plugins.plugins_disabled": ["Chrome PDF Viewer"],
                    "plugins.plugin_field_trial_triggered": False,
                }
                chrome_options.add_experimental_option("prefs", prefs)

                driver = webdriver.Chrome(options=chrome_options)
                driver.maximize_window()
                wait = WebDriverWait(driver, 20)
                driver.get(export_mod.URL_LOGIN)
                time.sleep(2)

                try:
                    btn_ciente = wait.until(EC.element_to_be_clickable((By.ID, "btnCiente")))
                    btn_ciente.click()
                    print("Botão 'Estou Ciente' clicado com sucesso!")
                except Exception:
                    print("Botão 'Estou Ciente' não encontrado ou já foi fechado.")

                try:
                    if tentativas == 0:
                        print(f"Processando linha {index + 1}: {row['Empresa']}")
                    else:
                        print(f"Tentativa {tentativas + 1} para a empresa: {row['Empresa']}")

                    export_mod.preencher_campo(driver, "cnpj", row["Usuário"], wait)
                    export_mod.preencher_campo(driver, "senha", row["Senha"], wait)

                    numeros = export_mod.extrair_numeros_imagem(driver, wait)
                    if numeros:
                        print(f"Números extraídos: {numeros}")
                        export_mod.digitar_captcha(driver, numeros, wait)
                        time.sleep(10)

                        if export_mod.processar_login(driver, wait):
                            login_bem_sucedido = True
                            executar_fluxo_simples(
                                driver,
                                wait,
                                row["Mês"],
                                row["Ano"],
                                row["Empresa"],
                                index,
                            )
                        else:
                            current_url = driver.current_url
                            if "msg=C%F3digo+de+Confirma%E7%E3o+Inv%E1lido" in current_url:
                                print(
                                    f"Captcha inválido para {row['Empresa']}, tentando novamente..."
                                )
                                tentativas += 1
                            elif "msg=Contribuinte+Inexistente+ou+Senha+Inv%E1lida" in current_url:
                                print(
                                    f"Login falhou por credenciais incorretas para {row['Empresa']}"
                                )
                                export_mod.atualizar_excel_status(
                                    index, "Não foi possivel realizar o login."
                                )
                                login_falhou_credenciais = True
                                break
                            else:
                                print("Login falhou por outro motivo, não preenchendo a data.")
                                break
                    else:
                        print("Nenhum número foi detectado!")
                        tentativas += 1
                        if tentativas < max_tentativas:
                            print(f"Tentando novamente ({tentativas}/{max_tentativas})...")
                finally:
                    driver.quit()

            if not login_bem_sucedido and not login_falhou_credenciais:
                print(
                    f"Excedido o número máximo de tentativas para {row['Empresa']}. "
                    "Indo para a próxima empresa."
                )
    except Exception as exc:
        print(f"Erro geral: {exc}")


if __name__ == "__main__":
    main()
