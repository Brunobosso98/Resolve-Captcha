import sys
import time
import os
import json
import re
import unicodedata

import pandas as pd
import pytesseract
from PIL import Image
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import StaleElementReferenceException, UnexpectedAlertPresentException


def get_resource_path(relative_path):
    """Obtém o caminho absoluto para recursos, funciona para desenvolvimento e para PyInstaller"""
    if hasattr(sys, "_MEIPASS"):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def click_element(wait, locator, descricao, tentativas=3):
    """Clica em um elemento garantindo nova referência quando ocorrer StaleElement."""
    ultima_excecao = None
    for _ in range(tentativas):
        try:
            elemento = wait.until(EC.element_to_be_clickable(locator))
            elemento.click()
            print(f"{descricao} clicado com sucesso!")
            return elemento
        except StaleElementReferenceException as exc:
            ultima_excecao = exc
            time.sleep(0.4)
    raise ultima_excecao or Exception("Não foi possível clicar no elemento.")


CAMINHO_TESSERACT = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
CAMINHO_EXCEL = get_resource_path("notasEmitir.xlsx")
URL_LOGIN = "https://itapira.sigiss.com.br/itapira/contribuinte/login.php"

DEFAULT_LOCAL = "J1"
DEFAULT_CNPJ_TOMADOR = "10493581000102"
DEFAULT_TIMEOUT_EMISSAO = 90

pytesseract.pytesseract.tesseract_cmd = CAMINHO_TESSERACT


def extrair_numeros_imagem(driver, wait):
    numeros = None
    try:
        elemento_imagem = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//*[@id="content"]/div[1]/div[2]/div/div[2]/div[4]/div/div/div/span/img')
            )
        )

        driver.save_screenshot("screenshot.png")
        screenshot = Image.open("screenshot.png")

        location = elemento_imagem.location
        size = elemento_imagem.size
        left = location["x"]
        top = location["y"]
        right = location["x"] + size["width"]
        bottom = location["y"] + size["height"]

        imagem = screenshot.crop((left, top, right, bottom))
        imagem.save("numero.png")

        imagem = imagem.convert("L")
        numeros = pytesseract.image_to_string(
            imagem, config="--psm 6 -c tessedit_char_whitelist=0123456789"
        )
        numeros = "".join(filter(str.isdigit, numeros))

    except Exception as e:
        print(f"Erro na extração: {type(e).__name__}: {e}")

    return numeros


def preencher_campo(driver, element_id, valor, wait):
    try:
        campo = wait.until(EC.element_to_be_clickable((By.ID, element_id)))
        campo.clear()
        campo.send_keys(valor)
        print(f"Campo {element_id} preenchido")
    except Exception as e:
        print(f"Erro ao preencher {element_id}: {e}")


def processar_login(driver, wait):
    try:
        if driver.current_url == "https://itapira.sigiss.com.br/itapira/contribuinte/main.php":
            return True

        current_url = driver.current_url
        if "msg=C%F3digo+de+Confirma%E7%E3o+Inv%E1lido" in current_url:
            print("Erro de login: Código de Confirmação Inválido (captcha)")
            return False

        if "msg=Contribuinte+Inexistente+ou+Senha+Inv%E1lida" in current_url:
            print("Erro de login: Contribuinte Inexistente ou Senha Inválida")
            return False

        try:
            erro_elemento = wait.until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/div[2]/div/font/b/center'))
            )
            if erro_elemento.is_displayed():
                print("Erro no login: ", erro_elemento.text)
                if (
                    "Confirmação" in erro_elemento.text
                    or "confirmação" in erro_elemento.text
                    or "Código" in erro_elemento.text
                ):
                    return False
                return False
        except Exception:
            pass

    except Exception as e:
        print(f"Erro no processo de login: {e}")
        return False


def digitar_captcha(driver, numeros, wait):
    try:
        campo = wait.until(EC.element_to_be_clickable((By.ID, "confirma")))
        campo.clear()
        campo.send_keys(numeros)
        print("Captcha digitado com sucesso!")

        botao_logar = wait.until(EC.element_to_be_clickable((By.ID, "btnOk")))
        botao_logar.click()

    except Exception as e:
        print(f"Erro ao digitar captcha: {e}")


def atualizar_excel_status(linha_index, mensagem):
    try:
        workbook = load_workbook(CAMINHO_EXCEL)
        worksheet = workbook.active

        coluna_status = None
        for col_idx, col_name in enumerate(worksheet[1], 1):
            if col_name.value == "Status Processo":
                coluna_status = col_idx
                break

        if coluna_status is not None:
            worksheet.cell(row=linha_index + 2, column=coluna_status, value=mensagem)
        else:
            coluna_status = worksheet.max_column + 1
            worksheet.cell(row=1, column=coluna_status, value="Status Processo")
            worksheet.cell(row=linha_index + 2, column=coluna_status, value=mensagem)

        workbook.save(CAMINHO_EXCEL)
        workbook.close()
    except Exception as e:
        print(f"Erro ao atualizar o Excel: {e}")


def normalizar(texto):
    if texto is None:
        return ""
    texto = str(texto).strip()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(ch for ch in texto if not unicodedata.combining(ch)).lower()


def obter_valor_row(row, candidatos, fallback=None):
    for col in candidatos:
        if col in row and pd.notna(row[col]):
            return row[col]
    return fallback


def normalizar_cnpj(valor):
    if valor is None:
        return ""
    return re.sub(r"\D", "", str(valor))


def formatar_valor_2casas(valor):
    try:
        if isinstance(valor, (int, float)):
            return f"{float(valor):.2f}".replace(".", ",")

        texto = str(valor).strip()
        if "," in texto and "." in texto:
            texto = texto.replace(".", "").replace(",", ".")
        elif "," in texto:
            texto = texto.replace(",", ".")
        numero = float(texto)
        return f"{numero:.2f}".replace(".", ",")
    except Exception:
        return "10,00"


def digitar_lento(elemento, texto, delay=0.08):
    for ch in str(texto):
        elemento.send_keys(ch)
        time.sleep(delay)


def setar_valor_js(driver, elemento, valor):
    driver.execute_script(
        "arguments[0].value = arguments[1];"
        "arguments[0].dispatchEvent(new Event('input', {bubbles: true}));"
        "arguments[0].dispatchEvent(new Event('change', {bubbles: true}));",
        elemento,
        valor,
    )


def escolher_local(wait, valor_excel):
    select_el = wait.until(EC.element_to_be_clickable((By.ID, "local")))
    select = Select(select_el)

    valor = str(valor_excel).strip() if valor_excel is not None else ""
    valor_norm = normalizar(valor)

    mapa = {
        "nota regime especial": "PFNI",
        "pessoa fisica": "F",
        "juridica do municipio": "J1",
        "juridica de fora (nao estabelecida no municipio)": "J2",
        "exportacao (empresa de fora do pais)": "J3",
        "produtor rural do municipio": "J4",
        "produtor rural de fora do municipio": "J5",
        "candidato politico": "J6",
    }

    try:
        if valor in {"PFNI", "F", "J1", "J2", "J3", "J4", "J5", "J6"}:
            select.select_by_value(valor)
            print(f"Local selecionado por valor: {valor}")
            return
        if valor_norm in mapa:
            select.select_by_value(mapa[valor_norm])
            print(f"Local selecionado por texto: {valor}")
            return
    except Exception as e:
        print(f"Falha ao selecionar local informado ({valor}): {e}")

    select.select_by_value(DEFAULT_LOCAL)
    print("Local padrão selecionado: Jurídica do Municipio (J1)")


def double_click_element(driver, elemento, descricao):
    try:
        ActionChains(driver).move_to_element(elemento).pause(0.2).double_click(elemento).perform()
        print(f"{descricao} com duplo clique!")
        return
    except Exception:
        pass

    try:
        driver.execute_script(
            "arguments[0].dispatchEvent(new MouseEvent('dblclick', {bubbles: true}));",
            elemento,
        )
        print(f"{descricao} com duplo clique via JS!")
    except Exception as e:
        print(f"Falha no duplo clique: {e}")
        return

def chamar_doubleclick_js(driver, elemento, descricao):
    try:
        driver.execute_script(
            "if (typeof doubleClick === 'function') { doubleClick(arguments[0]); }",
            elemento,
        )
        print("{descricao} via função doubleClick()!")
        return True
    except Exception as e:
        print("Falha ao chamar doubleClick(): {e}")
        return False

def aguardar_iframe_detail_sumir(driver, timeout=10):
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.ID, "detail"))
        )
        print("Iframe 'detail' oculto.")
    except Exception:
        print("Iframe 'detail' ainda visível; tentando continuar.")

def selecionar_linha_tomador(driver, wait, linha_tomador):
    def selecionada():
        return "lineSelected" in (linha_tomador.get_attribute("class") or "")

    if selecionada():
        return True

    # Seleciona via função do próprio sistema (evita toggling)
    try:
        driver.execute_script(
            "if (typeof lineSelected === 'function') { lineSelected(arguments[0]); }",
            linha_tomador,
        )
    except Exception:
        pass

    try:
        wait.until(lambda d: selecionada())
        return True
    except Exception:
        return False

def emitir_nfse(driver, wait, row, inicio=None, timeout=30):
    inicio = inicio or time.time()

    def expirou():
        return (time.time() - inicio) > timeout

    cnpj_excel = obter_valor_row(
        row,
        ["CNPJ Tomador", "CNPJ", "CPF/CNPJ", "Documento"],
        fallback=DEFAULT_CNPJ_TOMADOR,
    )
    # Extrai números primeiro
    numeros = re.sub(r"\D", "", str(cnpj_excel)).strip()
    # Trata como número para preservar zeros à esquerda
    if pd.notna(cnpj_excel):
        try:
            cnpj_num = int(float(cnpj_excel))
            numeros = str(cnpj_num)
        except:
            pass
    # Completa com zeros à esquerda conforme tipo (CPF=11, CNPJ=14)
    if len(numeros) <= 11:
        cnpj_texto = numeros.zfill(11)  # CPF
    else:
        cnpj_texto = numeros.zfill(14)  # CNPJ
    if not cnpj_texto or cnpj_texto == "0" * len(cnpj_texto):
        cnpj_texto = DEFAULT_CNPJ_TOMADOR

    def selecionar_tomador_com_local(local_codigo):
        try:
            escolher_local(wait, local_codigo)
            time.sleep(2)
            ActionChains(driver).send_keys(
                Keys.TAB, cnpj_texto
            ).perform()
            ActionChains(driver).send_keys(
                Keys.TAB, Keys.TAB, Keys.ENTER
            ).perform()

            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "detail")))
            print("Entrou no iframe 'detail' para seleção do tomador.")
            if expirou():
                driver.switch_to.parent_frame()
                return False, "Timeout 30s"

            linha_xpath = (
                "//div[contains(@class,'bluecubeGrid')]//table//tr[not(contains(@class,'header'))]"
                "[.//td[2][contains(normalize-space(.), '" + cnpj_texto + "')]][1]"
            )
            try:
                linha_tomador = WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.XPATH, linha_xpath)))
            except Exception:
                try:
                    botao_cancelar = wait.until(EC.element_to_be_clickable((By.ID, "btnCancelar")))
                    botao_cancelar.click()
                except Exception:
                    pass
                driver.switch_to.parent_frame()
                return False, "Tomador não encontrado"

            try:
                celula_cnpj = linha_tomador.find_element(By.XPATH, "./td[2]")
            except Exception:
                celula_cnpj = linha_tomador.find_element(By.XPATH, "./td[1]")

            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", celula_cnpj)
            except Exception:
                pass

            selecionar_linha_tomador(driver, wait, linha_tomador)

            try:
                botao_ok = wait.until(EC.element_to_be_clickable((By.ID, "btnOk")))
                botao_ok.click()
                time.sleep(2)
                print("Botão Ok clicado no popup de tomador.")
            except UnexpectedAlertPresentException:
                try:
                    alert = driver.switch_to.alert
                    print(f"Alerta ao clicar Ok: {alert.text}")
                    alert.accept()
                except Exception:
                    pass
            except Exception as e:
                print(f"Falha ao clicar no botão Ok do tomador: {e}")

            driver.switch_to.parent_frame()
            print("Retornou ao iframe 'main'.")
            return True, ""
        except Exception as e:
            print(f"Falha ao selecionar o tomador na tabela: {e}")
            try:
                driver.switch_to.parent_frame()
            except Exception:
                pass
            return False, "Falha ao selecionar tomador"

    locais_tentativa = ["F"] if len(cnpj_texto) == 11 else ["J1", "J2"]
    sucesso_tomador = False
    motivo_falha_tomador = "Tomador não encontrado"

    for idx, local_codigo in enumerate(locais_tentativa):
        try:
            driver.switch_to.default_content()
            click_element(
                wait,
                (By.XPATH, "//div[contains(@class,'btn-group')]/button[@id='dropdownMenu2' and contains(normalize-space(.), 'Serviços Prestados')]"),
                "Menu 'Serviços Prestados'",
            )
            click_element(
                wait,
                (By.XPATH, "//a[contains(@onclick, 'nfe/nfe.php')]"),
                "Link 'Emissão de NFSe'",
            )
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "main")))
            print("Entrou no iframe 'main'.")
        except Exception as e:
            print(f"Falha ao abrir Emissão de NFSe: {e}")
            driver.switch_to.default_content()
            return False, "Falha ao abrir emissão NFSe"

        sucesso_tomador, motivo_falha_tomador = selecionar_tomador_com_local(local_codigo)
        if sucesso_tomador:
            break

        if (
            motivo_falha_tomador == "Tomador não encontrado"
            and idx + 1 < len(locais_tentativa)
        ):
            print(
                f"Tomador não encontrado com local {local_codigo}; "
                f"tentando novamente com {locais_tentativa[idx + 1]}."
            )
            continue

        print("Tomador não encontrado; pulando empresa.")
        return False, motivo_falha_tomador

    try:
        click_element(wait, (By.XPATH, "//button[contains(@onclick, 'openFiltro')]"), "Botão de filtro")
    except Exception:
        botao_filtro = wait.until(
            EC.presence_of_element_located((By.XPATH, "//button[contains(@onclick, 'openFiltro')]"))
        )
        driver.execute_script("arguments[0].click();", botao_filtro)
        print("Botão de filtro clicado via JS.")

    try:
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "detail")))
        print("Entrou no iframe 'detail' para seleção do serviço.")
        try:
            campo_codigo = wait.until(EC.element_to_be_clickable((By.ID, "codigo")))
            campo_codigo.clear()
            campo_codigo.send_keys("1719")
            ActionChains(driver).send_keys(Keys.TAB, Keys.TAB, Keys.ENTER).perform()
            time.sleep(1)
        except Exception:
            pass

        linha_servico = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "//div[contains(@class,'bluecubeGrid')]//tr[not(contains(@class,'header'))]"
                    "[.//td[1][contains(normalize-space(.), '1719')]][1]",
                )
            )
        )
        selecionar_linha_tomador(driver, wait, linha_servico)
        time.sleep(0.2)

        def confirmar_servico():
            try:
                driver.execute_script(
                    "if (typeof confirmSelection === 'function') { confirmSelection(); return true; } return false;"
                )
                return True
            except Exception:
                botao_ok_servico = wait.until(EC.element_to_be_clickable((By.ID, "btnOk")))
                botao_ok_servico.click()
                return True

        try:
            confirmar_servico()
        except UnexpectedAlertPresentException:
            try:
                alert = driver.switch_to.alert
                texto_alerta = alert.text
                alert.accept()
                if "Nenhum registro" in texto_alerta:
                    try:
                        driver.execute_script(
                            "arguments[0].classList.add('lineSelected');"
                            "if (typeof lineSelected === 'function') { lineSelected(arguments[0]); }",
                            linha_servico,
                        )
                    except Exception:
                        pass
                    try:
                        confirmar_servico()
                    except UnexpectedAlertPresentException:
                        try:
                            alert = driver.switch_to.alert
                            alert.accept()
                        except Exception:
                            pass
            except Exception:
                pass

        print("Serviço 1719 selecionado e OK confirmado.")
        driver.switch_to.parent_frame()
        print("Retornou ao iframe 'main'.")
    except UnexpectedAlertPresentException:
        try:
            alert = driver.switch_to.alert
            print(f"Alerta no serviço: {alert.text}")
            alert.accept()
        except Exception:
            pass
        driver.switch_to.parent_frame()
        return False, "Alerta ao selecionar serviço"
    except Exception as e:
        print(f"Falha ao selecionar o serviço 1719: {e}")
        driver.switch_to.parent_frame()
        return False, "Falha ao selecionar serviço"

    if expirou():
        return False, "Timeout 30s"

    click_element(wait, (By.ID, "btnTributos"), "Botão 'Reforma Tributária'")
    time.sleep(1)

    try:
        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "main")))
        time.sleep(1)

        ActionChains(driver).send_keys(Keys.TAB * 19).perform()
        time.sleep(2)
        ActionChains(driver).send_keys(Keys.SPACE).perform()
        ActionChains(driver).send_keys("Itapira").perform()
        time.sleep(2)
        ActionChains(driver).send_keys(Keys.ENTER).perform()
        ActionChains(driver).send_keys(Keys.TAB).perform()
        time.sleep(1)
        ActionChains(driver).send_keys(Keys.TAB * 10).perform()
        time.sleep(1)
        ActionChains(driver).send_keys(Keys.ENTER).perform()
        time.sleep(5)
    except Exception as e:
        print(f"Falha no modal da Reforma Tributária: {e}")
        driver.switch_to.default_content()
        return False, "Falha no modal Reforma Tributária"

    driver.switch_to.default_content()
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "main")))

    valor_excel = obter_valor_row(row, ["Valor", "valor", "Valor NF", "Valor Nota"], fallback=10)
    valor_texto = formatar_valor_2casas(valor_excel)

    try:
        campo_valor = wait.until(EC.element_to_be_clickable((By.ID, "valor")))
        campo_valor.clear()
        setar_valor_js(driver, campo_valor, valor_texto)
    except Exception as e:
        print(f"Falha ao preencher valor: {e}")

    try:
        campo_desc = wait.until(EC.element_to_be_clickable((By.ID, "descricaoNF")))
        campo_desc.clear()
        campo_desc.send_keys("Serviços referente ao processo de renovação do alvará municipal 2026")
        time.sleep(1)
    except Exception as e:
        print(f"Falha ao preencher descrição: {e}")

    if expirou():
        driver.switch_to.default_content()
        return False, "Timeout 30s"

    try:
        botao_emitir = wait.until(EC.element_to_be_clickable((By.ID, "btnEmitirNF")))
        botao_emitir.click()
        try:
            alerta = WebDriverWait(driver, 20).until(EC.alert_is_present())
            alerta.accept()
            print("Alerta da emissão aceito.")
        except Exception:
            print("Nenhum alerta encontrado após emitir.")
            driver.switch_to.default_content()
            return False, "Sem alerta final na emissão"
    except Exception as e:
        print(f"Falha ao emitir NFSe: {e}")
        driver.switch_to.default_content()
        return False, "Falha ao emitir NFSe"

    driver.switch_to.default_content()
    return True, ""


def abrir_emissao_nfse(driver, wait):
    driver.switch_to.default_content()
    click_element(
        wait,
        (By.XPATH, "//div[contains(@class,'btn-group')]/button[@id='dropdownMenu2' and contains(normalize-space(.), 'Serviços Prestados')]"),
        "Menu 'Serviços Prestados'",
    )
    click_element(
        wait,
        (By.XPATH, "//a[contains(@onclick, 'nfe/nfe.php')]") ,
        "Link 'Emissão de NFSe'",
    )


def main():
    try:
        cnpjs_pular = set()
        if os.path.exists("cnpj_pular.json"):
            try:
                with open("cnpj_pular.json", "r", encoding="utf-8") as arquivo:
                    dados_pular = json.load(arquivo)
                for item in dados_pular.get("cnpjs", []):
                    cnpjs_pular.add(normalizar_cnpj(item))
            except Exception as e:
                print(f"Falha ao ler cnpj_pular.json: {e}")

        if os.path.exists(CAMINHO_EXCEL):
            df = pd.read_excel(CAMINHO_EXCEL, engine="openpyxl")
        else:
            df = pd.DataFrame([{}])
            print("Arquivo notasEmitir.xlsx não encontrado; usando valores padrão.")

        if "Valor" in df.columns:
            valores = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
            df = df[valores > 0]

        if "CNPJ" in df.columns and cnpjs_pular:
            df = df[~df["CNPJ"].apply(lambda x: normalizar_cnpj(x) in cnpjs_pular)]

        tentativas = 0
        max_tentativas = 8
        login_bem_sucedido = False

        while tentativas < max_tentativas and not login_bem_sucedido:
            chrome_options = Options()
            chrome_options.add_argument("--kiosk-printing")

            prefs = {
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
            driver.get(URL_LOGIN)

            time.sleep(2)

            try:
                btn_ciente = wait.until(EC.element_to_be_clickable((By.ID, "btnCiente")))
                btn_ciente.click()
                print("Botão 'Estou Ciente' clicado com sucesso!")
            except Exception:
                print("Botão 'Estou Ciente' não encontrado ou já foi fechado.")

            preencher_campo(driver, "cnpj", "06186859000113", wait)
            preencher_campo(driver, "senha", "18978", wait)

            numeros = extrair_numeros_imagem(driver, wait)

            if numeros:
                print(f"Números extraídos: {numeros}")
                digitar_captcha(driver, numeros, wait)
                time.sleep(10)

                if processar_login(driver, wait):
                    login_bem_sucedido = True
                else:
                    current_url = driver.current_url
                    if "msg=C%F3digo+de+Confirma%E7%E3o+Inv%E1lido" in current_url:
                        print("Captcha inválido, tentando novamente...")
                        tentativas += 1
                        driver.quit()
                    elif "msg=Contribuinte+Inexistente+ou+Senha+Inv%E1lida" in current_url:
                        print("Login falhou por credenciais incorretas")
                        driver.quit()
                        return
                    else:
                        print("Login falhou por outro motivo.")
                        driver.quit()
                        return
            else:
                print("Nenhum número foi detectado!")
                tentativas += 1
                driver.quit()

        if not login_bem_sucedido:
            print("Excedido o número máximo de tentativas de login.")
            return

        for index, row in df.iterrows():
            print(f"Processando linha {index + 1}")
            try:
                sucesso, motivo = emitir_nfse(
                    driver, wait, row, inicio=time.time(), timeout=DEFAULT_TIMEOUT_EMISSAO
                )
                if not sucesso and motivo == "Timeout 30s":
                    print(
                        f"Timeout na linha {index + 1}; reiniciando emissão da mesma linha."
                    )
                    driver.switch_to.default_content()
                    try:
                        driver.refresh()
                        time.sleep(2)
                        sucesso, motivo = emitir_nfse(
                            driver, wait, row, inicio=time.time(), timeout=DEFAULT_TIMEOUT_EMISSAO
                        )
                    except Exception:
                        pass
                if not sucesso:
                    cnpj_row = obter_valor_row(row, ["CNPJ"], fallback="")
                    valor_row = obter_valor_row(row, ["Valor"], fallback="")
                    with open("falhas_emissao.txt", "a", encoding="utf-8") as arquivo:
                        arquivo.write(
                            f"Linha {int(index + 1)} | CNPJ: {cnpj_row} | Valor: {valor_row} | Motivo: {motivo or 'Falha desconhecida'}\n"
                        )
                else:
                    cnpj_row = obter_valor_row(row, ["CNPJ"], fallback="")
                    valor_row = obter_valor_row(row, ["Valor"], fallback="")
                    with open("sucesso_emissao.txt", "a", encoding="utf-8") as arquivo:
                        arquivo.write(
                            f"Linha {int(index + 1)} | CNPJ: {cnpj_row} | Valor: {valor_row}\n"
                        )
                time.sleep(40)
            except Exception as e:
                print(f"Falha ao emitir na linha {index + 1}: {e}")
                cnpj_row = obter_valor_row(row, ["CNPJ"], fallback="")
                valor_row = obter_valor_row(row, ["Valor"], fallback="")
                with open("falhas_emissao.txt", "a", encoding="utf-8") as arquivo:
                    arquivo.write(
                        f"Linha {int(index + 1)} | CNPJ: {cnpj_row} | Valor: {valor_row} | Motivo: Exceção {type(e).__name__}\n"
                    )

    except Exception as e:
        print(f"Erro geral: {e}")


if __name__ == "__main__":
    main()
