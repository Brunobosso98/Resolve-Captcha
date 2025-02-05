from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import pytesseract
from PIL import Image
import time
import pandas as pd
import requests
import os

# Configuração do caminho do Tesseract
TESSERACT_PATH = r'C:\Users\bruno.martins\Desktop\robo Busca NFSe\dependencias sistema\Tesseract-OCR\tesseract.exe'
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

URL_LOGIN = 'https://itapira.sigiss.com.br/itapira/contribuinte/login.php'
EXCEL_PATH = r'C:\Users\bruno.martins\Desktop\ResolveCaptcha\Senha Municipio Itapira.xlsx'
DOWNLOAD_DIR = r'C:\Users\...\Downloads_Livros_Fiscais'  # Pasta específica para downloads


def iniciar_driver():
    """Inicializa o WebDriver."""
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get(URL_LOGIN)
    time.sleep(5)
    return driver

    # Configurações para download automático
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True
    }
    
    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    return driver

def ler_dados_excel(caminho):
    """Lê os dados incluindo CCM (Cadastro Municipal)"""
    dados = pd.read_excel(caminho)
    return [
        (str(row['Usuário']), str(row['Senha']), str(row['CCM']), str(row['Mês']), str(row['Ano']))
        for _, row in dados.iterrows()
    ]

def construir_url_download(ccm, usuario, mes, ano):
    """Monta a URL para download direto do XLS"""
    return (
        f'https://itapira.sigiss.com.br/cgi-local/contribuinte/livro/'
        f'livro_fiscal_mensal_banco_prestado_xls.php?'
        f'ccm={ccm}&usuario={usuario}&mes={mes.zfill(2)}&ano={ano}'
    )

def renomear_arquivo(usuario, mes, ano):
    """Renomeia o último arquivo baixado com nome significativo"""
    time.sleep(2)  # Aguarda conclusão do download
    
    # Lista arquivos na pasta de downloads
    arquivos = sorted(
        os.listdir(DOWNLOAD_DIR),
        key=lambda x: os.path.getmtime(os.path.join(DOWNLOAD_DIR, x)),
        reverse=True
    )
    
    if arquivos:
        ultimo_arquivo = os.path.join(DOWNLOAD_DIR, arquivos[0])
        novo_nome = os.path.join(DOWNLOAD_DIR, f"Livro_{usuario}_{mes}_{ano}.xls")
        
        # Evita sobrescrita
        if os.path.exists(novo_nome):
            os.remove(novo_nome)
        
        os.rename(ultimo_arquivo, novo_nome)
        print(f"Arquivo renomeado: {novo_nome}")

def baixar_livro_fiscal(driver, ccm, usuario, mes, ano):
    """Navega para a URL de download e gerencia o arquivo"""
    try:
        url_download = construir_url_download(ccm, usuario, mes, ano)
        driver.get(url_download)
        renomear_arquivo(usuario, mes, ano)
    except Exception as e:
        print(f"Erro no download para {usuario}: {str(e)}")

def ler_dados_excel(caminho):
    """Lê o arquivo Excel e retorna uma lista de usuários e senhas."""
    dados = pd.read_excel(caminho)
    return [(str(dados['Usuário'][i]), str(dados['Senha'][i]), str(dados['Mês'][i]), str(dados['Ano'][i])) for i in range(len(dados))]


def extrair_numeros_imagem(driver):
    """Extrai números da imagem do captcha usando OCR."""
    try:
        elemento_imagem = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/div[1]/div[2]/div/div[2]/div[4]/div/div/div/span/img'))
        )
        
        driver.save_screenshot('screenshot.png')
        screenshot = Image.open('screenshot.png')
        
        location = elemento_imagem.location
        size = elemento_imagem.size
        left, top = location['x'], location['y']
        right, bottom = left + size['width'], top + size['height']
        
        imagem = screenshot.crop((left, top, right, bottom))
        imagem = imagem.convert('L')  # Converter para escala de cinza
        
        numeros = pytesseract.image_to_string(imagem, config='--psm 6 -c tessedit_char_whitelist=0123456789')
        return ''.join(filter(str.isdigit, numeros))
    except Exception as e:
        print(f"Erro na extração do captcha: {e}")
        return None


def preencher_campos(driver, usuario, senha):
    """Preenche os campos de login."""
    try:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "cnpj"))).send_keys(usuario)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "senha"))).send_keys(senha)
    except Exception as e:
        print(f"Erro ao preencher os campos: {e}")


def digitar_captcha(driver, numeros):
    """Digita o captcha no campo apropriado."""
    try:
        campo = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "confirma")))
        campo.clear()
        campo.send_keys(numeros)
    except Exception as e:
        print(f"Erro ao digitar captcha: {e}")


def tentar_login(driver, usuario, senha):
    """Executa o fluxo de login, tentando até o captcha ser resolvido."""
    while True:
        numeros = extrair_numeros_imagem(driver)
        if numeros:
            print(f"Números extraídos: {numeros}")
            preencher_campos(driver, usuario, senha)
            digitar_captcha(driver, numeros)
            
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "btnOk"))).click()
            
            WebDriverWait(driver, 5).until(EC.url_changes(URL_LOGIN))
            if driver.current_url == "https://itapira.sigiss.com.br/itapira/contribuinte/main.php":
                print("Login bem-sucedido!")
                return True
            else:
                print("Falha no login. Tentando novamente...")
                driver.refresh()
                time.sleep(5)
        else:
            print("Captcha não reconhecido. Tentando novamente...")
            driver.refresh()
            time.sleep(5)


def selecionar_filtros(driver, mes, ano):
    """Seleciona os filtros de mês e ano no sistema."""
    try:
        botao_modificar = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "btnAlterar")))
        botao_modificar.click()
        time.sleep(1)
        # Localizar o elemento do select para o mês
        select_mes = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@name="mes"]'))
        )
        # Criar um objeto Select e escolher o mês
        select = Select(select_mes)
        select.select_by_visible_text(mes)  # Seleciona o mês

        # Localizar o elemento do select para o ano
        select_ano = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="panelFiltro"]/table/tbody/tr/td[7]/input'))
        )
        select_ano.clear()
        select_ano.send_keys(ano)  # Envia o ano

        botao_ok = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'btnOk')))
        botao_ok.click()
    except Exception as e:
        print(f"Erro ao selecionar filtros: {e}")


def baixar_xml(driver):
    """Baixa o XML da página após a seleção dos filtros."""    
    try:
        # Obter a URL atual da página
        url_xml = driver.current_url  # Usa a URL atual do driver
        response = requests.get(url_xml)
        
        if response.status_code == 200:
            with open('dados.xml', 'wb') as file:
                file.write(response.content)
            print("XML baixado com sucesso!")
        else:
            print(f"Erro ao baixar o XML: {response.status_code}")
    except Exception as e:
        print(f"Erro ao baixar o XML: {e}")


def main():
    dados = ler_dados_excel(EXCEL_PATH)
    
    for usuario, senha, ccm, mes, ano in dados:
        driver = iniciar_driver()
        try:
            driver.get(URL_LOGIN)
            
            if tentar_login(driver, usuario, senha):
                print(f"Baixando para {usuario} - {mes}/{ano}")
                baixar_livro_fiscal(driver, ccm, usuario, mes, ano)
            
        except Exception as e:
            print(f"Falha crítica no processo para {usuario}: {str(e)}")
        finally:
            driver.quit()
            time.sleep(2)  # Intervalo entre empresas

if __name__ == "__main__":
    main()
