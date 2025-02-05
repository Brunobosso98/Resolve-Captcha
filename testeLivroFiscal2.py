from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import pytesseract
from PIL import Image
import time
import pandas as pd
import pdfplumber
import requests
import os
import re

# Configuração do caminho do Tesseract
TESSERACT_PATH = r'C:\Users\bruno.martins\Desktop\robo Busca NFSe\dependencias sistema\Tesseract-OCR\tesseract.exe'
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

URL_LOGIN = 'https://itapira.sigiss.com.br/itapira/contribuinte/login.php'
EXCEL_PATH = r'C:\Users\bruno.martins\Desktop\ResolveCaptcha\Senha Municipio Itapira.xlsx'
DOWNLOAD_DIR = r'C:\Users\bruno.martins\Desktop\ResolveCaptcha\livro fiscal'  # Pasta específica para downloads
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_DIR,  # Define o diretório padrão
    "download.prompt_for_download": False,  # Evita o prompt de download
    "plugins.always_open_pdf_externally": True  # Faz com que o PDF seja baixado e não aberto no navegador
})


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
    """Lê o arquivo Excel e retorna uma lista de usuários, senhas, CNPJ, mês e ano."""
    dados = pd.read_excel(caminho)
    return [(str(dados['Usuário'][i]), str(dados['Senha'][i]), str(dados['Mês'][i]), str(dados['Ano'][i])) for i in range(len(dados))]
    
def executar_funcao_download(driver):
    """Executa a função livroMensalP() e captura a URL gerada."""
    try:
        driver.execute_script("livroMensalP();")
        time.sleep(5)  # Aguarda a URL ser gerada

        # Captura a URL gerada na nova aba
        janelas = driver.window_handles
        if len(janelas) > 1:
            driver.switch_to.window(janelas[1])  # Alterna para a nova aba
            pdf_url = driver.current_url  # Captura a URL do PDF
            driver.close()  # Fecha a nova aba
            driver.switch_to.window(janelas[0])  # Retorna à aba principal
            print(f"🔗 URL Capturada: {pdf_url}")
            return pdf_url

    except Exception as e:
        print(f"Erro ao executar a função livroMensalP(): {e}")
    
    return None

def baixar_pdf(pdf_url, download_path):
    """Baixa o PDF diretamente da URL capturada."""
    try:
        response = requests.get(pdf_url, stream=True)
        if response.status_code == 200:
            with open(download_path, "wb") as file:
                file.write(response.content)
            print(f"✅ PDF baixado com sucesso: {download_path}")
            return download_path
        else:
            print(f"❌ Erro ao baixar PDF: {response.status_code}")
    except Exception as e:
        print(f"Erro ao baixar PDF: {e}")
    
    return None

def formatar_cnpj(usuario):
    """Remove pontos, barras e traços do CNPJ."""
    return re.sub(r'\D', '', usuario)

def converter_mes_para_numero(mes):
    """Converte nome do mês para número."""
    meses = {
        "Janeiro": "01", "Fevereiro": "02", "Março": "03", "Abril": "04",
        "Maio": "05", "Junho": "06", "Julho": "07", "Agosto": "08",
        "Setembro": "09", "Outubro": "10", "Novembro": "11", "Dezembro": "12"
    }
    return meses.get(mes, "00")  # Retorna "00" se o mês não for encontrado (evita erro)

def extrair_ccm(driver):
    """Extrai o CCM da empresa a partir do HTML da página após login."""
    try:
        elemento_td = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="floatingHint"]/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[1]/td'))
        )
        
        texto_completo = elemento_td.get_attribute("innerHTML")  # Obtém todo o HTML interno
        print("Conteúdo da <td>:")
        print(texto_completo)  # Exibe todo o conteúdo da tag <td> para análise

        # Expressão regular para extrair o primeiro conjunto de números
        match = re.search(r'(\d+)', texto_completo)
        if match:
            print(match)
            return match.group(1)  # Retorna apenas o número do CCM
        else:
            print("CCM não encontrado no texto extraído!")
            return None
    except Exception as e:
        print(f"Erro ao extrair o CCM: {e}")
        return None

def baixar_livro_fiscal(driver, usuario, ccm, mes, ano):
    """Acessa a URL do livro fiscal mensal e baixa o PDF."""
    try:
        if not ccm:
            print(f"CCM não encontrado para CNPJ {usuario}. Pulando empresa.")
            return

        # Formatar CNPJ e mês corretamente
        cnpj_limpo = re.sub(r'\D', '', usuario)
        meses = {
            "Janeiro": "01", "Fevereiro": "02", "Março": "03", "Abril": "04",
            "Maio": "05", "Junho": "06", "Julho": "07", "Agosto": "08",
            "Setembro": "09", "Outubro": "10", "Novembro": "11", "Dezembro": "12"
        }
        mes_numerico = meses.get(mes, "00")

        # Construindo a URL correta para PDF
        url_livro_pdf = f"https://itapira.sigiss.com.br/cgi-local/contribuinte/livro/livro_fiscal_mensal_banco_prestado_pdf.php?ccm={ccm}&cnpj={cnpj_limpo}&mes={mes_numerico}&ano={ano}"

        print(f"Abrindo URL no navegador para download: {url_livro_pdf}")

        # Selenium acessa a URL diretamente (já autenticado)
        driver.get(url_livro_pdf)

        # Espera alguns segundos para o download iniciar
        time.sleep(5)

        print("Verifique a pasta de downloads do navegador!")
    except Exception as e:
        print(f"Erro ao tentar baixar o livro fiscal: {e}")

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

    if not os.path.exists(DOWNLOAD_DIR):
        os.makedirs(DOWNLOAD_DIR)

    dados = ler_dados_excel(EXCEL_PATH)  # Lê todos os usuários e senhas
    for usuario, senha, mes, ano in dados:  # Loop através de cada usuário e senha
        driver = iniciar_driver()
        
        if tentar_login(driver, usuario, senha):
            selecionar_filtros(driver, mes, ano)  # Passa mês e ano
            pdf_url = executar_funcao_download(driver)  # Captura a URL do PDF
            # ccm = extrair_ccm(driver)  # Obtém o CCM da empresa
            # baixar_livro_fiscal(driver, usuario, ccm, mes, ano)
            if pdf_url:
                pdf_path = os.path.join(DOWNLOAD_DIR, f"Livro_Fiscal_{usuario}_{mes}_{ano}.pdf")

            else:
                print("❌ Não foi possível capturar a URL do PDF.")
        driver.quit()


if __name__ == "__main__":
    main()

