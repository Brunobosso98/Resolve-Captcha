from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import pytesseract
from PIL import Image
import time
import pandas as pd

# Configuração do caminho do Tesseract
TESSERACT_PATH = r'C:\Users\bruno.martins\Desktop\robo Busca NFSe\dependencias sistema\Tesseract-OCR\tesseract.exe'
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

URL_LOGIN = 'https://itapira.sigiss.com.br/itapira/contribuinte/login.php'
EXCEL_PATH = r'C:\Users\bruno.martins\Desktop\ResolveCaptcha\Senha Municipio Itapira.xlsx'


def iniciar_driver():
    """Inicializa o WebDriver."""
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get(URL_LOGIN)
    time.sleep(5)
    return driver


def ler_dados_excel(caminho):
    """Lê o arquivo Excel e retorna o usuário e senha."""
    dados = pd.read_excel(caminho)
    return str(dados['Usuário'][0]), str(dados['Senha'][0])


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


def selecionar_filtros(driver):
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
        select.select_by_visible_text("Dezembro")  # Seleciona o mês "Dezembro"

        # Localizar o elemento do select para o ano
        select_ano = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="panelFiltro"]/table/tbody/tr/td[7]/input'))
        )
        select_ano.clear()
        select_ano.send_keys("2024")

        botao_ok = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'btnOk')))
        botao_ok.click()
    except Exception as e:
        print(f"Erro ao selecionar filtros: {e}")


def main():
    driver = iniciar_driver()
    usuario, senha = ler_dados_excel(EXCEL_PATH)
    
    if tentar_login(driver, usuario, senha):
        selecionar_filtros(driver)
    
    driver.quit()


if __name__ == "__main__":
    main()
