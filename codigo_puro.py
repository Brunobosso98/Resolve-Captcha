from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import pytesseract
from PIL import Image
import time
import pandas as pd  # Importar a biblioteca pandas

# Configurar o caminho do Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\bruno.martins\Desktop\robo Busca NFSe\dependencias sistema\Tesseract-OCR\tesseract.exe'

def extrair_numeros_imagem(driver):
    numeros = None
    try:
        # Localizar o elemento da imagem
        elemento_imagem = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/div[1]/div[2]/div/div[2]/div[4]/div/div/div/span/img'))
        )

        # Tirar screenshot e processar
        driver.save_screenshot('screenshot.png')
        screenshot = Image.open('screenshot.png')
        
        # Calcular coordenadas
        location = elemento_imagem.location
        size = elemento_imagem.size
        left = location['x']
        top = location['y']
        right = location['x'] + size['width']
        bottom = location['y'] + size['height']

        # Recortar e processar imagem
        imagem = screenshot.crop((left, top, right, bottom))
        imagem.save('numero.png')
        
        # OCR com pré-processamento
        imagem = imagem.convert('L')  # Converter para escala de cinza
        numeros = pytesseract.image_to_string(imagem, config='--psm 6 -c tessedit_char_whitelist=0123456789')
        numeros = ''.join(filter(str.isdigit, numeros))

    except Exception as e:
        print(f"Erro na extração: {e}")
    
    return numeros

def digitar_captcha(driver, numeros):
    try:
        # Localizar e interagir com o campo
        campo = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "confirma"))
        )
        campo.clear()
        campo.send_keys(numeros)
        print("Captcha digitado com sucesso!")
        
        # Adicionar qualquer lógica adicional pós-digitação aqui
        # Exemplo: clicar em enviar
        # driver.find_element(By.XPATH, "xpath_do_botao").click()
        
    except Exception as e:
        print(f"Erro ao digitar captcha: {e}")

def preencher_campos(driver, usuario, senha):
    try:
        # Preencher o campo de CNPJ
        campo_cnpj = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "cnpj"))
        )
        campo_cnpj.clear()
        campo_cnpj.send_keys(usuario)
        
        # Preencher o campo de senha
        campo_senha = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "senha"))
        )
        campo_senha.clear()
        campo_senha.send_keys(senha)
        
        print("Campos preenchidos com sucesso!")
        
    except Exception as e:
        print(f"Erro ao preencher campos: {e}")

def main():
    driver = None
    try:
        driver = webdriver.Chrome()
        driver.get('https://itapira.sigiss.com.br/itapira/contribuinte/login.php')
        
        # Esperar página carregar
        time.sleep(5)
        
        # Ler dados do Excel
        dados = pd.read_excel(r'C:\Users\bruno.martins\Desktop\ResolveCaptcha\Senha Municipio Itapira.xlsx')  # Substitua pelo caminho do seu arquivo
        usuario = str(dados['Usuário'][0])  # Converte para string
        senha = str(dados['Senha'][0])      # Converte para string
        
        while True:  # Loop para tentar novamente se o captcha não for reconhecido
            # Extrair números
            numeros = extrair_numeros_imagem(driver)
            
            if numeros:
                print(f"Números extraídos: {numeros}")
                # Preencher campos de CNPJ e senha
                preencher_campos(driver, usuario, senha)
                # Digitar no campo captcha
                digitar_captcha(driver, numeros)

                botao_acessar = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "btnOk")))
                botao_acessar.click()

                botao_modificar = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "btnAlterar")))
                botao_modificar.click()
                time.sleep(2)

                # Localizar o elemento do select para o mês
                select_mes = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@name="mes"]'))
                )
                # Criar um objeto Select e escolher o mês
                select = Select(select_mes)
                select.select_by_visible_text("Dezembro")  # Seleciona o mês "Dezembro"

                # Localizar o elemento do select para o mês
                select_ano = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="panelFiltro"]/table/tbody/tr/td[7]/input'))
                )
                select_ano.clear()
                select_ano.send_keys("2024")

                botao_ok = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, 'btnOk')))
                botao_ok.click()

                time.sleep(5)
                break  # Sai do loop se o captcha foi reconhecido
            else:
                print("Nenhum número foi detectado! Tentando novamente...")
                driver.quit()  # Fecha o navegador
                driver = webdriver.Chrome()  # Reinicia o navegador
                driver.get('https://itapira.sigiss.com.br/itapira/contribuinte/login.php')
                time.sleep(5)  # Espera a página carregar novamente
                    
    except Exception as e:
        print(f"Erro geral: {e}")
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    main()
