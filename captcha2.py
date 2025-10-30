from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pytesseract
from PIL import Image
import pandas as pd
import time

# Configurações
CAMINHO_TESSERACT = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
CAMINHO_EXCEL = r'C:\Users\bruno\Desktop\Automação\Resolve-Captcha\Senha Municipio Itapira.xlsx'
URL_LOGIN = 'https://itapira.sigiss.com.br/itapira/contribuinte/login.php'

pytesseract.pytesseract.tesseract_cmd = CAMINHO_TESSERACT

def extrair_numeros_imagem(driver, wait):
    numeros = None
    try:
        # Localizar o elemento da imagem
        elemento_imagem = wait.until(
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
        print(f"Erro na extração: {type(e).__name__}: {e}")
    
    return numeros

def preencher_campo(driver, element_id, valor, wait):
    try:
        campo = wait.until(
            EC.element_to_be_clickable((By.ID, element_id))
        )
        campo.clear()
        campo.send_keys(valor)
        print(f"Campo {element_id} preenchido")
    except Exception as e:
        print(f"Erro ao preencher {element_id}: {e}")

def processar_login(driver, wait):
    try:        
        # Verificar se o login foi bem-sucedido pelo URL
        if driver.current_url == 'https://itapira.sigiss.com.br/itapira/contribuinte/main.php':
            return True  # Login bem-sucedido

        # Verificar se há mensagem de erro de login específico (Contribuinte Inexistente ou Senha Inválida)
        current_url = driver.current_url
        if "msg=Contribuinte+Inexistente+ou+Senha+Inv%E1lida" in current_url:
            print("Erro de login: Contribuinte Inexistente ou Senha Inválida")
            return False  # Indica que o login falhou devido a credenciais incorretas

        # Verificar se há mensagem de erro genérica
        try:
            erro_elemento = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/div[2]/div/font/b/center')))
            if erro_elemento.is_displayed():
                print("Erro no login: ", erro_elemento.text)
                return False  # Indica que o login falhou
        except Exception:
            pass  # Se não encontrar o elemento de erro, assume que não houve erro

    except Exception as e:
        print(f"Erro no processo de login: {e}")
        return False  # Indica que o login falhou


def digitar_captcha(driver, numeros, wait):
    try:
        # Localizar e interagir com o campo
        campo = wait.until(
            EC.element_to_be_clickable((By.ID, "confirma"))
        )
        campo.clear()
        campo.send_keys(numeros)
        print("Captcha digitado com sucesso!")
        
        botao_logar = wait.until(EC.element_to_be_clickable((By.ID, "btnOk")))
        botao_logar.click()
        
    except Exception as e:
        print(f"Erro ao digitar captcha: {e}")

def preencher_data(driver, wait, mes, ano):
    try:
        # Localizar e interagir com o campo
        campo_modificar = wait.until(
            EC.element_to_be_clickable((By.ID, "btnAlterar"))
        )
        campo_modificar.click()

        campo_mes = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="panelFiltro"]/table/tbody/tr/td[3]/select')))
        campo_mes.send_keys(mes)
        print(f"Mês '{mes}' digitado com sucesso!")
        
        campo_ano = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="panelFiltro"]/table/tbody/tr/td[7]/input')))
        campo_ano.clear()
        campo_ano.send_keys(ano)
        print(f"Ano '{ano}' digitado com sucesso!")
        
        # Clicar no botão OK após preencher mês e ano
        botao_ok = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "btn-success")))
        botao_ok.click()
        print("Botão OK clicado com sucesso!")
        
        # Após preencher data, verificar e clicar em Encerramento Fiscal se disponível
        clicar_encerramento_fiscal(driver, wait)
                
    except Exception as e:
        print(f"Erro ao preencher data: {e}")

def clicar_encerramento_fiscal(driver, wait):
    try:
        # Verificar se o botão "Serviços Prestados" existe
        servicos_prestados_btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Serviços Prestados')]"))
        )
        
        # Clicar no botão "Serviços Prestados" para abrir o dropdown
        servicos_prestados_btn.click()
        print("Botão 'Serviços Prestados' clicado com sucesso!")
        
        # Aguardar o dropdown abrir e clicar em "Encerramento Fiscal"
        encerramento_fiscal_link = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@onclick, 'fechamento/prestado.php')]"))
        )
        encerramento_fiscal_link.click()
        print("Link 'Encerramento Fiscal' clicado com sucesso!")
        
        # Aguardar e mudar para a nova aba/janela
        time.sleep(3)  # Aguardar a nova aba carregar
        handles = driver.window_handles
        if len(handles) > 1:
            driver.switch_to.window(handles[-1])  # Mudar para a última aba aberta
            print("Mudança para nova aba realizada com sucesso!")
        
        # Procurar e clicar no botão "Encerrar Mes"
        botao_encerrar = wait.until(
            EC.element_to_be_clickable((By.ID, "btnSalvar"))
        )
        botao_encerrar.click()
        print("Botão 'Encerrar Mes' clicado com sucesso!")
        
        # Lidar com o alerta do navegador com fallback de 2 minutos
        try:
            # Aguardar até 2 minutos pelo alerta
            from selenium.webdriver.support.ui import WebDriverWait
            fallback_wait = WebDriverWait(driver, 120)  # 2 minutos de espera
            alert = fallback_wait.until(EC.alert_is_present())
            alert_text = alert.text
            print(f"Alerta encontrado: {alert_text}")
            alert.accept()  # Clica em OK no alerta
            print("Alerta aceito com sucesso!")
        except:
            print("Nenhum alerta encontrado ou não foi possível interagir com ele após 2 minutos. Continuando para a próxima empresa.")
        
    except Exception as e:
        print(f"Botão 'Serviços Prestados' não encontrado ou erro ao clicar em 'Encerramento Fiscal': {e}")

def main():
    try:
        # Ler dados do Excel
        df = pd.read_excel(CAMINHO_EXCEL, engine='openpyxl')

        # Para cada linha no Excel
        for index, row in df.iterrows():
            driver = webdriver.Chrome()
            driver.maximize_window()
            wait = WebDriverWait(driver, 20)
            driver.get(URL_LOGIN)
            
            time.sleep(2)  # Adiciona uma pausa de 5 segundos para garantir que a página carregue completamente
            
            try:
                print(f"Processando linha {index + 1}: {row['Empresa']}")
                
                # Preencher credenciais antes de extrair o captcha
                preencher_campo(driver, "cnpj", row['Usuário'], wait)
                preencher_campo(driver, "senha", row['Senha'], wait)
                
                # Extrair números
                numeros = extrair_numeros_imagem(driver, wait)
                
                if numeros:
                    print(f"Números extraídos: {numeros}")
                    # Digitar no campo
                    digitar_captcha(driver, numeros, wait)
                    time.sleep(10)
                else:
                    print("Nenhum número foi detectado!")
                
                if processar_login(driver, wait):
                    preencher_data(driver, wait, row['Mês'], row['Ano'])  # Chama preencher_data apenas se o login for bem-sucedido
                else:
                    print("Login falhou, não preenchendo a data.")
                    continue  # Pula para a próxima empresa se o login falhar
            finally:
                driver.quit()
                
    except Exception as e:
        print(f"Erro geral: {e}")

if __name__ == "__main__":
    main()

    