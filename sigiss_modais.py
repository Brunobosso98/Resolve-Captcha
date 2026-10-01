import time

from selenium.common.exceptions import (
    ElementClickInterceptedException,
    ElementNotInteractableException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


def _avisos_visiveis(driver: WebDriver) -> list[WebElement]:
    avisos = []
    for botao in driver.find_elements(By.CSS_SELECTOR, 'button[id="btnCiente"]'):
        try:
            if botao.is_displayed() and botao.is_enabled():
                avisos.append(botao)
        except StaleElementReferenceException:
            continue
    return avisos


def _clicar_aviso_visivel(driver: WebDriver) -> WebElement | bool:
    for botao in reversed(_avisos_visiveis(driver)):
        try:
            em_transicao = driver.execute_script(
                """
                const modal = arguments[0].closest('.modal');
                if (!modal || !window.jQuery) return false;
                const instancia = window.jQuery(modal).data('bs.modal');
                return Boolean(instancia && instancia._isTransitioning);
                """,
                botao,
            )
            if em_transicao:
                continue
            botao.click()
            return botao
        except (
            ElementClickInterceptedException,
            ElementNotInteractableException,
            StaleElementReferenceException,
        ):
            continue
    return False


def fechar_avisos_login(driver: WebDriver, timeout: float = 20) -> int:
    """Fecha avisos sucessivos ou sobrepostos, mesmo com IDs repetidos."""
    limite = time.monotonic() + timeout
    fechados = 0
    while time.monotonic() < limite:
        restante = limite - time.monotonic()
        espera = min(restante, 3) if fechados else restante
        try:
            WebDriverWait(driver, espera, poll_frequency=0.2).until(
                _avisos_visiveis
            )
        except TimeoutException:
            print(f"Avisos do login encerrados: {fechados}.")
            return fechados

        botao = WebDriverWait(
            driver, max(0, limite - time.monotonic()), poll_frequency=0.2
        ).until(
            _clicar_aviso_visivel,
            message=f"Aviso {fechados + 1} não ficou pronto para o clique.",
        )
        WebDriverWait(
            driver, max(0, limite - time.monotonic()), poll_frequency=0.2
        ).until(
            EC.invisibility_of_element(botao),
            message=f"Aviso {fechados + 1} continuou aberto após o clique.",
        )
        fechados += 1
        print(f"Aviso {fechados}: botão 'Estou Ciente' clicado com sucesso!")

    raise TimeoutException("Tempo esgotado ao fechar os avisos do login.")
