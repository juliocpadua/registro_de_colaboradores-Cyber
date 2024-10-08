import pyautogui as posicaoMouse
import pyautogui as tempoEspera
from selenium import webdriver as opcoesSelenium
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import pandas as pd

navegador = opcoesSelenium.Chrome()
df = pd.read_excel('caminho_da_planilha.xlsx')

tempoEspera.sleep(1)

navegador.get('https://intranetcyber.azurewebsites.net')
tempoEspera.sleep(2)

posicaoMouse.click(x=979, y=1007)
tempoEspera.sleep(2)

#insere email
navegador.find_element(By.NAME, "email").send_keys("rodrigo.mucedola@cybersolutions.com.br")
tempoEspera.sleep(2)

#insere senha
navegador.find_element(By.NAME, "senha").send_keys("rodrigo10*")
tempoEspera.sleep(1)

#dá o enter para fazer o login
posicaoMouse.hotkey('enter')
tempoEspera.sleep(5)

posicaoMouse.click(x=40, y=229)
tempoEspera.sleep(2)

posicaoMouse.click(x=258, y=819)
tempoEspera.sleep(3)

navegador.find_element(By.NAME, "Nome").send_keys("Júlio César Oliveira")
tempoEspera.sleep(1)

navegador.find_element(By.NAME, "CPF").send_keys("1290942619")
tempoEspera.sleep(1)

navegador.find_element(By.NAME, "DataNascimento").send_keys("24121999")
tempoEspera.sleep(2)

navegador.find_element(By.NAME, "Email").send_keys("julio.cesar@cybersolutions.com.br")
tempoEspera.sleep(1)

navegador.find_element(By.NAME, "SenhaTemporaria").send_keys("Senha12345")
tempoEspera.sleep(2)

pegarDropDown1 = navegador.find_element(By.NAME, "CargoCodigo")
itens = Select(pegarDropDown1)
itens.select_by_value("2")
tempoEspera.sleep(2)

pegarDropDown2 = navegador.find_element(By.NAME, "GerenteCPF")
itens1 = Select(pegarDropDown2)
itens1.select_by_value("12345678910")
tempoEspera.sleep(2)


pegarDropDown4 = navegador.find_element(By.NAME, "CodigoCliente")
itens3 = Select(pegarDropDown4)
itens3.select_by_value("1")
tempoEspera.sleep(2)

navegador.find_element(By.NAME, "DataAdmissao").send_keys("13092023")
tempoEspera.sleep(2)

radio_button = navegador.find_element(By.XPATH, f"//input[@name='Isadm' and @value='false']")
radio_button.click()

tempoEspera.sleep(10)

navegador.quit()


