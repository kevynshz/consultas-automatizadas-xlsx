import pandas as pd
import pyperclip
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

excel_file = "REFAZER 23 12--10 01--25 01automacao.xlsx"
df = pd.read_excel(excel_file, header=None, dtype=str)

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)
driver.get("https://protestosp.com.br/consulta-de-protesto")

def pausar(mensagem):
    print(f"\n[PAUSA] {mensagem}")
    input("Pressione Enter para continuar...")

def formatar_documento(documento):
    if pd.isna(documento):
        return ""
    documento = "".join(filter(str.isdigit, str(documento)))
    if len(documento) == 11:
        return f"{documento[:3]}.{documento[3:6]}.{documento[6:9]}-{documento[9:]}"
    elif len(documento) == 14:
        return f"{documento[:2]}.{documento[2:5]}.{documento[5:8]}/{documento[8:12]}-{documento[12:]}"
    else:
        return ""

for index, row in df.iterrows():
    cpf_cnpj = formatar_documento(row[1])

    if not cpf_cnpj:
        print(f"Linha {index + 1}: Documento vazio ou inválido. Ignorando...")
        continue

    coluna_c = row[2] if not pd.isna(row[2]) else ""

    tipo_documento = "CPF" if len(cpf_cnpj.replace(".", "").replace("-", "").replace("/", "")) == 11 else "CNPJ"
    print(f"Documento: {cpf_cnpj}, Tipo: {tipo_documento}")

    try:
        input_element = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "TipoDocumento"))
        )
        input_element.click()

        opcao = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f'//select[@id="TipoDocumento"]/option[text()="{tipo_documento}"]'))
        )
        opcao.click()

    except Exception as e:
        print(f"Erro ao selecionar {tipo_documento}: {e}")
        continue

    search_box = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "Documento"))
    )
    search_box.clear()

    pyperclip.copy(cpf_cnpj)

    search_box.send_keys(pyperclip.paste())

    consultar_button = driver.find_element(By.XPATH, '//*[@id="frmConsulta"]/input[4]')
    driver.execute_script("arguments[0].click();", consultar_button)

    pausar("Verifique e resolva o CAPTCHA no navegador, caso solicitado.")

    try:
        aviso_protesto = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="resumoConsulta"]'))
        )
        html_aviso = aviso_protesto.get_attribute("innerHTML")
        
        if "Constam protestos" in html_aviso:
            print(f"{tipo_documento} {cpf_cnpj}: Com Protesto (Coluna C: {coluna_c})")
        elif "Não constam protestos" in html_aviso:
            print(f"{tipo_documento} {cpf_cnpj}: Sem Protesto (Coluna C: {coluna_c})")
        else:
            print(f"{tipo_documento} {cpf_cnpj}: Sem Protesto (Coluna C: {coluna_c})")

    except Exception as e:
        print(f"Erro ao verificar protesto para {tipo_documento} {cpf_cnpj} (Coluna C: {coluna_c})")

    time.sleep(7)

    driver.get("https://protestosp.com.br/consulta-de-protesto")

driver.quit()
