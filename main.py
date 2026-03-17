import os
import time
import base64
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

df = pd.read_excel("Certificados.xlsx")

if "Status" not in df.columns:
    df["Status"] = ""

os.makedirs("certificados", exist_ok=True)

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless=new")  
chrome_options.add_argument("--disable-gpu")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

for i, url in enumerate(df["Link do certificado"], start=0):
    try:
        driver.get(url)
        time.sleep(2)  

        # Captura o título da aba
        titulo = driver.title.strip()

        # Remove caracteres inválidos para nome de arquivo no Windows
        for ch in r'\/:*?"<>|':
            titulo = titulo.replace(ch, "")

        # Gera o PDF via DevTools
        pdf = driver.execute_cdp_cmd("Page.printToPDF", {"format": "A4"})
        pdf_bytes = base64.b64decode(pdf['data'])

        # Salva com o nome do título
        output = f"certificados/{titulo}.pdf"
        with open(output, "wb") as f:
            f.write(pdf_bytes)

        # Atualiza status como "Baixado"
        df.at[i, "Status"] = "Baixado"
        print(f"✅ Gerado: {output}")

    except Exception as e:
        # Atualiza status como "Erro"
        df.at[i, "Status"] = "Erro"
        print(f"❌ Erro no link {url}: {e}")

    # Salva progresso a cada 100 registros para evitar perda de dados caso o script seja interrompido
    if (i + 1) % 100 == 0:
        df.to_excel("Certificados_Status.xlsx", index=False)
        print(f"💾 Progresso salvo até a linha {i+1}")

# Salva a versão final ao terminar tudo
df.to_excel("Certificados_Status.xlsx", index=False)
print("✅ Processamento concluído e planilha salva!")
driver.quit()
