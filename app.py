from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl

# Acesse o site
url = 'https://www.kabum.com.br/computadores/notebooks'
driver = webdriver.Chrome()
driver.get(url)

# Aguarde até que os elementos estejam presentes na página
wait = WebDriverWait(driver, 10)
titulos = wait.until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, 'span.sc-d79c9c3f-0.nlmfp.sc-cdc9b13f-16.eHyEuD.nameCard')))

precos = wait.until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, 'span.priceCard')))

# Criando a planilha
workbook = openpyxl.Workbook()
# Criando a página 'produtos'
sheet_produtos = workbook.create_sheet('produtos')
# Selecionar a página produtos
sheet_produtos['A1'].value = 'Produto'
sheet_produtos['B1'].value = 'Preço'

# Inserir os títulos e preços na planilha
for titulo, preco in zip(titulos, precos):
    sheet_produtos.append([titulo.text, preco.text])

# Salvar a planilha
workbook.save('produtos.xlsx')

# Fechar o navegador
driver.quit()
