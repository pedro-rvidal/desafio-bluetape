# Bibliotecas Selenium necessárias para automatizar o navegador
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Biblioteca Pandas para manipulação e exportação de dados em tabelas
import pandas as pd

# Biblioteca BeautifulSoup para analisar e obter dados do HTML
from bs4 import BeautifulSoup

# Configuração do WebDriver para usar o Chrome, sem precisar baixar manualmente o ChromeDriver
service = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=service)

# Acessa a página inicial do Reclame Aqui e maximiza a janela do navegador
navegador.maximize_window()
navegador.get("https://www.reclameaqui.com.br/")

# Esperar até que o elemento <body> esteja presente na página, garantindo que a página foi carregada
WebDriverWait(navegador, 20).until(
    EC.presence_of_element_located((By.TAG_NAME, "body")))

def scroll_page(pixels):
    navegador.execute_script(f"window.scrollBy(0, {pixels});")

scroll_page(1000)  # Rolando a página para baixo para exibir mais conteúdo

# Localizar e clicar no botão "E-commerce - Moda" para acessar a seção de moda
botao_filtro_moda = WebDriverWait(navegador, 20).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/section[2]/div/astro-island/div/nav/div[2]/button[4]")))
botao_filtro_moda.click()


# Inicializa uma lista para armazenar os dados obtidos sobre as empresas
dados_empresa = []

# Loop para navegar pelas 3 melhores empresas e obter os dados
for i in range(1, 4):
    # Construir o XPath do link da empresa e clicar para acessar sua página de perfil
    xpath_melhor = f"/html/body/section[2]/div/astro-island/div/div[3]/div/div[1]/a[{i}]"
    empresa = WebDriverWait(navegador, 20).until(
        EC.presence_of_element_located((By.XPATH, xpath_melhor)))
    empresa.click()

    # Captura o HTML da página da empresa e o analisa usando BeautifulSoup
    pagina_html = navegador.page_source
    soup = BeautifulSoup(pagina_html, 'html.parser')

    # Obter o nome da empresa da tag <h1> na página
    nome_empresa = soup.find("h1").text.strip()

    # Obter outras informações relevantes das tags <span> com a classe específica
    informacoes = soup.find_all("span", class_="go2549335548")
    if len(informacoes) >= 6:
        reclamacoes_respondidas = informacoes[1].text.strip()  # Segundo elemento contém as reclamações respondidas
        fariam_negocio_novamente = informacoes[4].text.strip()  # Quinto elemento contém a porcentagem de clientes que fariam negócio novamente
        taxa_de_solucao = informacoes[5].text.strip()  # Sexto elemento contém a taxa de solução dos problemas

    # Obter a nota do consumidor, que está em outra tag <span> com uma classe diferente
    nota_consumidor = soup.find("span", class_="go1306724026").text.strip()

    # Determinar a posição da empresa (melhor) e formatar o texto
    posicao = f"{i}º melhor"

    # Adicionar os dados à lista
    dados_empresa.append((nome_empresa, reclamacoes_respondidas, fariam_negocio_novamente, taxa_de_solucao, nota_consumidor, posicao))

    # Realizar o scroll novamente após voltar para a lista
    scroll_page(1000)  # Rolando a página novamente

    # Voltar para a página anterior (lista de empresas) para continuar o loop
    navegador.back()

    # Re-clicar no botão Moda para garantir que a lista de empresas será exibida corretamente
    botao_filtro_moda = WebDriverWait(navegador, 20).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/section[2]/div/astro-island/div/nav/div[2]/button[4]")))
    botao_filtro_moda.click()

# Loop para navegar pelas 3 piores empresas e obter os dados
for i in range(1, 4):
    # Construir o XPath do link da empresa e clicar para acessar sua página de perfil
    xpath_pior = f"/html/body/section[2]/div/astro-island/div/div[3]/div/div[2]/a[{i}]"
    empresa = WebDriverWait(navegador, 10).until(
        EC.presence_of_element_located((By.XPATH, xpath_pior)))
    empresa.click()

    # Captura o HTML da página da empresa e o analisa usando BeautifulSoup
    pagina_html = navegador.page_source
    soup = BeautifulSoup(pagina_html, 'html.parser')

    # Obter o nome da empresa da tag <h1> na página
    nome_empresa = soup.find("h1").text.strip()
    
    # Obter outras informações relevantes das tags <span> com a classe específica
    informacoes = soup.find_all("span", class_="go2549335548")
    if len(informacoes) >= 6:
        reclamacoes_respondidas = informacoes[1].text.strip()  # Segundo elemento contém as reclamações respondidas
        fariam_negocio_novamente = informacoes[4].text.strip()  # Quinto elemento contém a porcentagem de clientes que fariam negócio novamente
        taxa_de_solucao = informacoes[5].text.strip()  # Sexto elemento contém a taxa de solução dos problemas

    # Obter a nota do consumidor, que está em outra tag <span> com uma classe diferente
    nota_consumidor = soup.find("span", class_="go1306724026").text.strip()

    # Determinar a posição da empresa (pior) e formatar o texto
    posicao = f"{i}º pior"

    # Adicionar os dados à lista
    dados_empresa.append((nome_empresa, reclamacoes_respondidas, fariam_negocio_novamente, taxa_de_solucao, nota_consumidor, posicao))
    
    # Realizar o scroll novamente após voltar para a lista
    scroll_page(1000)  # Rolando a página novamente

    # Voltar para a página anterior (lista de empresas) para continuar o loop
    navegador.back()

    # Re-clicar no botão Moda para garantir que a lista de empresas será exibida corretamente
    botao_filtro_moda = WebDriverWait(navegador, 20).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/section[2]/div/astro-island/div/nav/div[2]/button[4]")))
    botao_filtro_moda.click()


# Criar um DataFrame com os dados obtidos e salvar em um arquivo Excel
df = pd.DataFrame(dados_empresa, columns=["Empresa", "Reclamações Respondidas", "Voltariam a Fazer Negócio", "Índice de Solução", "Nota do Consumidor", "Classificação"])
df.to_excel("resultados_empresas.xlsx", index=False)

# Fechar o navegador ao final do processo
navegador.quit()
