from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

# --------------------- FUNÇÃO PARA PEGAR PREÇO ---------------------
def pegar_preco(produto):
    try:
        preco = produto.find_element(By.CSS_SELECTOR, "span.a-offscreen").text.strip()
        print(preco)
    except:
        return None
    
    try: 
        inteiro = produto.find_element(By.CSS_SELECTOR, "span.a-price-whole").text.strip()
        decimal = produto.find_element(By.CSS_SELECTOR, "span.a-price-fraction").text.strip()
        return f"R$ {inteiro},{decimal}"
    except:
        return None

def pegar_nome(produto):
    try:
        nome = produto.find_element(By.CSS_SELECTOR, "h2 span").text.strip()
        return nome
    except:
        return None 

# --------------------- CONFIG DO NAVEGADOR ---------------------
options = webdriver.ChromeOptions() 
options.add_argument("--start-maximized")

navegador = webdriver.Chrome(options=options)
wait = WebDriverWait(navegador, 10)

# --------------------- LER EXCEL ---------------------
df = pd.read_excel("produtos_amazon.xlsx")

# normaliza nome das colunas
df.columns = df.columns.str.strip().str.lower()

# valida coluna
if 'produtos' not in df.columns:
    raise Exception(f"Coluna 'produtos' não encontrada. Colunas disponíveis: {df.columns}")

# --------------------- LOOP PRINCIPAL ---------------------
resultados = [] # Lista onde os resultados serão armazenados

for item in df['produtos'].astype(str).str.strip(): # Percorre cada produto do Excel Garante que é texto e sem espaços.
    if not item:
        continue

    print(f"Buscando: {item}")

    navegador.get("https://www.amazon.com.br/")

    # campo de busca
    campo_busca = wait.until(
        EC.element_to_be_clickable((By.ID, "twotabsearchtextbox")))
    campo_busca.clear()
    campo_busca.send_keys(item)
    campo_busca.send_keys(Keys.ENTER)

    # esperar carregar resultados
    wait.until(
        EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, "div.s-result-item[data-component-type='s-search-result']")))

    produtos = navegador.find_elements(
    By.CSS_SELECTOR,
    "div.s-result-item[data-component-type='s-search-result'][data-asin]:not([data-asin=''])")

    nome = None
    preco = None

    for p in produtos:
        nome_tmp = pegar_nome(p)
        preco_tmp = pegar_preco(p)

        if nome_tmp is not None and preco_tmp is not None: # Só aceita produtos que tenham nome E preço
            nome = nome_tmp
            preco = preco_tmp
            break

    if nome is None or preco is None:
        resultados.append({
            "Produto buscado": item,
            "Nome encontrado": "Não encontrado",
            "Preço": "Não encontrado"
        })
        print(f"Não encontrado: {item}")
        continue

    resultados.append({
        "Produto buscado": item,
        "Nome encontrado": nome,
        "Preço": preco
    })

# --------------------- FINALIZAÇÃO ---------------------
navegador.quit()

print(f"Total coletado: {len(resultados)}")

# salvar Excel
if resultados:
    df_resultado = pd.DataFrame(resultados)
    caminho = r"C:\Projetospy\Amazon-Automacao-de-Busca\resultado.xlsx"
    df_resultado.to_excel(caminho, index=False)
    print(f"Arquivo salvo em: {caminho}")
else:
    print("Nenhum dado foi coletado!")