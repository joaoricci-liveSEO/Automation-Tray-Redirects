from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC # IMPORTAÇÃO NECESSÁRIA PARA O TIMEOUT
import openpyxl
import time

# --- Configurações de Seletores ATUALIZADAS ---
# O XPath é mais seguro aqui, buscando um botão primário com o texto exato
BTN_ADICIONAR_REDIRECIONAMENTO_XPATH = "//button[contains(@class, 'btn-primary') and contains(., 'Adicionar redirecionamento')]"

# O ID é o seletor mais rápido e confiável para os campos de input
INPUT_URL_ORIGEM_ID = "redirection-input-source_url" 
INPUT_URL_DESTINO_ID = "redirection-input-target_url" 

# O XPath é usado para encontrar o botão salvar no modal (usando o texto e a classe)
BTN_SALVAR_MODAL_XPATH = "//button[contains(@class, 'redirection-url-sidebar__button--save') and contains(text(), 'Salvar')]" 

# --- Configurações Iniciais (Mantidas) ---
NOME_ARQUIVO = 'de_para_redirecionamento.xlsx'
COLUNA_STATUS = 3 
# URL da sua loja:
URL_REDIRECIONAMENTO_TRAY = "https://colorepedrarias.commercesuite.com.br/admin/settings/redirect?sort=-id&page%5Bsize%5D=25&page%5Bnumber%5D=1" 

def iniciar_automacao():
    """Inicia o WebDriver do Chrome."""
    service = Service(ChromeDriverManager().install())
    chrome_options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.maximize_window()
    return driver

def executar_automacao():
    # Carregar a planilha
    try:
        workbook = openpyxl.load_workbook(NOME_ARQUIVO)
        sheet = workbook.active
    except FileNotFoundError:
        print(f"ERRO: Arquivo '{NOME_ARQUIVO}' não encontrado.")
        return

    driver = iniciar_automacao()
    driver.get(URL_REDIRECIONAMENTO_TRAY)
    
    # --- BLOCO DE TIMEOUT (20s) APLICADO AQUI ---
    print(f"Acessando a página de redirecionamento. Você tem 20 segundos para completar o login, se necessário.")
    
    try:
        # Espera de 20 segundos até que o botão "+ Adicionar redirecionamento" esteja visível.
        WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.XPATH, BTN_ADICIONAR_REDIRECIONAMENTO_XPATH))
        )
        print("Página de Redirecionamento carregada. Iniciando os redirecionamentos...")
    except Exception as e:
        print(f"ERRO: O botão 'Adicionar redirecionamento' não apareceu em 20 segundos.")
        print("Verifique o login ou a URL. Finalizando o script.")
        driver.quit()
        return
    # --- FIM DO BLOCO DE TIMEOUT ---

    # Itera pelas linhas da planilha (começando da linha 2, após o cabeçalho)
    for row_index, row in enumerate(sheet.iter_rows(min_row=2), start=2): 
        url_de = row[0].value
        url_para = row[1].value
        status = row[COLUNA_STATUS - 1].value 
        
        # Garante que as URLs existem e que o redirecionamento ainda não foi feito (coluna C vazia ou diferente de True)
        if status == True or not url_de or not url_para:
            print(f"Linha {row_index}: Ignorada (Status TRUE ou URLs vazias).")
            continue

        print(f"Processando Linha {row_index}: DE='{url_de}' PARA='{url_para}'")

        try:
            # 4 - Executar o clique em "+Adicionar redirecionamento"
            # O elemento já deve estar visível devido ao WebDriverWait acima
            btn_adicionar = driver.find_element(By.XPATH, BTN_ADICIONAR_REDIRECIONAMENTO_XPATH)
            btn_adicionar.click()
            time.sleep(1) # Espera o modal abrir

            # 5 - Colar a URL DE (Usando o ID)
            # É bom usar um WebDriverWait também para esperar o modal carregar
            input_origem = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, INPUT_URL_ORIGEM_ID))
            )
            input_origem.send_keys(url_de)

            # 6 - Colar a URL PARA (Usando o ID)
            input_destino = driver.find_element(By.ID, INPUT_URL_DESTINO_ID)
            input_destino.send_keys(url_para)
            
            # 7 - Clicar no botão salvar
            btn_salvar = driver.find_element(By.XPATH, BTN_SALVAR_MODAL_XPATH)
            btn_salvar.click()
            
            # Adiciona uma espera para o salvamento e recarregamento da tabela
            # (Pode ser otimizado para esperar a barra de carregamento desaparecer)
            time.sleep(4) 
            
            # 8 - Marcar na planilha
            sheet.cell(row=row_index, column=COLUNA_STATUS, value=True)
            print(f"Linha {row_index} processada e marcada como TRUE.")

        except Exception as e:
            print(f"🚨 ERRO INESPERADO na linha {row_index} ao tentar salvar: {e}")
            # Tentar fechar o modal ou dar refresh para não travar o loop
            try:
                driver.refresh()
                time.sleep(5)
            except:
                pass # Se o refresh falhar, o próximo loop pode tentar novamente
            
    # Salvar as alterações na planilha
    workbook.save(NOME_ARQUIVO)
    print("\n✅ Processo de automação finalizado. Planilha salva com status TRUE.")
    driver.quit()

if __name__ == "__main__":
    executar_automacao()