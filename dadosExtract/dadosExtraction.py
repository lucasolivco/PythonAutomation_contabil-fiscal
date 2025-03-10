from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import pandas as pd
import re
import time
import os
from bs4 import BeautifulSoup

def parse_chat_data(html):
    soup = BeautifulSoup(html, 'lxml')
    
    # Encontrar todas as divs com o atributo 'title'
    divs_with_title = soup.find_all('div', title=True)
    
    for div in divs_with_title:
        # Substituir o conteúdo da div pelo valor do atributo 'title'
        div.string = div['title']
    
    # Converter o soup modificado de volta para texto
    modified_text = soup.get_text(separator='\n', strip=True)
    
    # Usar o padrão existente para extrair as informações
    pattern = r'(.*?)(\d{2}/\d{2}/\d{4} - \d{2}:\d{2})(.*?)'
    matches = re.findall(pattern, modified_text, re.DOTALL)
    
    data = []
    for match in reversed(matches):  # Invertemos a ordem para ficar cronológica
        sender, timestamp, message = match
        sender = sender.strip()
        message = message.strip()
        
        if "Atendimento" in sender:
            # Informação do sistema
            combined_message = f"Sistema {sender} {message}"
        else:
            # Mensagem normal
            combined_message = f"{sender} {message}"
        
        data.append({
            'Timestamp': timestamp,
            'Sender and message': combined_message
        })
    
    return pd.DataFrame(data)

def run_scraper():
    # Configuração do ChromeDriver
    chrome_version = "127.0.6533.120" # Substitua pela versão do seu Chrome
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("disable-infobars")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_experimental_option("useAutomationExtension", False)
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    
    # Caminho para o diretório de dados do usuário do Chrome
    user_data_dir = r"C:\Users\VMContabil\AppData\Local\Google\Chrome\User Data"
    chrome_options.add_argument(f"user-data-dir={user_data_dir}")
    
    # Nome do perfil (geralmente "Default" ou "Profile 1", "Profile 2", etc.)
    profile = "Default"
    chrome_options.add_argument(f"profile-directory={profile}")
    
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
    except Exception as e:
        print(f"Erro ao inicializar o ChromeDriver: {e}")
        print("Tentando especificar a versão manualmente...")
        try:
            service = Service(ChromeDriverManager(chrome_version=chrome_version).install())
            driver = webdriver.Chrome(service=service, options=chrome_options)
        except Exception as e:
            print(f"Erro ao inicializar o ChromeDriver com versão específica: {e}")
            return False
    
    # Garante que o navegador esteja maximizado
    driver.maximize_window()
    
    # Abre o navegador e navega para a URL especificada
    driver.get("https://app.gestta.com.br/messenger-dashboard/#/attendance-history")
    print("Navegador aberto e página carregada")
    
    # Inicializa o contador de TABs
    tab_count = 4
    max_attempts = 1000 # Limite máximo de tentativas para evitar loop infinito
    first_run = True
    # conta quantos chats foram obtidos
    Contador = 1
    
    while tab_count <= max_attempts:
        try:
            print(f"\nTentativa {tab_count}")
            
            body = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.TAG_NAME, 'body')))
            # Envia a tecla TAB o número de vezes definido pelo contador
            for _ in range(tab_count):
                body.send_keys(Keys.TAB)
                time.sleep(1)  # Pequena pausa entre cada TAB
            
            print(f"Foco realizado com {tab_count} TAB(s)")
            time.sleep(1)
            
            # Envia a tecla "Enter" diretamente para o elemento atualmente focado
            driver.switch_to.active_element.send_keys(Keys.RETURN)
            print("Tecla 'Enter' enviada para o elemento focado")
            
            try:
                # Esperar até que o elemento esteja presente
                elemento_titulo = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '#ContentBox > div.c-bTYtly.c-bTYtly-hTcQIf-theme-ONVIO.c-bTYtly-eZdROQ-visible-true.c-bTYtly-ijvnnZs-css > div.PJLV > div > div > div.c-bSkEHl > div.c-cphnPV.c-cphnPV-kGkktk-active-true > div > div > div:nth-child(2) > div:nth-child(3) > span.c-PJLV.c-PJLV-emhveM-theme-ONVIO.c-PJLV-ieFSdRm-css'))
                )
                
                # Obter o valor do atributo 'title'
                titulo = elemento_titulo.get_attribute('title')
                
                # Processar o título para usar como nome do arquivo
                nome_arquivo = titulo.replace(' - ', '_').replace(' ', '_').replace('(', '').replace(')', '').replace('/', '-')
                nome_arquivo = re.sub(r'[^\w\-_\. ]', '', nome_arquivo)  # Remove caracteres especiais
                nome_arquivo = nome_arquivo[:250] + f'_attempt_{tab_count}.xlsx'  # Adiciona o número da tentativa ao nome do arquivo
                
                print(f"Nome do arquivo: {nome_arquivo}")
                
            except Exception as e:
                print(f"Erro ao obter o título: {e}")
                if first_run:
                    print("Elemento título não encontrado na primeira tentativa. Reiniciando o processo.")
                    driver.quit()
                    return False
                nome_arquivo = f'dados_texto_div_attempt_{tab_count}.xlsx'  # Nome padrão com número da tentativa
            
            # botão histórico
            wait = WebDriverWait(driver, 12)
            elemento_hist = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ContentBox"]/div[3]/div[2]/div/div/div[1]/div/div[2]')))
            time.sleep(2)
            elemento_hist.click()
            print("Botão histórico clicado com sucesso")
            time.sleep(2)
            
            # Modifique a parte onde você extrai e processa os dados:
            elemento_div = driver.find_element(By.CSS_SELECTOR, '#ContentBox > div.c-bTYtly.c-bTYtly-hTcQIf-theme-ONVIO.c-bTYtly-eZdROQ-visible-true.c-bTYtly-ijvnnZs-css > div.PJLV > div > div > div.c-bSkEHl > div.c-cphnPV.c-cphnPV-kGkktk-active-true > div > div > div')
            
            # Extrair todo o conteúdo da `div` (HTML, texto, atributos)
            conteudo_div = elemento_div.get_attribute('outerHTML')  # Captura todo o HTML da `div`
            
            # Processar o HTML e criar o DataFrame
            df = parse_chat_data(conteudo_div)
            
            pasta_destino = r"C:\Users\VMContabil\Documents\Python automate\dadosExtract\Relatório chat onvio"
            # Certifique-se de que a pasta de destino existe
            if not os.path.exists(pasta_destino):
                os.makedirs(pasta_destino)
            
            # Função para gerar um nome de arquivo único
            def gerar_nome_arquivo_unico(pasta, nome_base):
                nome_arquivo = nome_base
                contador = 1
                while os.path.exists(os.path.join(pasta, nome_arquivo)):
                    nome_sem_extensao, extensao = os.path.splitext(nome_base)
                    nome_arquivo = f"{nome_sem_extensao}_{contador}{extensao}"
                    contador += 1
                return nome_arquivo
            
            # Use a função para gerar um nome de arquivo único
            nome_arquivo_unico = gerar_nome_arquivo_unico(pasta_destino, nome_arquivo)
            
            # Crie o caminho completo do arquivo
            caminho_completo = os.path.join(pasta_destino, nome_arquivo_unico)
            
            # Salve o arquivo
            df.to_excel(caminho_completo, index=False)
            
            print(f"Arquivo salvo como: {caminho_completo}")
            
            # Fechar aba histórico
            wait = WebDriverWait(driver, 12)
            botao_fechar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#ContentBox > div.c-bTYtly.c-bTYtly-hTcQIf-theme-ONVIO.c-bTYtly-eZdROQ-visible-true.c-bTYtly-ijvnnZs-css > div.c-dDfvGS > div > button > div.c-PJLV.c-PJLV-igMlOxk-css')))
            time.sleep(2)
            botao_fechar.click()
            print("Botão fechar clicado com sucesso")
            time.sleep(2)
            
            # implementei uma forma de que os tabs percorram de forma correta as conversas, sem pular nenhuma
            # ESSE CASO SÓ FUNCIONA CASO O ZOOM DA PAG FAÇA COM QUE APAREÇA APENAS 6 CONVERSAS POR VEZ
            # na primeira vez que rodar o tabcount permanece inalterado
            # e durante todo resto da execução sempre que ele chegar ao 9, reseta para 6
            if first_run:
                first_run = False
            else:
                if tab_count == 9:
                    tab_count = 6
                else:
                    tab_count += 1
            
            Contador += 1
            print(Contador)
            
        except (TimeoutException, NoSuchElementException) as e:
            print(f"Elemento histórico não encontrado após {tab_count} TAB(s). Finalizando o loop.")
            print(f"Erro: {e}")
            break  # Sai do loop se o elemento histórico não for encontrado
        
        except Exception as e:
            print(f"Erro inesperado: {e}")
            break  # Sai do loop em caso de erro inesperado
    
    # Fechar o navegador
    driver.quit()
    return True

# Executar o scraper até que seja bem-sucedido
while True:
    if run_scraper():
        break
    print("Reiniciando o scraper...")
    time.sleep(5)  # Pequena pausa antes de reiniciar