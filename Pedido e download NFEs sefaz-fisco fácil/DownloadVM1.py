import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from descompactar_e_mover import executar_descompactacao_e_movimentacao
import pyautogui
import time
import pyperclip
import os
import zipfile
import shutil
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import openpyxl
from datetime import datetime

def capture_screenshot(driver, base_path, file_name):
    """
    Captura uma screenshot da página atual e a salva com o nome e caminho especificados.
    
    :param driver: Instância do webdriver
    :param base_path: Caminho base para salvar a screenshot
    :param file_name: Nome do arquivo para a screenshot
    :return: Caminho do arquivo salvo
    """
    # Garante que o diretório base existe
    os.makedirs(base_path, exist_ok=True)

    # Cria o nome do arquivo com timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    full_file_name = f"{file_name}_{timestamp}.png"
    
    # Cria o caminho completo do arquivo
    file_path = os.path.join(base_path, full_file_name)
    
    # Captura a screenshot
    driver.save_screenshot(file_path)
    print(f"Screenshot salva em: {file_path}")
    return file_path

def mover_arquivos_xml(origem, destino):
    """
    Move arquivos .xml da pasta de origem para o destino especificado.
    
    :param origem: Caminho da pasta de origem dos arquivos .xml
    :param destino: Caminho da pasta de destino para os arquivos .xml
    :return: Uma string com o log da operação
    """
    log = []
    arquivos_xml = [f for f in os.listdir(origem) if f.endswith('.xml')]
    
    # Verificar se o diretório de destino existe, se não, criar
    if not os.path.exists(destino):
        os.makedirs(destino)
        log.append(f"Diretório criado: {destino}")
    
    # Mover os arquivos .xml para o destino
    for arquivo in arquivos_xml:
        origem_completa = os.path.join(origem, arquivo)
        destino_completo = os.path.join(destino, arquivo)
        
        try:
            shutil.move(origem_completa, destino_completo)
            log.append(f"Arquivo movido: {arquivo} para {destino}")
        except Exception as e:
            log.append(f"Erro ao mover {arquivo}: {str(e)}")
    
    return "\n".join(log)

def run_script():
    # Configuração do ChromeDriver
    chrome_version = "127.0.6533.120"  # Substitua pela versão do seu Chrome
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("disable-infobars")
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
            return
    

    try:
        
        # Garante que o navegador esteja maximizado
        driver.maximize_window()
        
        # Abre o navegador e navega para a URL especificada
        driver.get("https://ssa.fazenda.rj.gov.br/ssa/")
        print("Navegador aberto e página carregada")

        time.sleep(5)  

        pyautogui.click(896, 245)
        print("Clique certificado")
        time.sleep(2)

        pyautogui.click(535, 246)
        print("clique levi")
        time.sleep(2)

        pyautogui.click(794, 354)
        print("clique ok")
        time.sleep(2)

        # Aguarda até que o elemento esteja presente e clicável
        wait = WebDriverWait(driver, 10)

        button = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@onclick, \"carregarAutorizacoes('AUTO Fisco Fácil', 53917)\")]")))
        print("Botão encontrado, tentando clicar...")
        # Clica no botão diretamente
        button.click()
        print("Botão clicado")
        
        time.sleep(10)
        try:              
            
            wait = WebDriverWait(driver, 8)
            
            spanAutorizar = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'autorizado')]"))
            )
            print("Elemento 'autorizado' encontrado! Prosseguindo...")
            
            # Clica no span
            spanAutorizar.click()
            print("Elemento clicado")
        except (TimeoutException, NoSuchElementException) as e:
            print(f"Erro ao tentar encontrar ou clicar no elemento 'autorizado': {e}") 
        
        time.sleep(5)

        # Verifica se a página correta foi carregada verificando a presença do input com o ID 'FrmFisco:valorDaPesquisa'
        if driver.find_elements(By.ID, 'FrmFisco:valorDaPesquisa'):
            print("Página correta carregada. Continuando a execução...")
        else:
            print("Página correta não carregada. Reiniciando o navegador...")
            driver.quit()
            run_script()  # Executa o script novamente
        
        
        # Ler dados do Excel: colunas A, B, C, D etc. (0, 1, 2, 3 etc.) e linhas 4-95
        df = pd.read_excel(r'C:\Users\VMContabil\Documents\baseFiscalAuto - EndereçoPastasVM1.xlsx', usecols="A:M", skiprows=0, nrows=92)  # Ajuste o intervalo de colunas conforme necessário; nrows: quantidades de linhas que vai percorrer a partir das que pularam no skiprows
        # Criar um arquivo de log geral
        log_dir = r"C:\Users\VMContabil\Documents\log_movimentacao"
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, f"log_movimentacao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        # Iterar sobre as linhas do DataFrame
        #index = quantidade de vezes que vai percorrer de acordo com as linhas; row = numero da coluna
        with open(log_file, 'w') as log:
            for index, row in df.iterrows():
                print(f"Preenchendo dados da linha {index + 2}")

                # Exemplo: preencher múltiplos campos com dados de diferentes colunas
                valor_pesquisa = row.iloc[0]  # Substitua pelo índice da coluna no Excel

                # Especifique os caminhos das pastas que você quer limpar
                pasta_para_apagar1 = r"C:\Users\VMContabil\Documents\07. xmlDownload"  # Substitua pelo caminho real da primeira pasta
                pasta_para_apagar2 = r"C:\Users\VMContabil\Documents\destinoXml"   # Substitua pelo caminho real da segunda pasta

                def apagar_pasta(pasta):
                    if not os.path.exists(pasta):
                        print(f"O caminho '{pasta}' não existe.")
                        return False
                    if not os.path.isdir(pasta):
                        print(f"'{pasta}' não é uma pasta.")
                        return False
                    
                    try:
                        for item in os.listdir(pasta):
                            item_path = os.path.join(pasta, item)
                            if os.path.isfile(item_path):
                                os.remove(item_path)
                                print(f"Arquivo apagado: {item}")
                            elif os.path.isdir(item_path):
                                shutil.rmtree(item_path)
                                print(f"Pasta apagada: {item}")
                        print(f"Todos os itens na pasta '{pasta}' foram apagados.")
                        return True
                    except Exception as e:
                        print(f"Ocorreu um erro ao apagar os arquivos em '{pasta}': {e}")
                        return False

                     
                resultado1 = apagar_pasta(pasta_para_apagar1)
                resultado2 = apagar_pasta(pasta_para_apagar2)
                
                if resultado1 and resultado2:
                    print("Operação concluída com sucesso para ambas as pastas.")
                elif resultado1:
                    print("Operação concluída com sucesso apenas para a primeira pasta.")
                elif resultado2:
                    print("Operação concluída com sucesso apenas para a segunda pasta.")
                else:
                    print("A operação falhou para ambas as pastas.")
                

                pyperclip.copy(str(valor_pesquisa))

                # Preencher o campo de pesquisa
                wait = WebDriverWait(driver,12)
                input_field = wait.until(EC.element_to_be_clickable((By.ID, 'FrmFisco:valorDaPesquisa')))
                time.sleep(2)
                input_field.send_keys(Keys.CONTROL, 'v')
                time.sleep(1)
                print(f"Campo de pesquisa preenchido com: {valor_pesquisa}")

                time.sleep(2)

                # Adicione aqui qualquer outra ação que deseja realizar após preencher os campos
                wait = WebDriverWait(driver,12)
                botao_filtrar = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Filtrar')]")))
                botao_filtrar.click()
                print("Botão 'Filtrar' clicado")

                time.sleep(3)
                #click empresa
                pyautogui.click(831,493)
                time.sleep(2)
                
                wait = WebDriverWait(driver, 12)
                botao_solicitacoes = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@href='#frmHistInteracoes:tabsHist:tabSolictacao']")))
                time.sleep(2)
                botao_solicitacoes.click()
                print("Botão 'Solicitações' clicado")
                time.sleep(1)

                try:
                    # Extrai o caminho base e o nome do arquivo das colunas do Excel
                    base_path = row.iloc[5]  # Substitua pelo nome real da coluna
                    file_name = row.iloc[12]  # Substitua pelo nome real da coluna

                    capture_screenshot(driver, base_path, file_name)
                except Exception as e:
                    print(f"Ocorreu um erro: {e}")
                    # Captura screenshot em caso de erro (usando um diretório padrão)
                    capture_screenshot(driver, "screenshots", "erro_captura")


                #click 1º processada
                wait = WebDriverWait(driver, 12)
                processada1 = wait.until(EC.element_to_be_clickable((By.ID, "frmHistInteracoes:tabsHist:solicitacao:0:linkSolicitacao")))
                time.sleep(2)
                processada1.click()
                print("Link com id 'frmHistInteracoes:tabsHist:solicitacao:0:linkSolicitacao' clicado")
                time.sleep(1)

                def main(row):
                    print("Iniciando o processo principal...")

                    # Executar o subfluxo de descompactação e movimentação
                    executar_descompactacao_e_movimentacao()

                    # Após a descompactação, mover os arquivos XML
                    pasta_origem = r"C:\Users\VMContabil\Documents\destinoXml"
                    pasta_destino = row.iloc[4]  # Substitua 'CaminhoPasta' pelo nome real da coluna
                    
                    resultado_movimentacao = mover_arquivos_xml(pasta_origem, pasta_destino)
                    log.write(f"Movimentação para empresa da linha {index + 2}:\n")
                    log.write(resultado_movimentacao + "\n\n")
                    print(f"Arquivos XML movidos para {pasta_destino}")

                    print("Processo principal concluído.")

                def tentar_clicar_botao(driver):
                    try:
                        wait = WebDriverWait(driver, 12)
                        DownloadButton = wait.until(EC.element_to_be_clickable((By.ID, "FrmSolicitarExtracaoDfe:resultado:0:j_idt110")))
                        time.sleep(2)
                        DownloadButton.click()
                        print("Botão de download clicado")
                        time.sleep(1)
                        return True
                    except TimeoutException:
                        print("Botão de download não encontrado. Pulando o processo principal.")
                        return False

                if __name__ == "__main__":
                    # Assumindo que 'driver' já está definido e inicializado
                    botao_encontrado = tentar_clicar_botao(driver)
                    
                    if botao_encontrado:
                        main(row)
                        time.sleep(10)
                    else:
                        print("O processo principal foi pulado devido à falta do botão de download.")

                        

                print(f"Processo completo. Log de movimentação salvo em: {log_file}")

                #Click Página Principal
                wait = WebDriverWait(driver, 10)
                link = wait.until(EC.element_to_be_clickable((By.ID, "j_idt23")))
                time.sleep(2)
                link.click()
                print("Pagina principal clicado")
                
                #Click Solicitações
                wait = WebDriverWait(driver, 12)
                botao_solicitacoes = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@href='#frmHistInteracoes:tabsHist:tabSolictacao']")))
                time.sleep(2)
                botao_solicitacoes.click()
                print("Botão 'Solicitações' clicado")
                time.sleep(1)

                #Click processada2
                wait = WebDriverWait(driver, 12)
                processada2 = wait.until(EC.element_to_be_clickable((By.ID, "frmHistInteracoes:tabsHist:solicitacao:1:linkSolicitacao")))
                time.sleep(2)
                processada2.click()
                print("Link com id 'frmHistInteracoes:tabsHist:solicitacao:0:linkSolicitacao' clicado")
                time.sleep(1)
                    
                if __name__ == "__main__":
                    # Assumindo que 'driver' já está definido e inicializado
                    botao_encontrado = tentar_clicar_botao(driver)
                    
                    if botao_encontrado:
                        main(row)
                        time.sleep(10)
                    else:
                        print("O processo principal foi pulado devido à falta do botão de download.")

                        

                print(f"Processo completo. Log de movimentação salvo em: {log_file}")

                #Click Página Principal
                wait = WebDriverWait(driver, 10)
                link = wait.until(EC.element_to_be_clickable((By.ID, "j_idt23")))
                time.sleep(2)
                link.click()
                print("Pagina principal clicado")
                
                #Click Solicitações
                wait = WebDriverWait(driver, 12)
                botao_solicitacoes = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@href='#frmHistInteracoes:tabsHist:tabSolictacao']")))
                time.sleep(2)
                botao_solicitacoes.click()
                print("Botão 'Solicitações' clicado")
                time.sleep(1)

                #Click processada3
                wait = WebDriverWait(driver, 12)
                processada3 = wait.until(EC.element_to_be_clickable((By.ID, "frmHistInteracoes:tabsHist:solicitacao:2:linkSolicitacao")))
                time.sleep(2)
                processada3.click()
                print("Link com id 'frmHistInteracoes:tabsHist:solicitacao:0:linkSolicitacao' clicado")
                time.sleep(1)

                if __name__ == "__main__":
                    # Assumindo que 'driver' já está definido e inicializado
                    botao_encontrado = tentar_clicar_botao(driver)
                    
                    if botao_encontrado:
                        main(row)
                        time.sleep(10)
                    else:
                        print("O processo principal foi pulado devido à falta do botão de download.")

                        

                print(f"Processo completo. Log de movimentação salvo em: {log_file}")
                
                #Selecionar estabelecimento
                wait = WebDriverWait(driver, 12)
                selecionar_estabelecimento = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@style, 'margin-left:0.5em') and text()='Selecionar estabelecimento']")))
                time.sleep(2)
                selecionar_estabelecimento.click()
                print(f"selecionar estabelecimento")
       

    finally:
        # Fecha o navegador
        driver.quit()
        print("Navegador fechado")

run_script()