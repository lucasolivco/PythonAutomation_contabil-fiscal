import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import pyautogui
import time
import pyperclip
from datetime import datetime

def run_script():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("disable-infobars")

    # Configura o ChromeDriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    # Limpa o cache usando comandos do DevTools
    driver.execute_cdp_cmd('Network.clearBrowserCache', {})

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
            def verificar_mensagem_solicitacao(driver, timeout=10):
                try:
                    # Acessar o iframe pelo título usando XPath
                    iframe = WebDriverWait(driver, timeout).until(
                        EC.presence_of_element_located((By.XPATH, "//iframe[@title='Solicitação de Extração de DFe']"))
                    )
                    driver.switch_to.frame(iframe)
                    print("Mudou para o iframe 'Solicitação de Extração de DFe'.")

                    # Procurar pelo elemento específico dentro do iframe
                    elemento = WebDriverWait(driver, timeout).until(
                        EC.presence_of_element_located((By.XPATH, 
                            "//div[contains(@class, 'ui-panelgrid-cell') and contains(text(), 'O resultado será apresentado')]"
                        ))
                    )
                    
                    print(f"Mensagem encontrada: {elemento.text}")
                    return True

                except TimeoutException as e:
                    print(f"Timeout: Elemento não encontrado. Erro: {e}")
                    return False
                except Exception as e:
                    print(f"Erro ao verificar o pedido: {e}")
                    return False
                finally:
                    # Sempre volta ao contexto padrão, mesmo se ocorrer uma exceção
                    driver.switch_to.default_content()
                    print("Voltou ao contexto padrão da página.")
                        
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
        df = pd.read_excel(r'C:\Users\VMContabil\Documents\baseFiscalAuto - EndereçoPastasVM1.xlsx', usecols="A:D", skiprows=5, nrows=89)  # Ajuste o intervalo de colunas conforme necessário; nrows: quantidades de linhas que vai percorrer a partir das que pularam no skiprows

        # Iterar sobre as linhas do DataFrame
        #index = quantidade de vezes que vai percorrer de acordo com as linhas; row = numero da coluna
        for index, row in df.iterrows():
            print(f"Preenchendo dados da linha {index + 2}")

            # Exemplo: preencher múltiplos campos com dados de diferentes colunas
            valor_pesquisa = row.iloc[0]  # Substitua pelo índice da coluna no Excel
            valor_dataInicio = row.iloc[2]
            valor_dataFinal = row.iloc[3]

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

            pyautogui.click(831,493)
            time.sleep(2)
            
            WebDriverWait(driver, 20)
            botao_extracao = wait.until(EC.element_to_be_clickable((By.ID, "frmMenuLateral:fieldExtracaoID")))
            botao_extracao.click()
            print("Botão 'Extração de documentos fiscais' clicado")
            time.sleep(1)

            pyperclip.copy(str(valor_dataInicio))

            #preencher campo data Inicio
            wait = WebDriverWait(driver,12)
            data_inicio = wait.until(EC.presence_of_element_located((By.ID, "FrmSolicitarExtracaoDfe:dtInicio_input")))
            data_inicio.clear()
            time.sleep(1)
            data_inicio.click()
            time.sleep(2)
            data_inicio.send_keys(Keys.CONTROL, 'v')
            print(f"DataInicio {valor_dataInicio} preenchida")
            time.sleep(2)

            pyperclip.copy(str(valor_dataFinal))

            #preencher campo data Final
            wait = WebDriverWait(driver, 12)
            data_final = wait.until(EC.presence_of_element_located((By.ID, "FrmSolicitarExtracaoDfe:dtFim_input")))
            data_final.clear()
            time.sleep(1)
            data_final.click()
            time.sleep(2)
            data_final.send_keys(Keys.CONTROL, 'v')
            print(f"DataFinal {valor_dataFinal} preenchida")
            time.sleep(1)

            # Espera até que o elemento esteja clicável
            wait = WebDriverWait(driver, 12)
            nfe_checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[@for='FrmSolicitarExtracaoDfe:tpDocumento:2']")))
            time.sleep(2)
            nfe_checkbox.click()
            time.sleep(1)
            print(f"nfe_checkbox preenchida")

            # Espera até que o elemento esteja clicável
            wait = WebDriverWait(driver, 12)
            destinatario_checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[@for='FrmSolicitarExtracaoDfe:tpParticipante:0']")))
            time.sleep(2)
            destinatario_checkbox.click()
            time.sleep(1)
            print(f"destinatario marcado")

            wait = WebDriverWait(driver, 12)
            confirmar_button = wait.until(EC.element_to_be_clickable((By.ID, "FrmSolicitarExtracaoDfe:submitPesquisa")))
            time.sleep(2)
            confirmar_button.click()
            time.sleep(1)
            print(f"Confirmar clicado")

            if verificar_mensagem_solicitacao(driver):

                wait = WebDriverWait(driver, 12)
                close_icon = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span.ui-icon.ui-icon-closethick")))
                time.sleep(2)
                close_icon.click()
                time.sleep(1)
                print("Clique fechar pop up")
            
            else:
                # Pausa a execução
                input("Mensagem não encontrada. Pressione Enter para continuar...")

                wait = WebDriverWait(driver, 3)
                close_icon = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span.ui-icon.ui-icon-closethick")))
                time.sleep(2)
                close_icon.click()
                time.sleep(1)
                print("Clique fechar pop up")
                
            wait = WebDriverWait(driver, 12)
            emitente_label = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[@for='FrmSolicitarExtracaoDfe:tpParticipante:1']")))
            time.sleep(2)
            emitente_label.click()
            time.sleep(1)
            print(f"emitente marcado")

            wait = WebDriverWait(driver, 12)
            confirmar_button = wait.until(EC.element_to_be_clickable((By.ID, "FrmSolicitarExtracaoDfe:submitPesquisa")))
            time.sleep(2)
            confirmar_button.click()
            time.sleep(1)
            print(f"Clique confirmar")
            
            if verificar_mensagem_solicitacao(driver):

                wait = WebDriverWait(driver, 12)
                close_icon = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span.ui-icon.ui-icon-closethick")))
                time.sleep(2)
                close_icon.click()
                time.sleep(1)
                print("Clique fechar pop up")

            else:
                # Pausa a execução
                input("Mensagem não encontrada. Pressione Enter para continuar...")
                
                wait = WebDriverWait(driver, 3)
                close_icon = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span.ui-icon.ui-icon-closethick")))
                time.sleep(2)
                close_icon.click()
                time.sleep(1)
                print("Clique fechar pop up")
                
            wait = WebDriverWait(driver, 12)
            nfce_label = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[@for='FrmSolicitarExtracaoDfe:tpDocumento:3']")))
            time.sleep(2)
            nfce_label.click()
            time.sleep(1)
            print(f"marcar Nfce")

            wait = WebDriverWait(driver, 12)
            emitente_label = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[@for='FrmSolicitarExtracaoDfe:tpParticipante:1']")))
            time.sleep(2)
            emitente_label.click()
            time.sleep(1)
            print(f"emitente marcado")

            wait = WebDriverWait(driver, 12)
            confirmar_button = wait.until(EC.element_to_be_clickable((By.ID, "FrmSolicitarExtracaoDfe:submitPesquisa")))
            time.sleep(2)
            confirmar_button.click()
            time.sleep(1)
            print(f"clique confirmar")
            
                
            if verificar_mensagem_solicitacao(driver):

                wait = WebDriverWait(driver, 12)
                close_icon = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span.ui-icon.ui-icon-closethick")))
                time.sleep(2)
                close_icon.click()
                time.sleep(1)
                print("Clique fechar pop up")

            else:
                #pausa a execução
                input("Mensagem não encontrada. Pressione Enter para continuar...")
                
                wait = WebDriverWait(driver, 3)
                close_icon = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span.ui-icon.ui-icon-closethick")))
                time.sleep(2)
                close_icon.click()
                time.sleep(1)
                print("Clique fechar pop up")
                
                
            wait = WebDriverWait(driver, 12)
            selecionar_estabelecimento = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@style, 'margin-left:0.5em') and text()='Selecionar estabelecimento']")))
            time.sleep(2)
            selecionar_estabelecimento.click()
            print(f"selecionar estabelecimento")
                
            time.sleep(1)  # Aguarde um tempo antes de passar para a próxima iteração

        print("Processo de preenchimento concluído.")
       

    finally:
        # Fecha o navegador
        driver.quit()
        print("Navegador fechado")

run_script()