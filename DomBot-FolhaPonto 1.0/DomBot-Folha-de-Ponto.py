# Importação de bibliotecas necessárias
import customtkinter as ctk          # Interface gráfica moderna e personalizada baseada em Tkinter
import pandas as pd                  # Manipulação de dados tabulares (Excel)
from pywinauto.application import Application  # Automação de interfaces Windows
from pywinauto.keyboard import send_keys      # Simulação de teclado
from pywinauto import findwindows, timings    # Encontrar janelas e controlar tempos de espera
import win32gui                      # API do Windows para manipulação de janelas
import win32con                      # Constantes do Windows
import time                          # Funções de tempo e espera
import logging                       # Registro de logs
from datetime import datetime        # Manipulação de datas e horas
import os                            # Operações do sistema operacional
from PIL import Image, ImageTk       # Manipulação de imagens

class AutomacaoGUI:
    def __init__(self):
        # Configuração do tema da interface - escuro com cores verdes
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")

        # Criação da janela principal do aplicativo
        self.window = ctk.CTk()
        self.window.title("DomBot - Folha de Ponto")
        self.window.geometry("700x500")  # Define o tamanho inicial da janela

       # Configura o ícone da janela
        self.set_window_icon()

        # Cria diretório para armazenar logs se não existir
        self.logs_dir = os.path.join(os.path.dirname(__file__), "logs")
        if not os.path.exists(self.logs_dir):
            os.makedirs(self.logs_dir)
        
        # Configura sistema de logging para arquivos
        self.setup_file_logging()

    def setup_file_logging(self):
        """Configura o logging para arquivos - criando logs separados para sucesso e erro"""
        data_atual = datetime.now().strftime("%Y-%m-%d")  # Data atual para nome dos arquivos
        
        # Logger para registros de sucesso
        self.success_logger = logging.getLogger('SuccessLog')
        self.success_logger.setLevel(logging.INFO)
        success_handler = logging.FileHandler(
            os.path.join(self.logs_dir, f'success_{data_atual}.log'),
            encoding='utf-8'
        )
        success_handler.setFormatter(
            logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S')
        )
        self.success_logger.addHandler(success_handler)
        
        # Logger para registros de erro
        self.error_logger = logging.getLogger('ErrorLog')
        self.error_logger.setLevel(logging.ERROR)
        error_handler = logging.FileHandler(
            os.path.join(self.logs_dir, f'error_{data_atual}.log'),
            encoding='utf-8'
        )
        error_handler.setFormatter(
            logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S')
        )
        self.error_logger.addHandler(error_handler)

    def set_window_icon(self):
        """Configura o ícone da janela principal do aplicativo"""
        try:
            # Caminho para o arquivo de ícone
            icon_path = os.path.join(os.path.dirname(__file__), "assets", "bot-folha-de-pagamento.ico")
            
            # Define o ícone apenas para Windows (nt)
            if os.name == 'nt':
                self.window.iconbitmap(icon_path)
                
        except Exception as e:
            print(f"Erro ao carregar ícone: {e}")
        
        # Variáveis para armazenar inputs do usuário
        self.arquivo_excel = ctk.StringVar()  # Caminho do arquivo Excel
        self.linha_inicial = ctk.StringVar(value="1")  # Linha inicial (padrão: 1)
        self.status_var = ctk.StringVar(value="Aguardando início...")  # Status da automação
        
        # Configuração do logger principal
        self.logger = logging.getLogger('AutomacaoDominio')
        self.logger.setLevel(logging.INFO)
        self.logger.handlers = []  # Limpa handlers existentes
        
        # Cria a interface gráfica
        self.criar_interface()
    
    def selecionar_arquivo(self):
        """Abre diálogo para seleção de arquivo Excel"""
        filename = ctk.filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Selecione o arquivo Excel"
        )
        if filename:
            self.arquivo_excel.set(filename)
            self.adicionar_log(f"Arquivo selecionado: {filename}")
        
    def criar_interface(self):
        """Cria todos os elementos da interface gráfica"""
        # Frame principal que contém todos os elementos
        main_frame = ctk.CTkFrame(self.window)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Frame superior para inputs do usuário
        input_frame = ctk.CTkFrame(main_frame)
        input_frame.pack(fill="x", padx=10, pady=10)
        
        # Área de seleção do arquivo Excel
        ctk.CTkLabel(input_frame, text="Arquivo Excel:").pack(anchor="w", padx=5, pady=2)
        
        file_frame = ctk.CTkFrame(input_frame)
        file_frame.pack(fill="x", pady=2)
        
        ctk.CTkEntry(file_frame, textvariable=self.arquivo_excel, width=400).pack(side="left", padx=5)
        ctk.CTkButton(file_frame, text="Procurar", command=self.selecionar_arquivo, width=100).pack(side="left", padx=5)
        
        # Seleção da linha inicial
        linha_frame = ctk.CTkFrame(input_frame)
        linha_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(linha_frame, text="Iniciar da linha:").pack(side="left", padx=5)
        ctk.CTkEntry(linha_frame, textvariable=self.linha_inicial, width=100).pack(side="left", padx=5)
        
        # Botão para iniciar a automação
        ctk.CTkButton(
            input_frame, 
            text="Iniciar", 
            command=self.iniciar_automacao,
            height=35
        ).pack(pady=15)
        
        # Barra de Progresso para mostrar avanço da automação
        self.progress_var = ctk.DoubleVar()
        self.progress_bar = ctk.CTkProgressBar(main_frame)
        self.progress_bar.pack(fill="x", padx=10, pady=10)
        self.progress_bar.set(0)  # Inicializa em 0%
        
        # Label para mostrar status atual
        ctk.CTkLabel(
            main_frame, 
            textvariable=self.status_var,
            height=25
        ).pack(pady=5)
        
        # Área de log onde são exibidas as mensagens em tempo real
        log_frame = ctk.CTkFrame(main_frame)
        log_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.log_text = ctk.CTkTextbox(log_frame, height=200)
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)

    def atualizar_progresso(self, atual, total):
        """Atualiza a barra de progresso e o texto de status"""
        porcentagem = (atual / total)
        self.progress_bar.set(porcentagem)
        self.status_var.set(f"Processando: {atual}/{total} ({porcentagem*100:.1f}%)")
        self.window.update()  # Força atualização da interface
            
    def adicionar_log(self, mensagem):
        """Adiciona uma mensagem na área de log com timestamp"""
        self.log_text.insert("end", f"{datetime.now().strftime('%H:%M:%S')} - {mensagem}\n")
        self.log_text.see("end")  # Rola para mostrar a mensagem mais recente
        self.window.update()  # Força atualização da interface
        
    def iniciar_automacao(self):
        """Função principal que inicia o processo de automação"""
        # Validações iniciais
        if not self.arquivo_excel.get():
            self.adicionar_log("Erro: Selecione um arquivo Excel")
            return
            
        try:
            linha_inicial = int(self.linha_inicial.get())
            if linha_inicial < 1:
                raise ValueError
        except ValueError:
            self.adicionar_log("Erro: Linha inicial inválida")
            return
            
        self.adicionar_log("Iniciando automação...")
        self.status_var.set("Em execução...")
        
        # Handler personalizado para redirecionar logs para a interface
        class GUIHandler(logging.Handler):
            def __init__(self, gui):
                super().__init__()
                self.gui = gui
                
            def emit(self, record):
                msg = self.format(record)
                self.gui.adicionar_log(msg)
        
        # Configura logging para mostrar na interface
        gui_handler = GUIHandler(self)
        formatter = logging.Formatter('%(message)s')
        gui_handler.setFormatter(formatter)
        self.logger.addHandler(gui_handler)
        
        try:
            # Carrega o arquivo Excel com pandas
            df = pd.read_excel(self.arquivo_excel.get())
            total_linhas = len(df) - (linha_inicial - 1)
            self.adicionar_log(f"Arquivo Excel carregado com {total_linhas} linhas para processar")
            
            # Inicializa a barra de progresso
            self.progress_var.set(0)
            
            # Cria objeto de automação
            automacao = DominioAutomation(self.logger)
            
            # Tenta conectar ao sistema Domínio
            if not automacao.connect_to_dominio():
                erro_msg = "Erro: Não foi possível conectar ao Domínio"
                self.adicionar_log(erro_msg)
                self.error_logger.error(erro_msg)
                return
                
            # Loop principal que processa cada linha do Excel
            for idx, (index, row) in enumerate(df.iloc[linha_inicial-1:].iterrows()):
                # Atualiza indicador de progresso
                self.atualizar_progresso(idx + 1, total_linhas)
                
                try:
                    # Processa uma linha do Excel
                    success = automacao.processar_linha(row, index)
                    
                    # Prepara mensagem de log
                    log_msg = (f"Linha {index + 1} - Nº {row['Nº']} - "
                            f"EMPRESA: {row.get('EMPRESA', 'N/A')}")
                    
                    # Registra resultado nos logs apropriados
                    if success:
                        self.success_logger.info(f"{log_msg} - Enviado com sucesso")
                        self.adicionar_log(f"Linha {index + 1} processada com sucesso")
                    else:
                        self.error_logger.error(f"{log_msg} - Erro no envio")
                        self.adicionar_log(f"Processo interrompido na linha {index + 1}")
                        break
                    
                    # Espera entre processamentos
                    time.sleep(2)
                    
                except Exception as e:
                    erro_msg = f"{log_msg} - Erro: {str(e)}"
                    self.error_logger.error(erro_msg)
                    self.adicionar_log(erro_msg)
                    break
            
            # Atualiza status ao finalizar
            self.status_var.set("Processamento concluído")
            self.progress_var.set(100)
            
        except Exception as e:
            # Tratamento de erros críticos
            erro_msg = f"Erro crítico: {str(e)}"
            self.error_logger.error(erro_msg)
            self.adicionar_log(erro_msg)
            self.status_var.set("Erro no processamento")
            
    def executar(self):
        """Inicia o loop principal da interface gráfica"""
        self.window.mainloop()

class DominioAutomation:
    """Classe que implementa a automação do sistema Domínio Folha"""
    def __init__(self, logger):
        # Aumenta timeout para encontrar janelas
        timings.Timings.window_find_timeout = 20
        self.app = None
        self.main_window = None
        self.logger = logger
        
    def log(self, message):
        """Registra mensagem no logger"""
        self.logger.info(message)

    def find_dominio_window(self):
        """Procura pela janela do Domínio Folha usando expressão regular no título"""
        try:
            windows = findwindows.find_windows(title_re=".*Domínio Folha.*")
            if windows:
                return windows[0]  # Retorna o handle da primeira janela encontrada
            self.log("Nenhuma janela do Domínio Folha encontrada.")
            return None
        except Exception as e:
            self.log(f"Erro ao procurar a janela do Domínio Folha: {str(e)}")
            return None

    def connect_to_dominio(self):
        """Conecta à janela do sistema Domínio já aberto"""
        try:
            handle = self.find_dominio_window()
            if not handle:
                self.log("Não foi possível encontrar a janela do Domínio Folha.")
                return False

            # Restaura a janela se estiver minimizada
            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                time.sleep(1)

            # Conecta à aplicação usando interface UIA
            self.app = Application(backend="uia").connect(handle=handle)
            self.main_window = self.app.window(handle=handle)
            return True
        except Exception as e:
            self.log(f"Erro ao conectar ao Domínio Folha: {str(e)}")
            return False

    def wait_for_window(self, titulo, timeout=30):
        """Espera por uma janela com o título especificado até um timeout"""
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # Tenta encontrar a janela pelo título
                window = self.app.window(title=titulo)
                if window.exists():
                    return window
            except Exception:
                pass
            time.sleep(0.5)
        raise TimeoutError(f"Timeout esperando pela janela: {titulo}")

    def processar_linha(self, row, index):
        """Processa uma linha do Excel, interagindo com o sistema Domínio"""
        try:
            self.log(f"Processando linha {index + 1}")
        
            # Reconecta à janela principal (garante que temos controle da janela correta)
            handle = self.find_dominio_window()
            if not handle:
                self.log("Não foi possível encontrar a janela do Domínio Folha.")
                return False

            # Restaura a janela se estiver minimizada
            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                time.sleep(1)

            app = Application(backend="uia").connect(handle=handle)
            main_window = app.window(handle=handle)
        
            # Garante que a janela principal está em foco
            main_window.set_focus()
            time.sleep(1)
            
            # Pressiona F8 para trocar de empresa
            send_keys('{F8}')
            time.sleep(2)

            # Busca a janela "Troca de empresas"
            try:
                # Tenta encontrar usando critérios específicos
                troca_empresas_window = main_window.child_window(
                    title="Troca de empresas",
                    class_name="FNWND3190"  # Classe da janela
                )
                
                # Alternativa caso não encontre com os critérios específicos
                if not troca_empresas_window.exists():
                    troca_empresas_windows = main_window.children(title="Troca de empresas")
                    if troca_empresas_windows:
                        troca_empresas_window = troca_empresas_windows[0]
                    else:
                        self.log("Janela 'Troca de empresas' não encontrada.")
                        return False
            except Exception as e:
                self.log(f"Erro ao localizar janela 'Troca de empresas': {str(e)}")
                return False

            self.log("Janela 'Troca de empresas' visível")

            # Digita o número da empresa
            send_keys(f"{int(row['Nº'])}")  # Envia o valor da coluna Nº
            time.sleep(0.5)

            # Confirma com Enter
            send_keys('{ENTER}')
            time.sleep(9)  # Espera a troca de empresa

            # Aguarda até que a janela "Troca de empresas" feche ou timeout
            max_wait = 25  # tempo máximo de espera em segundos
            wait_interval = 0.5  # intervalo entre verificações
            start_time = time.time()

            while time.time() - start_time < max_wait:
                try:
                    # Verifica se a janela ainda está visível
                    troca_empresas_window = main_window.child_window(
                    title="Troca de empresas",
                    class_name="FNWND3190"  
                )
                    if not troca_empresas_window.exists() or not troca_empresas_window.is_visible():
                        # Janela fechou com sucesso
                        self.log("Janela de troca de empresas fechada com sucesso")
                        
                        # Verifica se há janela de avisos de vencimento
                        aviso_window = main_window.child_window(
                            title="Avisos de Vencimento",
                            class_name="FNWND3190"
                        )
                        
                        if aviso_window.exists() and aviso_window.is_visible():
                            self.log("Janela 'Avisos de Vencimento' encontrada - executando fechamento")
                            send_keys('{ESC}')
                            time.sleep(5)  # Espera entre ESCs
                            send_keys('{ESC}')
                            self.log("ESCs executados para fechar 'Avisos de Vencimento'")
                        else:
                            self.log("Janela 'Avisos de Vencimento' não está visível - continuando")
                        break
                except Exception:
                    # Janela não encontrada (já fechou)
                    self.log("Janela de troca de empresas fechada com sucesso")
                    
                    # Verifica se há janela de avisos de vencimento
                    aviso_window = main_window.child_window(
                        title="Avisos de Vencimento",
                        class_name="FNWND3190"
                    )
                    
                    if aviso_window.exists() and aviso_window.is_visible():
                        self.log("Janela 'Avisos de Vencimento' encontrada - executando fechamento")
                        send_keys('{ESC}')
                        time.sleep(5)  # Espera entre ESCs
                        send_keys('{ESC}')
                        self.log("ESCs executados para fechar 'Avisos de Vencimento'")
                    else:
                        self.log("Janela 'Avisos de Vencimento' não está visível - continuando")
                    break
                time.sleep(2)
            else:
                # Timeout atingido
                self.log("Aviso: Tempo máximo de espera atingido para fechamento da janela")

            # Navega até o menu de relatórios
            self.main_window.set_focus()
            send_keys('%r')  # ALT+R para abrir menu Relatórios
            time.sleep(1)
            send_keys('i')   # Seleção do submenu
            time.sleep(1)
            send_keys('i')   # Segunda seleção
            time.sleep(1)
            send_keys('{ENTER}')  # Confirma seleção

            # Interage com o Gerenciador de Relatórios
            try:
                # Encontra a janela do Gerenciador 
                relatorio_window = main_window.child_window(
                    title="Gerenciador de Relatórios",
                    class_name="FNWND3190"
                )
                
                if not relatorio_window.exists():
                    self.log("Janela 'Gerenciador de Relatórios' não encontrada.")
                    return False

                # Localiza elementos da árvore de relatórios
                app = Application(backend='uia').connect(handle=relatorio_window.handle)
                tree = app.window(class_name="FNWND3190").child_window(class_name="PBTreeView32_100")
                
                # Navega para selecionar o relatório correto
                try:
                    folha_21_20 = tree.child_window(title="Folha de Ponto_21 a 20 - II")
                    folha_ponto = tree.child_window(title="Folha - Ponto")

                    if folha_ponto.exists():
                        folha_ponto.set_focus()
                        folha_ponto.click_input(double=True)  # Duplo clique para expandir
                        time.sleep(0.5)

                        folha_21_20.set_focus()
                    else:
                        self.log("Item 'Folha - Ponto' não encontrado na árvore.")
                        return False
                    
                except Exception as e:
                    self.log(f"Erro ao clicar no item: {str(e)}")
                    return False

            except Exception as e:
                self.log(f"Erro ao interagir com o Gerenciador de Relatórios: {str(e)}")
                self.log(f"Detalhes do erro: {str(e.__class__.__name__)}")
                return False

            # Preenche campos de datas com TAB
            send_keys('{TAB}*')  # Move para o próximo campo
            send_keys('{TAB}' + str(row['data inicio']))  # Preenche data de início
            send_keys('{TAB}' + str(row['data final']))   # Preenche data final

            time.sleep(1)

            # Clica no botão de executar pelo ID
            button_1007 = relatorio_window.child_window(auto_id="1007", class_name="Button")
            button_1007.click_input()

            time.sleep(2)

            # Clica no botão da maleta (para salvar)
            self.main_window.set_focus()
            button_1005 = self.main_window.child_window(auto_id="1005", class_name="FNUDO3190")
            button_1005.click_input()

            try:
                # Encontra a janela de publicação de documentos
                pub_doc_window = main_window.child_window(
                    title="Publicação de Documentos",
                    class_name="FNWNS3190"
                )

                # Seleciona a pasta de destino
                combo_box = pub_doc_window.child_window(auto_id="1007", class_name="ComboBox")
                combo_box.click_input()
                send_keys("Pessoal/Folha de Ponto{ENTER}")  # Seleciona pasta específica

                time.sleep(1)

                # Define o nome do arquivo PDF
                edit_field = pub_doc_window.child_window(auto_id="1014", class_name="Edit")
                edit_field.set_text(str(row['nome pdf']))  # Usa valor da coluna "nome pdf"

                time.sleep(1)

                # Comentado: ativa/desativa checkbox de pasta padrão
                # check_box = pub_doc_window.child_window(auto_id="1008", class_name="Button")
                # check_box.click_input()

                # Clica no botão Gravar
                button_1016 = pub_doc_window.child_window(auto_id="1016", class_name="Button")
                button_1016.click()

                # Aguarda fechamento da janela de publicação
                max_wait = 25  # tempo máximo de espera
                wait_interval = 0.5
                start_time = time.time()

                # Loop que verifica se a janela ainda está visível
                while time.time() - start_time < max_wait:
                    try:
                        pub_doc_window = main_window.child_window(
                            title="Publicação de Documentos",
                            class_name="FNWNS3190"
                        )
                        if not pub_doc_window.exists() or not pub_doc_window.is_visible():
                            self.log("Janela de publicação fechada com sucesso")
                            break
                    except Exception:
                        self.log("Janela de publicação fechada com sucesso")
                        break
                    time.sleep(wait_interval)
                else:
                    self.log("Aviso: Tempo máximo de espera atingido para fechamento da janela")

                # Limpa janelas residuais com ESC
                send_keys('{ESC}')
                time.sleep(2)
                
                # Garante foco na janela principal
                self.main_window.set_focus()
                time.sleep(1)
                
                # Segundo ESC para garantir
                send_keys('{ESC}')
                
                # Verifica se ainda existe alguma janela aberta que deveria estar fechada
                max_wait = 10
                start_time = time.time()
                
                # Loop para tentar fechar qualquer janela residual
                while time.time() - start_time < max_wait:
                    try:
                        relatorio_window = main_window.child_window(
                            title="Gerenciador de Relatórios",
                            class_name="FNWND3190"
                        )
                        
                        if relatorio_window.exists() and relatorio_window.is_visible():
                            # Tenta fechar novamente
                            self.main_window.set_focus()
                            time.sleep(0.5)
                            send_keys('{ESC}')
                            time.sleep(1)
                        else:
                            self.log("Todas as janelas fechadas com sucesso")
                            break
                            
                    except Exception:
                        # Janela não encontrada, provavelmente fechada
                        break
                        
                    time.sleep(0.5)
                else:
                    self.log("Aviso: Possível janela ainda aberta após tentativas de fechamento")

            except Exception as e:
                self.log(f"Erro ao interagir com o Gerenciador de Relatórios: {str(e)}")
                self.log(f"Detalhes do erro: {str(e.__class__.__name__)}")
                return False
            
            # Processo concluído com sucesso
            self.log(f"Linha {index + 1} processada com sucesso")
            return True

        except Exception as e:
            # Captura qualquer erro na execução
            self.log(f"Erro ao processar linha {index + 1}: {str(e)}")
            return False

# Função principal original - executada diretamente, não pela GUI
def main():
    try:
        # Carrega Excel diretamente
        df = pd.read_excel('automação folha de ponto.xlsx')
        logging.info(f"Arquivo Excel carregado com {len(df)} linhas")

        # Inicia automação sem interface
        automacao = DominioAutomation()
        
        # Conecta ao Domínio
        if not automacao.connect_to_dominio():
            logging.error("Não foi possível conectar ao Domínio Folha")
            return

        # Processa cada linha do Excel
        for index, row in df.iterrows():
            success = automacao.processar_linha(row, index)
            if not success:
                logging.warning(f"Parando o processamento na linha {index + 1}")
                break
            time.sleep(2)

    except Exception as e:
        logging.error(f"Erro crítico na execução: {str(e)}")

# Função main substituída para iniciar a interface gráfica
def main():
    gui = AutomacaoGUI()
    gui.executar()

# Ponto de entrada do programa
if __name__ == "__main__":
    main()