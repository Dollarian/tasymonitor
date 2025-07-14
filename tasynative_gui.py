import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog
import threading
import time
import os
import subprocess
import socket
import psutil
import configparser
import sys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException

# Importações para lidar com atalhos do Windows
import pythoncom
import win32com.client

# --- CONFIGURAÇÕES DO ARQUIVO ---g
CONFIG_FILE = "config.ini"
LOG_FILE = "monitoramento_log.txt"

# Variáveis globais
rodando = False
driver = None
monitor_thread = None

# Dicionário para armazenar as configurações carregadas
app_settings = {
    "ATALHO_TASY_PATH": r"C:\Program Files\Philips EMR\TasyNative\TasyNative.lnk", # Caminho padrão para o atalho principal
    "ATALHO_TASY_PATH_FALLBACK": r"C:\Program Files\Philips EMR\tasy-native\tasy-native.lnk", # Caminho de fallback
    "TASY_NATIVE_INSTALL_PATHS": "C:\\Program Files\\Philips EMR\\TasyNative,C:\\Program Files\\Philips EMR\\tasy-native", # Caminhos para buscar o TasyNative.exe
    "TASY_URL": "https://tasyweb.spdm.org.br/",
    "USERNAME": "usr_spaa",
    "PASSWORD": "hospital@102030",
    "REMOTE_DEBUGGING_PORT": 9222,
    "BUNDLED_CHROMEDRIVER_RELATIVE_PATH": "chromedriver.exe"
}

# --- FUNÇÕES DE CONFIGURAÇÃO ---
def load_settings():
    """Carrega as configurações do arquivo config.ini."""
    config = configparser.ConfigParser()
    if os.path.exists(CONFIG_FILE):
        config.read(CONFIG_FILE)
        if 'SETTINGS' in config:
            for key in app_settings:
                if key in config['SETTINGS']:
                    if key == "REMOTE_DEBUGGING_PORT":
                        try:
                            app_settings[key] = int(config['SETTINGS'][key])
                        except ValueError:
                            escrever_log(f"Aviso: Valor inválido para {key} no config.ini. Usando padrão.", None)
                    elif key == "TASY_NATIVE_INSTALL_PATHS":
                        app_settings[key] = [p.strip() for p in config['SETTINGS'][key].split(',') if p.strip()]
                    else:
                        app_settings[key] = config['SETTINGS'][key]
        escrever_log(f"Configurações carregadas de {CONFIG_FILE}")
    else:
        escrever_log(f"Arquivo de configuração '{CONFIG_FILE}' não encontrado. Usando valores padrão.")
    return app_settings

def save_settings(settings_to_save):
    """Salva as configurações no arquivo config.ini."""
    config = configparser.ConfigParser()
    settings_to_save_str = {k: str(v) for k, v in settings_to_save.items()}
    if isinstance(settings_to_save["TASY_NATIVE_INSTALL_PATHS"], list):
        settings_to_save_str["TASY_NATIVE_INSTALL_PATHS"] = ",".join(settings_to_save["TASY_NATIVE_INSTALL_PATHS"])

    config['SETTINGS'] = settings_to_save_str
    with open(CONFIG_FILE, 'w') as configfile:
        config.write(configfile)
    escrever_log(f"Configurações salvas em {CONFIG_FILE}")

# --- LOG ---
def escrever_log(msg, console=None):
    """Escreve uma mensagem no console (se fornecido) e no arquivo de log."""
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    linha = f"[{timestamp}] {msg}"
    print(linha)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(linha + "\n")
    if console:
        console.configure(state='normal')
        console.insert(tk.END, linha + "\n")
        console.see(tk.END)
        console.configure(state='disabled')

# --- FUNÇÕES DE GERENCIAMENTO DE PROCESSO ---
def encerrar_chrome_debugger(port, console):
    """
    Encerra processos do Chrome que estão usando a porta de depuração remota especificada.
    Também encerra qualquer processo que esteja escutando na porta.
    """
    escrever_log(f"Tentando encerrar processos do Chrome ou outros na porta {port}...", console)
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            if proc.info['name'] == "chrome.exe" or proc.info['name'] == "chromium":
                cmdline = " ".join(proc.info['cmdline'])
                if f"--remote-debugging-port={port}" in cmdline:
                    escrever_log(f"Encerrando Chrome com debugger (PID {proc.pid}) na porta {port})", console)
                    proc.kill()
                    continue

            if hasattr(proc, 'connections'):
                try:
                    for conn in proc.connections(kind='inet'):
                        if conn.laddr.port == port and conn.status == psutil.CONN_LISTEN:
                            escrever_log(f"Encerrando processo '{proc.name()}' (PID {proc.pid}) escutando na porta {port}", console)
                            proc.kill()
                            break
                except psutil.AccessDenied:
                    escrever_log(f"Acesso negado ao verificar conexões para o processo {proc.pid} ({proc.name()}).", console)
                except Exception as conn_e:
                    escrever_log(f"Erro ao verificar conexões para o processo {proc.pid} ({proc.name()}): {conn_e}", console)

        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
        except Exception as e:
            escrever_log(f"Erro ao analisar processo {proc.pid}: {e}", console)

def porta_esta_ocupada(porta):
    """Verifica se uma porta está em uso."""
    for conn in psutil.net_connections(kind='inet'):
        if conn.laddr.port == porta:
            return True
    return False

# --- ABRIR ATALHO DO TASY NATIVE ---
def abrir_atalho(atalho_path_main, atalho_path_fallback, tasy_native_install_paths, console):
    """
    Abre um atalho .lnk do Windows. Tenta o caminho principal, depois o fallback.
    Se nenhum atalho existente for encontrado, tenta criar um novo.
    Esta função é específica para Windows.
    """
    if os.name != 'nt':
        escrever_log("⚠️ Abrir atalhos .lnk só funciona no Windows. Esta funcionalidade pode não funcionar em outros sistemas.", console)
        raise Exception("Funcionalidade de atalho .lnk não suportada neste sistema operacional.")

    caminhos_atalho_existentes = [atalho_path_main]
    if atalho_path_fallback and atalho_path_fallback != atalho_path_main:
        caminhos_atalho_existentes.append(atalho_path_fallback)

    escrever_log(f"Caminhos de atalho existentes a tentar: {caminhos_atalho_existentes}", console)

    # Tenta abrir atalhos existentes primeiro
    for path in caminhos_atalho_existentes:
        escrever_log(f"Verificando atalho existente em: '{path}'", console)
        if os.path.exists(path):
            escrever_log(f"Atalho encontrado em '{path}'. Tentando abrir...", console)
            try:
                pythoncom.CoInitialize()
                shell = win32com.client.Dispatch("WScript.Shell")
                atalho = shell.CreateShortCut(path)
                
                target_path = atalho.TargetPath
                arguments = atalho.Arguments
                working_directory = atalho.WorkingDirectory
                
                escrever_log(f"Propriedades do atalho: TargetPath='{target_path}', Arguments='{arguments}', WorkingDirectory='{working_directory}'", console)
                
                # NOVO: Adiciona o argumento de depuração remota se não estiver presente
                # Garante que os argumentos originais sejam preservados e o novo adicionado
                if f"--remote-debugging-port={app_settings['REMOTE_DEBUGGING_PORT']}" not in arguments:
                    arguments = f"{arguments} --remote-debugging-port={app_settings['REMOTE_DEBUGGING_PORT']}".strip()
                    escrever_log(f"Adicionado argumento de depuração remota ao atalho existente. Novos argumentos: '{arguments}'", console)

                # NOVO: Usa subprocess.Popen com cwd para garantir o diretório de trabalho correto
                # e loga o comando final e cwd
                command_to_execute = f'"{target_path}" {arguments}'
                escrever_log(f"Comando final para subprocess.Popen: '{command_to_execute}'", console)
                escrever_log(f"Diretório de trabalho (cwd) para subprocess.Popen: '{working_directory}'", console)

                subprocess.Popen(command_to_execute, cwd=working_directory)
                escrever_log(f"Atalho do Tasy Native aberto com sucesso de '{path}'.", console)
                return # Sai da função se o atalho foi aberto com sucesso
            except Exception as e:
                escrever_log(f"Erro ao abrir atalho existente de '{path}': {e}", console)
            finally:
                pythoncom.CoUninitialize()
        else:
            escrever_log(f"Atalho não encontrado em '{path}'.", console)

    # Se nenhum atalho existente funcionou, tenta criar um novo
    escrever_log("Nenhum atalho existente funcionou. Tentando criar um novo atalho...", console)
    
    tasy_native_exe_name = "TasyNative.exe"
    created_shortcut_path = None

    for base_path in tasy_native_install_paths:
        exe_path = os.path.join(base_path, tasy_native_exe_name)
        escrever_log(f"Verificando executável do Tasy Native em: '{exe_path}'", console)
        if os.path.exists(exe_path):
            escrever_log(f"Executável do Tasy Native encontrado em '{exe_path}'.", console)
            
            # Determine where to create the new shortcut
            if getattr(sys, 'frozen', False):
                # If running as PyInstaller executable, create in the same directory as the .exe
                shortcut_dir = os.path.dirname(sys.executable)
            else:
                # If running from source, create in the script's directory
                shortcut_dir = os.path.dirname(os.path.abspath(__file__))
            
            new_shortcut_name = "TasyNative_Auto.lnk"
            created_shortcut_path = os.path.join(shortcut_dir, new_shortcut_name)

            escrever_log(f"Tentando criar novo atalho em: '{created_shortcut_path}'", console)
            try:
                pythoncom.CoInitialize()
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(created_shortcut_path)
                shortcut.TargetPath = exe_path
                # Inclui o argumento de depuração remota
                shortcut.Arguments = f"--remote-debugging-port={app_settings['REMOTE_DEBUGGING_PORT']}"
                shortcut.WindowStyle = 1 # 1 for normal window
                shortcut.Description = "Atalho automático para Tasy Native com depuração remota"
                shortcut.WorkingDirectory = os.path.dirname(exe_path) # Define o diretório de trabalho para a pasta do TasyNative
                shortcut.Save()
                escrever_log(f"Novo atalho criado com sucesso em '{created_shortcut_path}'.", console)
                
                # Agora tenta abrir o atalho recém-criado
                # Loga o comando final e cwd também para o atalho criado
                command_to_execute = f'"{shortcut.TargetPath}" {shortcut.Arguments}'
                escrever_log(f"Comando final para subprocess.Popen (novo atalho): '{command_to_execute}'", console)
                escrever_log(f"Diretório de trabalho (cwd) para subprocess.Popen (novo atalho): '{shortcut.WorkingDirectory}'", console)

                subprocess.Popen(command_to_execute, cwd=shortcut.WorkingDirectory)
                escrever_log(f"Novo atalho do Tasy Native aberto com sucesso de '{created_shortcut_path}'.", console)
                return # Sai da função se o novo atalho foi criado e aberto com sucesso
            except Exception as e:
                escrever_log(f"Erro ao criar ou abrir o novo atalho em '{created_shortcut_path}': {e}", console)
            finally:
                pythoncom.CoUninitialize()
        else:
            escrever_log(f"Executável do Tasy Native não encontrado em '{exe_path}'.", console)

    # Se chegou até aqui, nenhuma opção funcionou
    raise Exception("Falha ao abrir ou criar atalho do Tasy Native: Nenhuma opção funcionou ou o atalho está corrompido.")


# --- CONECTAR CHROME (com ChromeDriver empacotado) ---
def setup_driver(port, bundled_chromedriver_relative_path, console):
    """Configura e retorna o driver do Selenium usando o ChromeDriver empacotado."""
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")

    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
        escrever_log(f"Rodando como executável PyInstaller. Base path: {base_path}", console)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
        escrever_log(f"Rodando em ambiente de desenvolvimento. Base path: {base_path}", console)

    chromedriver_path = os.path.join(base_path, bundled_chromedriver_relative_path)

    escrever_log(f"Caminho do ChromeDriver a ser usado: '{chromedriver_path}'", console)
    escrever_log(f"Verificando existência do ChromeDriver em '{chromedriver_path}': {os.path.exists(chromedriver_path)}", console)

    if not os.path.exists(chromedriver_path):
        escrever_log(f"Erro: ChromeDriver empacotado não encontrado em '{chromedriver_path}'. Verifique o caminho relativo nas configurações e a inclusão no PyInstaller.", console)
        raise Exception("ChromeDriver empacotado não encontrado.")

    escrever_log(f"Usando ChromeDriver empacotado de: {chromedriver_path}", console)
    service = Service(executable_path=chromedriver_path)

    try:
        return webdriver.Chrome(service=service, options=chrome_options)
    except Exception as e:
        if "session not created: This version of ChromeDriver only supports Chrome version" in str(e):
            escrever_log(f"ERRO DE COMPATIBILIDADE: A versão do ChromeDriver ({os.path.basename(chromedriver_path)}) não é compatível com a versão do Chrome que o Tasy Native abriu. Detalhes: {e}", console)
            raise Exception(f"Incompatibilidade de versão entre ChromeDriver e Chrome. Por favor, baixe o ChromeDriver correto para a versão do Chrome que o Tasy Native utiliza. Erro: {e}")
        else:
            escrever_log(f"Erro ao inicializar o driver do Selenium com ChromeDriver empacotado: {e}. Verifique a compatibilidade entre o Chrome e o ChromeDriver, e se o Chrome foi iniciado pelo Tasy Native.", console)
            raise Exception("Falha ao configurar ChromeDriver empacotado.")


def login_tasy(driver_instance, url, username, password, console):
    """Realiza o login no Tasy."""
    driver_instance.get(url)
    escrever_log(f"URL carregada: {url}", console)

    try:
        WebDriverWait(driver_instance, 20).until(EC.visibility_of_element_located((By.ID, "loginUsername"))).send_keys(username)
        WebDriverWait(driver_instance, 20).until(EC.visibility_of_element_located((By.ID, "loginPassword"))).send_keys(password)
        WebDriverWait(driver_instance, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' and @value='Entrar']"))).click()
        escrever_log("Login realizado com sucesso. Aguardando carregamento da página pós-login...", console)

        try:
            WebDriverWait(driver_instance, 30).until(EC.url_changes(url))
            escrever_log("URL alterada. Navegação para página pós-login detectada.", console)
        except Exception as e:
            escrever_log(f"⚠️ A URL não mudou da página de login após 30 segundos. Login pode ter falhado ou a navegação está lenta: {e}", console)
            raise Exception("Login não sustentado ou navegação para página pós-login falhou.")

        try:
            WebDriverWait(driver_instance, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            escrever_log("Elemento 'body' presente na nova página. Login provavelmente bem-sucedido.", console)
        except Exception as e:
            escrever_log(f"⚠️ Elemento 'body' não detectado na nova página após 15 segundos. A página pode não ter carregado corretamente: {e}", console)
            raise Exception("Página pós-login não carregada completamente.")

    except Exception as e:
        escrever_log(f"Erro durante o processo de login inicial ou elementos não encontrados: {e}", console)
        raise Exception("Falha no login ou elementos não encontrados.")
    
    escrever_log("Aguardando 10 segundos antes de verificar erros na sessão e fechar popups...", console)
    time.sleep(10)

    try:
        btn_fechar = WebDriverWait(driver_instance, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[span[normalize-space(text())='Fechar']]"))
        )
        btn_fechar.click()
        escrever_log("Popup de boas-vindas/informativo fechado.", console)
    except:
        escrever_log("Popup não encontrado ou já fechado.", console)
    


def is_logged_out(driver_instance, console):
    """Verifica se a sessão foi deslogada ou se há erros na tela."""
    try:
        if driver_instance.find_elements(By.ID, "loginUsername"):
            escrever_log("⚠️ Sessão deslogada (campo de login visível).", console)
            return True

        error_keywords = ["Erro", "SQL", "falha", "expirada", "session expired", "error"]
        for keyword in error_keywords:
            if driver_instance.find_elements(By.XPATH, f"//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{keyword.lower()}')]"):
                escrever_log(f"⚠️ Erro detectado na tela com a palavra-chave: '{keyword}'.", console)
                return True

        return False
    except WebDriverException as e:
        escrever_log(f"Erro ao verificar sessão (WebDriverException): {e}. Presumindo que a sessão está instável e requer reinício.", console)
        return True
    except Exception as e:
        escrever_log(f"Erro ao verificar sessão: {e}. Presumindo que a sessão está instável.", console)
        return True

# --- LOOP PRINCIPAL DE MONITORAMENTO ---
def monitorar(console):
    """Função principal que gerencia o ciclo de vida da automação."""
    global rodando, driver
    settings = app_settings

    while rodando:
        driver_inicializado = False
        try:
            escrever_log("Verificando e liberando porta de depuração...", console)
            encerrar_chrome_debugger(settings["REMOTE_DEBUGGING_PORT"], console)
            time.sleep(2)

            escrever_log("Abrindo TasyNative...", console)
            abrir_atalho(settings["ATALHO_TASY_PATH"], settings["ATALHO_TASY_PATH_FALLBACK"], settings["TASY_NATIVE_INSTALL_PATHS"], console)
            
            escrever_log(f"Aguardando porta {settings['REMOTE_DEBUGGING_PORT']} ficar disponível após lançamento do Tasy Native...", console)
            max_wait_for_port = 60 # Tempo máximo de espera em segundos
            waited_time_for_port = 0
            while not porta_esta_ocupada(settings["REMOTE_DEBUGGING_PORT"]) and waited_time_for_port < max_wait_for_port:
                escrever_log(f"Porta {settings['REMOTE_DEBUGGING_PORT']} ainda não está ocupada. Aguardando... ({waited_time_for_port}s/{max_wait_for_port}s)", console)
                time.sleep(5)
                waited_time_for_port += 5
            
            if not porta_esta_ocupada(settings["REMOTE_DEBUGGING_PORT"]):
                raise Exception(f"Porta {settings['REMOTE_DEBUGGING_PORT']} não ficou disponível após {max_wait_for_port} segundos. O Tasy Native pode não ter iniciado o Chrome com a depuração remota ou o Chrome não está respondendo.")
            
            escrever_log(f"Porta {settings['REMOTE_DEBUGGING_PORT']} está ocupada. Conectando ao Chrome via porta de depuração...", console)
            driver = setup_driver(settings["REMOTE_DEBUGGING_PORT"], settings["BUNDLED_CHROMEDRIVER_RELATIVE_PATH"], console)
            driver_inicializado = True

            login_tasy(driver, settings["TASY_URL"], settings["USERNAME"], settings["PASSWORD"], console)

            while rodando:
                if is_logged_out(driver, console):
                    escrever_log("Sessão deslogada ou erro detectado. Reiniciando automação...", console)
                    break
                escrever_log("🟢 Sessão ativa.", console)
                time.sleep(30)

        except Exception as e:
            escrever_log(f"Erro crítico na automação: {e}. Tentando reiniciar...", console)
            if driver_inicializado and driver:
                try:
                    driver.quit()
                    driver = None
                    escrever_log("Driver do Selenium fechado após erro.", console)
                except Exception as quit_e:
                    escrever_log(f"Erro ao tentar fechar o driver do Selenium: {quit_e}", console)
            
            try:
                encerrar_chrome_debugger(settings["REMOTE_DEBUGGING_PORT"], console)
                escrever_log("Processos do Chrome e depurador encerrados após erro.", console)
            except Exception as cleanup_e:
                escrever_log(f"Erro durante a limpeza de processos do Chrome: {cleanup_e}", console)
            
            if not rodando:
                escrever_log("Sinal de parada recebido. Encerrando o ciclo de monitoramento.", console)
                break
            time.sleep(15)

    escrever_log("Monitoramento encerrado.", console)
    try:
        if driver:
            driver.quit()
            driver = None
        encerrar_chrome_debugger(settings["REMOTE_DEBUGGING_PORT"], console)
    except Exception as final_cleanup_e:
        escrever_log(f"Erro durante a limpeza final: {final_cleanup_e}", console)

# --- GUI CALLBACKS ---
def abrir_configuracoes(janela_principal, console):
    """Abre uma nova janela para configurações."""
    config_window = tk.Toplevel(janela_principal)
    config_window.title("Configurações")
    config_window.geometry("600x600")

    labels = [
        "Caminho do Atalho TasyNative (.lnk):",
        "Caminho do Atalho TasyNative (Fallback):",
        "Caminhos de Instalação do Tasy Native (separados por vírgula):",
        "URL do Tasy:",
        "Usuário:",
        "Senha:",
        "Porta de Depuração Remota:",
        "Caminho Relativo ChromeDriver Empacotado:"
    ]
    keys = [
        "ATALHO_TASY_PATH",
        "ATALHO_TASY_PATH_FALLBACK",
        "TASY_NATIVE_INSTALL_PATHS",
        "TASY_URL",
        "USERNAME",
        "PASSWORD",
        "REMOTE_DEBUGGING_PORT",
        "BUNDLED_CHROMEDRIVER_RELATIVE_PATH"
    ]
    entries = {}

    for i, label_text in enumerate(labels):
        tk.Label(config_window, text=label_text).grid(row=i, column=0, padx=5, pady=2, sticky='w')
        
        if keys[i] == "PASSWORD":
            entry = tk.Entry(config_window, width=50, show='*')
        else:
            entry = tk.Entry(config_window, width=50)
        
        entry.grid(row=i, column=1, padx=5, pady=2, sticky='ew')
        entries[keys[i]] = entry

        if keys[i] == "TASY_NATIVE_INSTALL_PATHS":
            entry.insert(0, ",".join(app_settings.get(keys[i], [])))
        else:
            entry.insert(0, str(app_settings.get(keys[i], "")))

        if keys[i] in ["ATALHO_TASY_PATH", "ATALHO_TASY_PATH_FALLBACK"]:
            btn_browse = tk.Button(config_window, text="Procurar...", command=lambda e=entry: browse_lnk_path(e))
            btn_browse.grid(row=i, column=2, padx=5, pady=2)
        elif keys[i] == "TASY_NATIVE_INSTALL_PATHS":
            btn_browse = tk.Button(config_window, text="Adicionar Pasta...", command=lambda e=entry: add_install_path(e))
            btn_browse.grid(row=i, column=2, padx=5, pady=2)


    def browse_lnk_path(entry_widget):
        file_path = filedialog.askopenfilename(
            title="Selecione o Atalho do Tasy Native (.lnk)",
            filetypes=[("Atalhos do Windows", "*.lnk"), ("Todos os arquivos", "*.*")]
        )
        if file_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, file_path)

    def add_install_path(entry_widget):
        folder_path = filedialog.askdirectory(
            title="Selecione a Pasta de Instalação do Tasy Native"
        )
        if folder_path:
            current_paths = entry_widget.get()
            if current_paths:
                new_paths = f"{current_paths},{folder_path}"
            else:
                new_paths = folder_path
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, new_paths)


    def salvar_e_fechar():
        """Salva as configurações e fecha o janela."""
        new_settings = {}
        for key, entry_widget in entries.items():
            value = entry_widget.get()
            if key == "REMOTE_DEBUGGING_PORT":
                try:
                    new_settings[key] = int(value)
                except ValueError:
                    messagebox.showerror("Erro de Configuração", "A porta de depuração deve ser um número inteiro válido.")
                    return
            elif key == "TASY_NATIVE_INSTALL_PATHS":
                new_settings[key] = [p.strip() for p in value.split(',') if p.strip()]
            else:
                new_settings[key] = value

        app_settings.update(new_settings)
        save_settings(app_settings)
        escrever_log("Configurações atualizadas e salvas.", console)
        config_window.destroy()

    btn_salvar = tk.Button(config_window, text="Salvar e Fechar", command=salvar_e_fechar)
    btn_salvar.grid(row=len(labels), column=0, columnspan=3, pady=10)

    config_window.grid_columnconfigure(1, weight=1)

def iniciar(console, start_button, stop_button):
    """Inicia o monitoramento em uma nova thread."""
    global rodando, monitor_thread
    if not rodando:
        rodando = True
        start_button.config(state=tk.DISABLED)
        stop_button.config(state=tk.NORMAL)
        
        if monitor_thread is None or not monitor_thread.is_alive():
            monitor_thread = threading.Thread(target=monitorar, args=(console,), daemon=True)
            monitor_thread.start()
            escrever_log("Monitoramento iniciado.", console)
        else:
            escrever_log("Monitoramento já está rodando (thread ativa).", console)
    else:
        escrever_log("Monitoramento já está rodando.", console)

def parar(console, start_button, stop_button):
    """Interrompe o monitoramento."""
    global rodando, driver, monitor_thread
    if rodando:
        rodando = False
        escrever_log("Sinal para parar monitoramento enviado. Aguarde a finalização...", console)
        start_button.config(state=tk.NORMAL)
        stop_button.config(state=tk.DISABLED)
        
        if monitor_thread and monitor_thread.is_alive():
            monitor_thread.join(timeout=20)
            if monitor_thread.is_alive():
                escrever_log("Aviso: Thread de monitoramento não encerrou completamente dentro do tempo.", console)
        
        try:
            if driver:
                driver.quit()
                driver = None
            encerrar_chrome_debugger(app_settings["REMOTE_DEBUGGING_PORT"], console)
            escrever_log("Limpeza de processos do Chrome concluída ao parar.", console)
        except Exception as cleanup_e:
            escrever_log(f"Erro durante a limpeza ao parar: {cleanup_e}", console)
        
        monitor_thread = None

    else:
        escrever_log("Monitoramento não está rodando.", console)

# --- INTERFACE GRÁFICA ---
def criar_interface():
    """Cria a interface gráfica principal."""
    janela = tk.Tk()
    janela.title("Monitoramento TasyNative")
    janela.geometry("800x600")

    button_frame = tk.Frame(janela)
    button_frame.pack(pady=10)

    btn_iniciar = tk.Button(button_frame, text="Iniciar Monitoramento", bg="green", fg="white", font=('Arial', 12))
    btn_iniciar.pack(side=tk.LEFT, padx=5)

    btn_parar = tk.Button(button_frame, text="Parar Monitoramento", bg="red", fg="white", font=('Arial', 12), state=tk.DISABLED)
    btn_parar.pack(side=tk.LEFT, padx=5)

    btn_config = tk.Button(button_frame, text="Configurações", bg="blue", fg="white", font=('Arial', 12), command=lambda: abrir_configuracoes(janela, log_box))
    btn_config.pack(side=tk.LEFT, padx=5)

    log_box = scrolledtext.ScrolledText(janela, state='disabled', height=25, font=('Courier', 10))
    log_box.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    btn_iniciar.config(command=lambda: iniciar(log_box, btn_iniciar, btn_parar))
    btn_parar.config(command=lambda: parar(log_box, btn_iniciar, btn_parar))

    load_settings()
    escrever_log("Aplicação iniciada. Verifique as configurações.", log_box)

    janela.mainloop()

# --- EXECUÇÃO ---
if __name__ == "__main__":
    criar_interface()
