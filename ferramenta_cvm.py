import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import os
import re
import requests
from datetime import datetime
import json
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading


PASTA_PRINCIPAL_PDFS = "Decisoes_PDFs"
ARQUIVO_SAIDA_LOG = "log_de_extracao.xlsx"
ARQUIVO_LOG_SUCESSO_TXT = "log_sucesso.txt"
ARQUIVO_LOG_ERROS_TXT = "log_erros.txt"
MAX_TENTATIVAS_PAGINACAO = 5
REINICIAR_NAVEGADOR_A_CADA_PAGINAS = 50

class CvmApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ferramenta de Busca e Extração CVM")
        self.root.geometry("850x700") 


        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", font=('Segoe UI', 9))
        style.configure("TLabel", padding=5, font=('Segoe UI', 9))
        style.configure("TEntry", padding=5)
        style.configure("TLabelframe.Label", font=('Segoe UI', 10, 'bold'))

        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        controls_frame = ttk.LabelFrame(main_frame, text="Opções de Busca Específica", padding="10")
        controls_frame.pack(fill=tk.X, expand=False, pady=5)
        controls_frame.columnconfigure(1, weight=1)

        ttk.Label(controls_frame, text="Página Específica:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.page_entry = ttk.Entry(controls_frame, width=10)
        self.page_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        self.page_button = ttk.Button(controls_frame, text="Buscar por Página", command=self.iniciar_busca_por_pagina)
        self.page_button.grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(controls_frame, text="Termo/Processo:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.term_entry = ttk.Entry(controls_frame)
        self.term_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        self.term_button = ttk.Button(controls_frame, text="Buscar por Termo", command=self.iniciar_busca_por_termo)
        self.term_button.grid(row=1, column=2, padx=5, pady=5)

        ttk.Label(controls_frame, text="Período de:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        date_frame = ttk.Frame(controls_frame)
        date_frame.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        self.start_date_entry = ttk.Entry(date_frame, width=15)
        self.start_date_entry.pack(side=tk.LEFT)
        self.start_date_entry.insert(0, "dd/mm/aaaa")
        ttk.Label(date_frame, text=" até ").pack(side=tk.LEFT, padx=5)
        self.end_date_entry = ttk.Entry(date_frame, width=15)
        self.end_date_entry.pack(side=tk.LEFT)
        self.end_date_entry.insert(0, "dd/mm/aaaa")
        self.date_button = ttk.Button(controls_frame, text="Buscar por Período", command=self.iniciar_busca_por_data)
        self.date_button.grid(row=2, column=2, padx=5, pady=5)


        full_scan_frame = ttk.LabelFrame(main_frame, text="Busca Completa", padding="10")
        full_scan_frame.pack(fill=tk.X, expand=False, pady=10)


        self.full_scan_button = ttk.Button(full_scan_frame, text="Verificar e Extrair Todas as Decisões (Pode levar muito tempo)", command=self.iniciar_busca_completa)
        self.full_scan_button.pack(fill=tk.X, expand=True)


        log_frame = ttk.LabelFrame(main_frame, text="Log de Atividades", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=90, height=25, font=('Courier New', 9))
        self.log_area.pack(fill=tk.BOTH, expand=True)
        self.log_area.configure(state='disabled')
        
        self.buttons = [self.page_button, self.term_button, self.date_button, self.full_scan_button]

    def log(self, message):
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, str(message) + "\n")
        self.log_area.configure(state='disabled')
        self.log_area.see(tk.END)
        self.root.update_idletasks()

    def toggle_buttons(self, enabled):
        state = "normal" if enabled else "disabled"
        for button in self.buttons:
            button.config(state=state)

    def iniciar_busca(self, target_function, *args):
        self.toggle_buttons(False)
        self.log_area.configure(state='normal')
        self.log_area.delete('1.0', tk.END)
        self.log_area.configure(state='disabled')
        
        thread = threading.Thread(target=target_function, args=args)
        thread.daemon = True
        thread.start()

    def iniciar_busca_completa(self):
        self.iniciar_busca(self.executar_busca, self._buscar_tudo_logic)
        
    def iniciar_busca_por_pagina(self):
        page_num = self.page_entry.get()
        self.iniciar_busca(self.executar_busca, self._buscar_por_pagina_logic, page_num)

    def iniciar_busca_por_termo(self):
        term = self.term_entry.get()
        self.iniciar_busca(self.executar_busca, self._buscar_por_termo_logic, term)

    def iniciar_busca_por_data(self):
        start_date = self.start_date_entry.get()
        end_date = self.end_date_entry.get()
        self.iniciar_busca(self.executar_busca, self._buscar_por_data_logic, start_date, end_date)
        
    def executar_busca(self, funcao_busca, *args):
        try:
            df_resultados = funcao_busca(*args)
            if df_resultados is not None and not df_resultados.empty:
                num_links = len(df_resultados)
                self.log(f"\nBusca concluída. {num_links} links foram encontrados.")
                if messagebox.askyesno("Confirmar Processamento", f"{num_links} links encontrados. Deseja iniciar o download dos PDFs?"):
                    self.processar_links(df_resultados)
                else:
                    self.log("Processamento de PDFs cancelado pelo usuário.")
            else:
                self.log("\nNenhum resultado encontrado para os critérios de busca informados.")
        except Exception as e:
            self.log(f"\n--- ERRO INESPERADO NA EXECUÇÃO ---\n{e}")
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro inesperado:\n{e}")
        finally:
            self.log("\n--- Tarefa Concluída ---")
            self.toggle_buttons(True)

    def limpar_nome(self, nome, tamanho_max=150):
        nome = re.sub(r'[\\/*?:"<>|]', "-", nome)
        nome = ' '.join(nome.split())
        return nome.strip()[:tamanho_max]

    def iniciar_driver_para_processamento(self):
        self.log("Iniciando sessão do navegador...")
        options = Options()
        options.add_argument("--start-maximized")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        settings = {"recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}], "selectedDestinationId": "Save as PDF", "version": 2}
        caminho_absoluto_pdfs = os.path.join(os.getcwd(), PASTA_PRINCIPAL_PDFS)
        os.makedirs(caminho_absoluto_pdfs, exist_ok=True)
        prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings), 'savefile.default_directory': caminho_absoluto_pdfs}
        options.add_experimental_option('prefs', prefs)
        options.add_argument('--kiosk-printing')
        try:
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            driver.set_script_timeout(120)
            return driver
        except Exception as e:
            self.log(f"[!] ERRO ao iniciar o WebDriver: {e}")
            messagebox.showerror("Erro de Driver", f"Não foi possível iniciar o navegador. Verifique sua conexão ou se o Chrome está atualizado.\n\nErro: {e}")
            return None

    def raspar_links_da_pagina_atual(self, driver):
        links_encontrados = []
        base_url = "https://conteudo.cvm.gov.br"
        soup = BeautifulSoup(driver.page_source, 'lxml')
        artigos = soup.select("section.listaResultados article")
        for artigo in artigos:
            link_tag = artigo.select_one("h3 a")
            if link_tag:
                titulo = link_tag.get_text(strip=True)
                url_completa = urljoin(base_url, link_tag.get('href'))
                links_encontrados.append({'Título': titulo, 'URL': url_completa})
        return links_encontrados

    def lidar_com_paginacao_e_raspar_tudo(self, driver, wait):
        todos_os_links = []
        pagina_num = 1
        while True:
            if pagina_num > 1 and (pagina_num - 1) % REINICIAR_NAVEGADOR_A_CADA_PAGINAS == 0:
                self.log(f"\n--- ATINGIDO O LIMITE DE {REINICIAR_NAVEGADOR_A_CADA_PAGINAS} PÁGINAS. REINICIANDO O NAVEGADOR PARA LIBERAR MEMÓRIA... ---")
                
                pagina_a_retornar = pagina_num
                url_da_busca = driver.current_url
                
                driver.quit()
                time.sleep(5)
                driver = self.iniciar_driver_para_processamento()
                wait = WebDriverWait(driver, 30)
                
                self.log(f"Retornando para a busca e navegando para a página {pagina_a_retornar}...")
                driver.get(url_da_busca)
                
                wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "section.listaResultados article")))
                
                if pagina_a_retornar > 1:
                    try:
                        primeiro_titulo_antes = driver.find_element(By.CSS_SELECTOR, "section.listaResultados article:first-of-type a").text
                        campo_pagina = driver.find_element(By.ID, 'irPara')
                        campo_pagina.clear()
                        campo_pagina.send_keys(str(pagina_a_retornar))
                        

                        botao_ir = driver.find_element(By.ID, 'irParaButton')
                        driver.execute_script("arguments[0].click();", botao_ir)

                        def o_titulo_da_pagina_mudou(d):
                            try:
                                novo_titulo = d.find_element(By.CSS_SELECTOR, "section.listaResultados article:first-of-type a").text
                                return novo_titulo != primeiro_titulo_antes
                            except (NoSuchElementException, StaleElementReferenceException):
                                return False
                        wait.until(o_titulo_da_pagina_mudou)
                        self.log(f"✔️ Retorno para a página {pagina_a_retornar} concluído.")
                    except Exception as e:
                        self.log(f"[!] Falha ao tentar retornar para a página {pagina_a_retornar} após reiniciar. Erro: {e}")
                        break

            self.log(f"Coletando links da página de resultados nº {pagina_num}...")
            links_da_pagina = self.raspar_links_da_pagina_atual(driver)
            
            if not links_da_pagina and not todos_os_links:
                self.log("Nenhum link encontrado nesta página de resultados.")
                break
                
            todos_os_links.extend(links_da_pagina)
            self.log(f"Encontrados {len(links_da_pagina)} links. Total até agora: {len(todos_os_links)}")
            
            sucesso_na_paginacao = False
            for tentativa in range(MAX_TENTATIVAS_PAGINACAO):
                try:
                    primeiro_titulo_pagina_atual = links_da_pagina[0]['Título'] if links_da_pagina else ""
                    
                    xpath_botao_proxima = '//li/button[normalize-space()="Próxima"]'
                    wait.until(EC.element_to_be_clickable((By.XPATH, xpath_botao_proxima)))

                    li_botao_proxima_container = driver.find_element(By.XPATH, '//li[button[normalize-space()="Próxima"]]')
                    if "disabled" in li_botao_proxima_container.get_attribute("class"):
                        self.log("Fim dos resultados da busca (botão 'Próxima' desabilitado).")
                        return todos_os_links, driver

                    botao = driver.find_element(By.XPATH, xpath_botao_proxima)
                    botao.click()
                    
                    def o_titulo_mudou(d):
                        try:
                            novo_titulo = d.find_element(By.CSS_SELECTOR, "section.listaResultados article:first-of-type a").text
                            return novo_titulo != primeiro_titulo_pagina_atual
                        except (NoSuchElementException, StaleElementReferenceException):
                            return False
                    wait.until(o_titulo_mudou)
                    
                    sucesso_na_paginacao = True
                    break

                except Exception as e:
                    self.log(f"  [!] Ocorreu um erro na paginação (tentativa {tentativa + 1}/{MAX_TENTATIVAS_PAGINACAO}).")
                    if tentativa < MAX_TENTATIVAS_PAGINACAO - 1:
                        self.log("    Aguardando 20 segundos para o site estabilizar e tentando novamente...")
                        time.sleep(20)
                    else:
                        self.log(f"    [!] Falha final na paginação após {MAX_TENTATIVAS_PAGINACAO} tentativas.")
            
            if not sucesso_na_paginacao:
                self.log("Não foi possível continuar a paginação.")
                break
                
            pagina_num += 1
                
        return todos_os_links, driver

    def _buscar_tudo_logic(self):
        driver = self.iniciar_driver_para_processamento()
        if not driver: return pd.DataFrame()
        wait = WebDriverWait(driver, 30)
        url_base_busca = "https://conteudo.cvm.gov.br/decisoes/index.html?lastNameShow=&lastName=&filtro=todos&dataInicio=&dataFim=&buscadoDecisao=false&categoria=decisao"
        driver.get(url_base_busca)
        links = []
        try:
            self.log("Iniciando busca completa em todas as páginas...")
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "section.listaResultados article")))
            links, driver = self.lidar_com_paginacao_e_raspar_tudo(driver, wait)
        except Exception as e:
            self.log(f"[!] Ocorreu um erro durante a busca completa: {e}")
        finally:
            if 'driver' in locals() and driver.service.is_connectable(): driver.quit()
        return pd.DataFrame(links)

    def _buscar_por_pagina_logic(self, num_pagina):
        if not num_pagina.isdigit() or int(num_pagina) < 1:
            self.log("[!] Entrada inválida. Por favor, digite um número de página válido.")
            return pd.DataFrame()
        driver = self.iniciar_driver_para_processamento()
        if not driver: return pd.DataFrame()
        wait = WebDriverWait(driver, 30)
        url_base_busca = "https://conteudo.cvm.gov.br/decisoes/index.html?lastNameShow=&lastName=&filtro=todos&dataInicio=&dataFim=&buscadoDecisao=false&categoria=decisao"
        driver.get(url_base_busca)
        links = []
        try:
            self.log(f"Navegando para a página {num_pagina}...")
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "section.listaResultados")))
            primeiro_titulo_antes = driver.find_element(By.CSS_SELECTOR, "section.listaResultados article:first-of-type a").text
            campo_pagina = driver.find_element(By.ID, 'irPara')
            campo_pagina.clear()
            campo_pagina.send_keys(str(num_pagina))
            driver.find_element(By.ID, 'irParaButton').click()
            def o_titulo_da_pagina_mudou(d):
                try:
                    novo_titulo = d.find_element(By.CSS_SELECTOR, "section.listaResultados article:first-of-type a").text
                    return novo_titulo != primeiro_titulo_antes
                except NoSuchElementException: return False
            if int(num_pagina) > 1: wait.until(o_titulo_da_pagina_mudou)
            self.log("Página carregada. Coletando links...")
            
            links = self.raspar_links_da_pagina_atual(driver)

        except Exception as e:
            self.log(f"[!] Ocorreu um erro durante a busca por página: {e}")
        finally:
            if 'driver' in locals() and driver.service.is_connectable(): driver.quit()
        return pd.DataFrame(links)
    

    def _buscar_por_termo_logic(self, termo):
        driver = self.iniciar_driver_para_processamento()
        if not driver: return pd.DataFrame()
        wait = WebDriverWait(driver, 30)
        url_base_busca = "https://conteudo.cvm.gov.br/decisoes/index.html?lastNameShow=&lastName=&filtro=todos&dataInicio=&dataFim=&buscadoDecisao=false&categoria=decisao"
        driver.get(url_base_busca)
        links = []
        try:
            self.log(f"Buscando por '{termo}' com a opção 'Expressão exata'...")
            wait.until(EC.visibility_of_element_located((By.ID, 'termoShow')))
            driver.find_element(By.ID, 'termoShow').send_keys(termo)
            driver.find_element(By.ID, 'expressao').click()
            driver.find_element(By.CSS_SELECTOR, 'button.submit').click()
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "section.listaResultados article")))
            links, driver = self.lidar_com_paginacao_e_raspar_tudo(driver, wait)
        except TimeoutException: self.log("A busca não retornou resultados ou demorou demais para carregar.")
        except Exception as e: self.log(f"[!] Ocorreu um erro durante a busca por termo: {e}")
        finally:
            if 'driver' in locals() and driver.service.is_connectable(): driver.quit()
        return pd.DataFrame(links)

    def _buscar_por_data_logic(self, data_inicio, data_fim):
        driver = self.iniciar_driver_para_processamento()
        if not driver: return pd.DataFrame()
        wait = WebDriverWait(driver, 30)
        url_base_busca = "https://conteudo.cvm.gov.br/decisoes/index.html?lastNameShow=&lastName=&filtro=todos&dataInicio=&dataFim=&buscadoDecisao=false&categoria=decisao"
        driver.get(url_base_busca)
        links = []
        try:
            self.log(f"Buscando no período de {data_inicio} a {data_fim}...")
            wait.until(EC.visibility_of_element_located((By.ID, 'dataInicio')))
            driver.find_element(By.ID, 'dataInicio').send_keys(data_inicio)
            driver.find_element(By.ID, 'dataFim').send_keys(data_fim)
            driver.find_element(By.CSS_SELECTOR, 'button.submit').click()
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "section.listaResultados article")))
            links, driver = self.lidar_com_paginacao_e_raspar_tudo(driver, wait)
        except TimeoutException: self.log("A busca não retornou resultados ou demorou demais para carregar.")
        except Exception as e: self.log(f"[!] Ocorreu um erro durante a busca por data: {e}")
        finally:
            if 'driver' in locals() and driver.service.is_connectable(): driver.quit()
        return pd.DataFrame(links)
        
    def processar_links(self, df_para_processar):
        if df_para_processar.empty:
            self.log("Nenhum link para processar.")
            return

        self.log(f"\n--- INICIANDO PROCESSAMENTO DE LINKS ---")
        
        log_extracao = []
        links_ja_processados = set()
        if os.path.exists(ARQUIVO_SAIDA_LOG):
            df_log_existente = pd.read_excel(ARQUIVO_SAIDA_LOG)
            if 'URL Original' in df_log_existente.columns:
                links_ja_processados.update(df_log_existente['URL Original'].dropna().astype(str).str.strip())
            log_extracao = df_log_existente.to_dict('records')
            self.log(f"✔️ {len(links_ja_processados)} links já foram processados anteriormente e serão pulados.")

        df_links_a_fazer = df_para_processar[~df_para_processar['URL'].isin(links_ja_processados)].reset_index(drop=True)

        if df_links_a_fazer.empty:
            self.log("Todos os links encontrados na busca já foram processados anteriormente.")
            return
        
        self.log(f"Novos links a processar: {len(df_links_a_fazer)}")
            
        driver = self.iniciar_driver_para_processamento()
        if not driver: return
        
        wait = WebDriverWait(driver, 30)
        base_url = "https://conteudo.cvm.gov.br"
        caminho_absoluto_pdfs = os.path.join(os.getcwd(), PASTA_PRINCIPAL_PDFS)

        for index, row in df_links_a_fazer.iterrows():
            url = row['URL']
            titulo_original_excel = row['Título']
            
            self.log(f"\n--- Processando link #{index + 1} de {len(df_links_a_fazer)}: {url} ---")
            
            linhas_relatorio = []
            houve_erro_no_link = False

            try:
                driver.get(url)
                wait.until(EC.visibility_of_element_located((By.ID, "main")))
                soup = BeautifulSoup(driver.page_source, 'lxml')
                main_div = soup.select_one('div#main')
                if not main_div: raise Exception("Div 'main' não encontrada.")
                    
                h2_tag = main_div.select_one("h2")
                texto_h2 = h2_tag.get_text(strip=True) if h2_tag else ""
                match_data = re.search(r'(\d{2}/\d{2}/\d{4})', texto_h2)
                data_formatada = datetime.strptime(match_data.group(1), '%d/%m/%Y').strftime('%Y.%m.%d') if match_data else "DataNaoEncontrada"

                titulo_da_decisao = titulo_original_excel
                num_proc_str = ""
                
                true_title_tag = main_div.select_one("p.text-uppercase b")
                if true_title_tag:
                    true_title_text = true_title_tag.get_text(strip=True, separator=" ")
                    regex_processos = r'\bPROC\.\s*[A-Z]{2}\s*\d{1,4}/\d+\b|\b(?:PAS|PROC|PROCESSOS)\s+(?:Nº\s*)?\d+/\d{4}\b|\b(?:PROC|PAS|PROCESSOS)(?:\s*SEI)?\s+[A-Z]{2}\s*\d{1,4}/\d+\b|\b[A-Z]{2}\s*\d{1,4}/\d+\b|\b\d{5}\.\d{6}/\d{4}-\d{2}\b|\bPROC\.\s*\d{2}/\d+\b|\bPROC\.\s*[A-Z]{2}´\d{4}/\d+\b'
                    processos_encontrados = re.findall(regex_processos, true_title_text, re.IGNORECASE)
                    processos_unicos = sorted(list(set(processos_encontrados)))
                    
                    if len(processos_unicos) > 2:
                        processos_unicos = processos_unicos[:2]
                    
                    if processos_unicos:
                        num_proc_str = ' & '.join(processos_unicos)

                    titulo_limpo = true_title_text
                    for proc in processos_unicos:
                        titulo_limpo = titulo_limpo.replace(proc, '')
                    titulo_da_decisao = re.sub(r'\s*[-–]\s*(?:PROC|PAS|PROCESSOS|SEI).*$', '', titulo_limpo, flags=re.IGNORECASE).strip()
                    if not titulo_da_decisao: titulo_da_decisao = titulo_original_excel
                
                if not num_proc_str:
                    raise Exception("Número do processo não foi encontrado na página.")

                numero_limpo = re.sub(r'^(PROC|PAS|PROCESSOS)\s*\.?\s*', '', num_proc_str, flags=re.IGNORECASE)
                nome_base_arquivo = f"{data_formatada} - Nº PROC. {numero_limpo} - {titulo_da_decisao}"
                
                nome_arquivo_seguro = self.limpar_nome(nome_base_arquivo) + ".pdf"
                caminho_final_arquivo = os.path.join(caminho_absoluto_pdfs, nome_arquivo_seguro)
                
                contador_duplicado = 2
                while os.path.exists(caminho_final_arquivo):
                    caminho_final_arquivo = os.path.join(caminho_absoluto_pdfs, f"{self.limpar_nome(nome_base_arquivo)} ({contador_duplicado}).pdf")
                    contador_duplicado += 1
                
                nome_arquivo_final = os.path.basename(caminho_final_arquivo)
                driver.execute_script("document.title = 'arquivo_temp_cvm';")
                arquivos_antes = set(os.listdir(caminho_absoluto_pdfs))
                driver.execute_script('window.print();')
                
                arquivo_baixado = None
                for _ in range(90):
                    time.sleep(1)
                    novos_arquivos = set(os.listdir(caminho_absoluto_pdfs)) - arquivos_antes
                    if novos_arquivos:
                        nome_novo_arquivo = novos_arquivos.pop()
                        if not nome_novo_arquivo.endswith('.crdownload'):
                            arquivo_baixado = os.path.join(caminho_absoluto_pdfs, nome_novo_arquivo)
                            break
                
                if not arquivo_baixado: raise Exception(f"Falha crítica ao gerar PDF para o link {url}.")

                os.rename(arquivo_baixado, caminho_final_arquivo)
                self.log(f"  ✔️ PDF Principal salvo como: '{nome_arquivo_final}'")
                log_extracao.append({'Data da Extração': datetime.now().strftime('%d/%m/%Y'), 'Hora da Extração': datetime.now().strftime('%H:%M:%S'), 'Título Completo': titulo_da_decisao, 'Números de Processo': num_proc_str, 'Data da Decisão': data_formatada, 'Nome do Arquivo': nome_arquivo_final, 'Local Salvo': caminho_final_arquivo, 'URL Original': url})

                tabela_anexos = soup.select("div.boxVejaMais table a")
                
                self.log("  - Verificando anexos...")
                if not tabela_anexos:
                    self.log("    - Nenhum anexo encontrado.")
                else:
                    for link_anexo_tag in tabela_anexos:
                        nome_anexo = link_anexo_tag.get_text(strip=True)
                        if nome_anexo and link_anexo_tag.get('href'):
                            pass 
            except Exception as e:
                houve_erro_no_link = True
                self.log(f"  [!] ERRO GERAL ao processar o link {url}. Erro: {e}")
                log_extracao.append({'Data da Extração': datetime.now().strftime('%d/%m/%Y'), 'Hora da Extração': datetime.now().strftime('%H:%M:%S'), 'Título Completo': f"ERRO - {titulo_original_excel}", 'Números de Processo': 'ERRO', 'Data da Decisão': 'ERRO', 'Nome do Arquivo': 'ERRO', 'Local Salvo': f"ERRO ao processar URL: {url}", 'URL Original': url})
            
        if 'driver' in locals() and driver.service.is_connectable():
            driver.quit()

        self.log("\n--- Coleta finalizada. Salvando log em Excel... ---")
        if log_extracao:
            pd.DataFrame(log_extracao).to_excel(ARQUIVO_SAIDA_LOG, index=False)
        
        self.log(f"\n🔎 Processo concluído.")


if __name__ == "__main__":
    root = tk.Tk()
    app = CvmApp(root)
    root.mainloop()