from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from openpyxl import load_workbook
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import re
import os
import datetime
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from queue import Queue

class LogManager:
    def __init__(self, log_dir: str = "log"):
        """
        Inicializa o gerenciador de logs com arquivo nomeado por data/hora
        """
        self.log_dir = log_dir
        
        # Criar diretório se não existir
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        # Gerar nome do arquivo com data e hora atual
        now = datetime.datetime.now()
        timestamp = now.strftime("%d_%m_%y__%H_%Mh")
        self.log_file = os.path.join(log_dir, f"log_{timestamp}.txt")
        
        # Criar arquivo de log
        with open(self.log_file, 'w', encoding='utf-8') as f:
            f.write("=== LOG DE LANÇAMENTOS DE AULAS ===\n")
            f.write(f"Arquivo criado em: {now.strftime('%d/%m/%Y %H:%M:%S')}\n")
            f.write("-" * 50 + "\n\n")
        
        print(f"📝 Arquivo de log criado: {self.log_file}")
    
    def registrar_lancamento(self, id_aluno: str, aulas_lancadas: list, status: str = "SUCESSO", observacoes: str = ""):
        """
        Registra um lançamento de aulas no log
        """
        timestamp = datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        
        log_entry = f"LANÇAMENTO REGISTRADO\n"
        log_entry += f"Data/Hora: {timestamp}\n"
        log_entry += f"ID do Aluno: {id_aluno}\n"
        log_entry += f"Status: {status}\n"
        log_entry += f"Aulas Processadas: {len(aulas_lancadas)}\n"
        log_entry += f"Detalhes das Aulas: {', '.join(aulas_lancadas) if aulas_lancadas else 'Nenhuma aula processada'}\n"
        
        if observacoes:
            log_entry += f"Observações: {observacoes}\n"
        
        log_entry += "-" * 30 + "\n\n"
        
        try:
            with open(self.log_file, 'a', encoding='utf-8') as f:
                f.write(log_entry)
            print(f"📝 Log registrado para aluno {id_aluno}")
        except Exception as e:
            print(f"❌ Erro ao registrar no log: {e}")
    
    def registrar_erro(self, id_aluno: str, erro: str):
        """
        Registra um erro no processamento
        """
        timestamp = datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        
        log_entry = f"ERRO REGISTRADO\n"
        log_entry += f"Data/Hora: {timestamp}\n"
        log_entry += f"ID do Aluno: {id_aluno}\n"
        log_entry += f"Erro: {erro}\n"
        log_entry += "-" * 30 + "\n\n"
        
        try:
            with open(self.log_file, 'a', encoding='utf-8') as f:
                f.write(log_entry)
            print(f"📝 Erro registrado para aluno {id_aluno}")
        except Exception as e:
            print(f"❌ Erro ao registrar erro no log: {e}")


class StopRequested(Exception):
    pass


def processar_atestados(user, password, status_cb=None, resume_event=None, stop_event=None, excel_path='atestados.xlsx'):
    """
    Executa toda a automação. 
    - status_cb: função para receber mensagens de status (str)
    - resume_event: Event controlado pela UI. Quando limpo (clear), a automação pausa. Quando setado, continua.
    - stop_event: Event para parada total e imediata assim que possível.
    - excel_path: caminho do arquivo da planilha (mantido externo ao programa).
    """
    def notify(msg: str):
        print(msg)
        if status_cb:
            try:
                status_cb(msg)
            except Exception:
                pass

    if resume_event is None:
        resume_event = threading.Event()
        resume_event.set()
    if stop_event is None:
        stop_event = threading.Event()

    def check_abort():
        if stop_event.is_set():
            raise StopRequested("Parado pelo usuário")

    def wait_if_paused():
        # Bloqueia aqui quando em pausa, mas sai se stop for solicitado
        while not resume_event.is_set():
            if stop_event.is_set():
                raise StopRequested("Parado pelo usuário")
            sleep(0.2)

    log_manager = LogManager()

    # Carregar planilha (externa ao programa)
    try:
        excel = load_workbook(excel_path)
    except FileNotFoundError:
        notify(f"Planilha não encontrada: {excel_path}. Deixe 'atestados.xlsx' na mesma pasta do programa.")
        raise
    lista_atestados = excel['Plan1']

    driver = webdriver.Chrome()
    try:
        notify("Abrindo página de login...")
        driver.get('https://senaconline-interno.sp.senac.br/psp/cs90pss/?cmd=login&languageCd=POR')

        # Helpers locais com referência ao driver atual
        def colocar_assim_aparecer(tipo, nome, conteudo, timeout=15):
            try:
                elemento = WebDriverWait(driver, timeout).until(
                    EC.presence_of_element_located((tipo, nome))
                )
                driver.execute_script("document.activeElement.blur();")
                elemento.clear()
                elemento.send_keys(str(conteudo))
                print("✅ Conteúdo inserido com sucesso.")
            except Exception:
                pass

        def clicar_assim_aparecer(tipo, nome, timeout=15):
            try:
                elemento = WebDriverWait(driver, timeout).until(
                    EC.element_to_be_clickable((tipo, nome))
                )
                elemento.click()
                print("✅ Clicou com sucesso.")
            except Exception:
                pass

        check_abort()
        notify("Realizando login...")
        colocar_assim_aparecer(By.ID, 'userid', user)
        colocar_assim_aparecer(By.ID, 'pwd', password)
        clicar_assim_aparecer(By.XPATH, '//*[@id="login"]/div/div[1]/div[8]/input')

        wait_if_paused(); check_abort()
        notify("Acessando lista de atividades por aluno...")
        clicar_assim_aparecer(By.XPATH, """//*[@id="pthnavbca_PORTAL_ROOT_OBJECT"]""")
        clicar_assim_aparecer(By.XPATH, """//*[@id="fldra_HCSR_CURRICULUM_MANAGEMENT"]""")
        clicar_assim_aparecer(By.XPATH, """//*[@id="fldra_HCSR_ATTENDANCE_ROSTER"]""")
        clicar_assim_aparecer(By.XPATH, """//*[@id="crefli_HC_STDNT_ATTENDANCE_GBL"]/a""")

        total_alunos = lista_atestados.max_row - 1
        notify(f"Iniciando processamento de {total_alunos} alunos...")

        # Loop nos IDs da planilha
        for row in lista_atestados.iter_rows(min_row=2, max_row=lista_atestados.max_row, min_col=1, max_col=lista_atestados.max_column):
            wait_if_paused(); check_abort()
            current_id = row[1].value
            current_inicio = row[2].value
            current_fim = row[3].value

            notify(f"Processando aluno {current_id}...")
            # Lista para armazenar as aulas processadas para este aluno
            aulas_processadas = []
            status_processamento = "SUCESSO"
            observacoes = ""

            try:
                driver.switch_to.default_content()
                WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ptifrmtgtframe")))

                # colocando id
                colocar_assim_aparecer(By.XPATH, """//*[@id="OR_ATND_SRCH_EMPLID"]""", current_id)
                # clicando em pesquisar
                clicar_assim_aparecer(By.XPATH, """//*[@id="#ICSearch"]""")

                driver.switch_to.default_content()
                WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ptifrmtgtframe")))

                # Tabela de resultados
                tabela = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.ID, "PTSRCHRESULTS"))
                )

                linhas = driver.find_elements(By.XPATH, '//*[@id="PTSRCHRESULTS"]/tbody/tr')

                encontrou = False
                for linha in linhas:
                    texto = linha.text.upper()
                    if "2025" in texto and "EMÉDIO" in texto:
                        link = linha.find_element(By.XPATH, ".//a")
                        link.click()
                        encontrou = True
                        break

                if not encontrou:
                    status_processamento = "ERRO"
                    observacoes = "Aluno não encontrado ou não possui curso EMÉDIO 2025"
                    log_manager.registrar_lancamento(current_id, aulas_processadas, status_processamento, observacoes)
                    notify(f"Aluno {current_id}: não encontrado/sem EMÉDIO 2025")
                    continue

                driver.switch_to.default_content()
                WebDriverWait(driver, 20).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ptifrmtgtframe")))

                # Pega a tabela de aulas
                tabela = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="STDNT_ENRL$scroll$0"]'))
                )

                linhas = tabela.find_elements(By.TAG_NAME, "tr")

                primeira_coluna_elementos = []

                for linha in linhas:
                    colunas = linha.find_elements(By.TAG_NAME, "td")
                    if colunas:
                        primeira_coluna_elementos.append(colunas[0])

                elementos_aulas = []

                for elemento in primeira_coluna_elementos:
                    texto = elemento.text.strip()
                    if re.fullmatch(r'\d+', texto):  # Só números
                        elementos_aulas.append(elemento)

                notify(f"Encontradas {len(elementos_aulas)} aulas para o aluno {current_id}.")

                # Agora vamos clicar um a um nos elementos de aula
                for i in range(len(elementos_aulas)):
                    wait_if_paused(); check_abort()
                    # Retorna para o contexto padrão
                    driver.switch_to.default_content()
                    WebDriverWait(driver, 20).until(
                        EC.frame_to_be_available_and_switch_to_it((By.ID, "ptifrmtgtframe"))
                    )

                    # Localiza novamente a tabela e os elementos
                    tabela = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.XPATH, '//*[@id="STDNT_ENRL$scroll$0"]'))
                    )
                    linhas = tabela.find_elements(By.TAG_NAME, "tr")

                    primeira_coluna_elementos = []
                    for linha in linhas:
                        colunas = linha.find_elements(By.TAG_NAME, "td")
                        if colunas:
                            primeira_coluna_elementos.append(colunas[0])

                    # Filtra os elementos que você quer clicar (apenas números de aula)
                    elementos_aulas = [elem for elem in primeira_coluna_elementos if len(elem.text.strip()) <= 4]

                    # Agora pega o elemento atual baseado no índice do loop
                    elemento = elementos_aulas[i]
                    numero_aula = elemento.text.strip()

                    try:
                        link = elemento.find_element(By.TAG_NAME, "a")
                        driver.execute_script("arguments[0].scrollIntoView();", link)
                        link.click()

                        print(f"✅ Clicou no link do número da aula {i+1}: {numero_aula}")
                        sleep(2)  # Aguarda um pouco após o clique

                    except Exception as e:
                        print(f"❌ Erro ao clicar no link do elemento {i+1}: {e}")
                        continue

                    # Aqui dentro da página aberta, verifica se o texto "Matric" está presente
                    try:
                        tabela_verificacao = WebDriverWait(driver, 20).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="ACE_DERIVED_AA2_"]'))
                        )
                        # printar a informação que apararece em tabela_verificacao
                        print(tabela_verificacao.text)

                        if "Matric" in tabela_verificacao.text:
                            print("✅ O texto 'Matric' foi encontrado na tabela.")
                            clicar_assim_aparecer(By.XPATH, '//*[@id="ICTAB_1"]')

                            # Inserir as datas e informações
                            colocar_assim_aparecer(By.XPATH, '//*[@id="DIG_APR_EST_WRK_START_DT"]', current_inicio.strftime('%d/%m/%Y'))
                            colocar_assim_aparecer(By.XPATH, '//*[@id="DIG_APR_EST_WRK_END_DT"]', current_fim.strftime('%d/%m/%Y'))

                            # Selecionar "Amparo Legal" no select
                            select_element = WebDriverWait(driver, 15).until(
                                EC.presence_of_element_located((By.XPATH, '//*[@id="DIG_APR_EST_WRK_ATTEND_REASON"]'))
                            )
                            for option in select_element.find_elements(By.TAG_NAME, 'option'):
                                if option.text.strip() == "Amparo Legal":
                                    option.click()
                                    print("✅ Opção 'Amparo Legal' selecionada.")
                                    break

                            # Inserir o motivo
                            colocar_assim_aparecer(By.XPATH, '//*[@id="DIG_APR_EST_WRK_REASON_DESCR"]', "0000000001")

                            print("✅ Dados inseridos com sucesso.")

                            # clicar no botao de aplicar
                            clicar_assim_aparecer(By.XPATH, '//*[@id="DIG_APR_EST_WRK_PROCESS_BTN"]')
                            # clicar em botao de salvar
                            clicar_assim_aparecer(By.XPATH, '//*[@id="#ICSave"]')
                            # clicar no link de voltar para a lista
                            clicar_assim_aparecer(By.XPATH, '//*[@id="DERIVED_AA2_DERIVED_LINK10$0"]')

                            # Adicionar aula processada com sucesso à lista
                            aulas_processadas.append(f"Aula {numero_aula} - Lançamento realizado")
                            sleep(2)
                            
                        else:
                            print("❌ O texto 'Matric' NÃO foi encontrado na tabela.")
                            clicar_assim_aparecer(By.XPATH, '//*[@id="DERIVED_AA2_DERIVED_LINK10$0"]')
                            # Adicionar aula que não foi processada à lista
                            aulas_processadas.append(f"Aula {numero_aula} - Não processada (sem Matric)")
                            sleep(2)

                    except Exception as e:
                        print(f"❌ Erro ao buscar a tabela: {e}")
                        aulas_processadas.append(f"Aula {numero_aula} - Erro: {str(e)}")

                    # Se necessário, clique para retornar após cada verificação, para não quebrar o próximo loop
                    try:
                        voltar = driver.find_element(By.XPATH, '//*[@id="DERIVED_AA2_DERIVED_LINK10$0"]')
                        voltar.click()
                        sleep(2)
                    except Exception as e:
                        print(f"❌ Erro ao retornar: {e}")

            except StopRequested:
                notify("Parada solicitada. Encerrando processamento atual...")
                break
            except Exception as e:
                status_processamento = "ERRO"
                observacoes = f"Erro geral no processamento: {str(e)}"
                log_manager.registrar_erro(current_id, str(e))

            # Registrar o lançamento no log para este aluno
            if aulas_processadas:
                observacoes_final = f"Período: {current_inicio.strftime('%d/%m/%Y')} a {current_fim.strftime('%d/%m/%Y')}"
                if observacoes:
                    observacoes_final += f" | {observacoes}"
                log_manager.registrar_lancamento(current_id, aulas_processadas, status_processamento, observacoes_final)
            else:
                log_manager.registrar_lancamento(current_id, [], "SEM_AULAS", "Nenhuma aula foi encontrada para processar")

            check_abort()
            driver.switch_to.default_content()
            # Acessando lista de atividades por aluno
            clicar_assim_aparecer(By.XPATH, """//*[@id=\"pthnavbca_PORTAL_ROOT_OBJECT\"]""")
            clicar_assim_aparecer(By.XPATH, """//*[@id=\"fldra_HCSR_CURRICULUM_MANAGEMENT\"]""")
            clicar_assim_aparecer(By.XPATH, """//*[@id=\"fldra_HCSR_ATTENDANCE_ROSTER\"]""")
            clicar_assim_aparecer(By.XPATH, """//*[@id=\"crefli_HC_STDNT_ATTENDANCE_GBL\"]/a""")

            notify(f"Aluno {current_id} finalizado.")

        notify("🎉 Processamento concluído! Verifique o arquivo de log na pasta 'log'.")

    except StopRequested:
        notify("Execução interrompida pelo usuário.")
    finally:
        try:
            driver.quit()
        except Exception:
            pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AutoAtestado - SENAC")
        self.geometry("460x300")
        self.resizable(False, False)

        # Estado
        self.worker_thread = None
        self.resume_event = threading.Event(); self.resume_event.set()
        self.stop_event = threading.Event()
        self.is_running = False
        self.is_paused = False
        self.pending_restart = False
        self.status_queue = Queue()

        # UI
        padding = {"padx": 8, "pady": 6}
        frm = ttk.Frame(self)
        frm.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        ttk.Label(frm, text="Usuário:").grid(row=0, column=0, sticky="w", **padding)
        self.user_var = tk.StringVar()
        self.user_entry = ttk.Entry(frm, textvariable=self.user_var, width=32)
        self.user_entry.grid(row=0, column=1, columnspan=3, sticky="w", **padding)

        ttk.Label(frm, text="Senha:").grid(row=1, column=0, sticky="w", **padding)
        self.pass_var = tk.StringVar()
        self.pass_entry = ttk.Entry(frm, textvariable=self.pass_var, width=32, show="*")
        self.pass_entry.grid(row=1, column=1, columnspan=3, sticky="w", **padding)

        self.start_btn = ttk.Button(frm, text="Iniciar", command=self.on_start)
        self.start_btn.grid(row=2, column=0, **padding)

        self.pause_btn = ttk.Button(frm, text="Pausar", command=self.on_pause_resume, state=tk.DISABLED)
        self.pause_btn.grid(row=2, column=1, **padding)

        self.stop_btn = ttk.Button(frm, text="Parar", command=self.on_stop, state=tk.DISABLED)
        self.stop_btn.grid(row=2, column=2, **padding)

        self.restart_btn = ttk.Button(frm, text="Reiniciar", command=self.on_restart)
        self.restart_btn.grid(row=2, column=3, **padding)

        ttk.Separator(frm).grid(row=3, column=0, columnspan=4, sticky="ew", pady=(10, 0))

        ttk.Label(frm, text="Status:").grid(row=4, column=0, sticky="w", **padding)
        self.status_var = tk.StringVar(value="Pronto.")
        self.status_lbl = ttk.Label(frm, textvariable=self.status_var, wraplength=400, anchor="w", justify="left")
        self.status_lbl.grid(row=4, column=1, columnspan=3, sticky="w", **padding)

        # Info da planilha (externa)
        ttk.Label(frm, text="Planilha: 'atestados.xlsx' (mesma pasta do programa)").grid(row=5, column=0, columnspan=4, sticky="w", padx=8)

        self.bind('<Return>', lambda e: self.on_start() if not self.is_running else None)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # Poll da fila de status
        self.after(200, self._poll_status)

    def _enqueue_status(self, msg: str):
        self.status_queue.put(msg)

    def _poll_status(self):
        try:
            while True:
                msg = self.status_queue.get_nowait()
                self.status_var.set(msg)
        except Exception:
            pass
        # Se houver reinício pendente e não estiver rodando, reinicia
        if self.pending_restart and not self.is_running:
            self.pending_restart = False
            self.on_start()
        self.after(200, self._poll_status)

    def on_start(self):
        if self.is_running:
            return
        user = self.user_var.get().strip()
        pwd = self.pass_var.get().strip()
        if not user or not pwd:
            messagebox.showwarning("Dados obrigatórios", "Informe usuário e senha.")
            return
        self.is_running = True
        self.is_paused = False
        self.resume_event.set()
        self.stop_event.clear()

        # Travar inputs enquanto roda
        self.user_entry.configure(state=tk.DISABLED)
        self.pass_entry.configure(state=tk.DISABLED)
        self.start_btn.configure(state=tk.DISABLED)
        self.pause_btn.configure(state=tk.NORMAL, text="Pausar")
        self.stop_btn.configure(state=tk.NORMAL)

        self.status_var.set("Iniciando automação...")

        def target():
            try:
                processar_atestados(user, pwd, status_cb=self._enqueue_status, resume_event=self.resume_event, stop_event=self.stop_event)
                # Notificar término (se não foi parado manualmente)
                if not self.stop_event.is_set():
                    self._enqueue_status("Concluído. Veja o log na pasta 'log'.")
                    self.after(0, lambda: messagebox.showinfo("AutoAtestado", "Processamento concluído!"))
            except Exception as e:
                self._enqueue_status(f"Erro: {e}")
                self.after(0, lambda: messagebox.showerror("Erro", f"Falha na execução: {e}"))
            finally:
                # Restaurar UI
                def restore():
                    self.is_running = False
                    self.is_paused = False
                    self.user_entry.configure(state=tk.NORMAL)
                    self.pass_entry.configure(state=tk.NORMAL)
                    self.start_btn.configure(state=tk.NORMAL)
                    self.pause_btn.configure(state=tk.DISABLED, text="Pausar")
                    self.stop_btn.configure(state=tk.DISABLED)
                self.after(0, restore)

        self.worker_thread = threading.Thread(target=target, daemon=True)
        self.worker_thread.start()

    def on_pause_resume(self):
        if not self.is_running:
            return
        if not self.is_paused:
            # Pausar
            self.resume_event.clear()
            self.is_paused = True
            self.pause_btn.configure(text="Retomar")
            self.status_var.set("Pausado. Clique em Retomar para continuar.")
        else:
            # Retomar
            self.resume_event.set()
            self.is_paused = False
            self.pause_btn.configure(text="Pausar")
            self.status_var.set("Retomando...")

    def on_stop(self):
        if not self.is_running:
            return
        if messagebox.askyesno("Parar", "Deseja interromper o processamento atual?"):
            self.stop_event.set()
            self.resume_event.set()  # garante sair da pausa
            self.status_var.set("Parando... aguarde a etapa atual finalizar.")
            # Botões: desabilita parar para evitar múltiplos cliques
            self.stop_btn.configure(state=tk.DISABLED)

    def on_restart(self):
        if self.is_running:
            if not messagebox.askyesno("Reiniciar", "Isso vai interromper a execução atual e iniciar do começo da planilha. Continuar?"):
                return
            self.pending_restart = True
            self.on_stop()
        else:
            self.on_start()

    def on_close(self):
        if self.is_running:
            if not messagebox.askyesno("Sair", "A automação ainda está em execução. Deseja encerrar mesmo assim?"):
                return
            # Solicita parada
            self.stop_event.set()
            self.resume_event.set()
        try:
            pass
        finally:
            self.destroy()


if __name__ == '__main__':
    app = App()
    app.mainloop()
