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

# Inicializar o sistema de logging
log_manager = LogManager()

# remover futuramente
user = input("Digite seu usuário: ")
password = input("Digite sua senha: ")

# my dict
excel = load_workbook('atestados.xlsx')
lista_atestados = excel['Plan1']

driver = webdriver.Chrome()
driver.get('https://senaconline-interno.sp.senac.br/psp/cs90pss/?cmd=login&languageCd=POR')

def colocar_assim_aparecer(tipo, nome, conteudo, timeout=15):
    try:
        elemento = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((tipo, nome))
        )
        driver.execute_script("document.activeElement.blur();")
        elemento.clear()
        elemento.send_keys(str(conteudo))
        print("✅ Conteúdo inserido com sucesso.")
    except Exception as e:
        pass

def clicar_assim_aparecer(tipo, nome, timeout=15):
    try:
        elemento = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((tipo, nome))
        )
        elemento.click()
        print("✅ Clicou com sucesso.")
    except Exception as e:
        pass

# Fazer login
colocar_assim_aparecer(By.ID, 'userid', user)
colocar_assim_aparecer(By.ID, 'pwd', password)
clicar_assim_aparecer(By.XPATH, '//*[@id="login"]/div/div[1]/div[8]/input')

# Acessando lista de atividades por aluno
clicar_assim_aparecer(By.XPATH, """//*[@id="pthnavbca_PORTAL_ROOT_OBJECT"]""")
clicar_assim_aparecer(By.XPATH, """//*[@id="fldra_HCSR_CURRICULUM_MANAGEMENT"]""")
clicar_assim_aparecer(By.XPATH, """//*[@id="fldra_HCSR_ATTENDANCE_ROSTER"]""")
clicar_assim_aparecer(By.XPATH, """//*[@id="crefli_HC_STDNT_ATTENDANCE_GBL"]/a""")

# Loop nos IDs da planilha
for row in lista_atestados.iter_rows(min_row=2, max_row=lista_atestados.max_row, min_col=1, max_col=lista_atestados.max_column):
    current_id = row[1].value
    current_inicio = row[2].value
    current_fim = row[3].value
    
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

        # Agora vamos clicar um a um nos elementos de aula
        print(f"🛠️ Encontradas {len(elementos_aulas)} aulas para clicar.")

        for i in range(len(elementos_aulas)):
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

                    # Inserir as datas e informações conforme o TODO
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

                    # clicar no botao de aplicar: //*[@id="DIG_APR_EST_WRK_PROCESS_BTN"]
                    clicar_assim_aparecer(By.XPATH, '//*[@id="DIG_APR_EST_WRK_PROCESS_BTN"]')
                    # clicar em botao de salvar: //*[@id="#ICSave"]
                    clicar_assim_aparecer(By.XPATH, '//*[@id="#ICSave"]')
                    # clicar no link de voltar para a lista: //*[@id="DERIVED_AA2_DERIVED_LINK10$0"]
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

    driver.switch_to.default_content()
    # Acessando lista de atividades por aluno
    clicar_assim_aparecer(By.XPATH, """//*[@id="pthnavbca_PORTAL_ROOT_OBJECT"]""")
    clicar_assim_aparecer(By.XPATH, """//*[@id="fldra_HCSR_CURRICULUM_MANAGEMENT"]""")
    clicar_assim_aparecer(By.XPATH, """//*[@id="fldra_HCSR_ATTENDANCE_ROSTER"]""")
    clicar_assim_aparecer(By.XPATH, """//*[@id="crefli_HC_STDNT_ATTENDANCE_GBL"]/a""")

print("🎉 Processamento concluído! Verifique o arquivo de log na pasta 'log'.")