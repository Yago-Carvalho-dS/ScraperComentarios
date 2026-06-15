import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Configurações do Selenium
chrome_options = Options()
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

import os
import glob

def process_excel_files(directory, output_file="comentarios_consolidados.csv"):
    """
    Lê os arquivos Excel baixados, filtra colunas com 'comente' e salva.
    """
    print("\n--- Iniciando Processamento Pandas ---")
    files = glob.glob(os.path.join(directory, "*.xlsx"))
    all_data = []

    for file in files:
        try:
            df = pd.read_excel(file)
            # Filtra colunas que contêm 'comente' (case insensitive)
            cols_to_keep = [col for col in df.columns if 'comente' in str(col).lower()]
            
            if cols_to_keep:
                filtered_df = df[cols_to_keep].copy()
                filtered_df['projeto_origem'] = os.path.basename(file)
                all_data.append(filtered_df)
                print(f"  [Pandas] Extraídas {len(cols_to_keep)} colunas de {file}")
        except Exception as e:
            print(f"  [Erro Pandas] Falha ao processar {file}: {e}")

    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        final_df.to_csv(output_file, index=False, encoding='utf-8-sig')
        print(f"\n[Sucesso] Arquivo consolidado gerado: {output_file}")
    else:
        print("\n[Aviso] Nenhum comentário encontrado nos arquivos.")

def main():
    try:
        # --- CONFIGURAÇÕES DE SELETORES ---
        LOGIN_URL = "https://survey.fgv.br/login/identity-provider-select?path=%2Fsurvey-builder%2FSV_d9UOpJLoiapfYTc%2Fedit&stateID=f074b9a8-0fd8-41be-b135-85ee8fd4048d"
        PROJECTS_URL = "https://survey.fgv.br/Q/MyProjectsSection"
        
        # Seletores da Lista de Projetos
        SEL_THREE_DOTS_ALL = "button[id^='options-SV_']" # Seleciona todos os botões de 3 pontinhos
        SEL_NEXT_PAGE = "button[aria-label='Próxima página']"
        
        # Seletores do Fluxo de Exportação
        SEL_DATA_ANALYSIS_XPATH = "//*[@role='menuitem' and contains(., 'Dados e análise')]"
        SEL_EXPORT_DROPDOWN = "[data-testid='export-and-import-menu']"
        SEL_EXPORT_BTN_INNER = "//div[@role='menuitem' and contains(., 'Exportar dados')]"
        SEL_EXCEL_TAB = "#export-data-modal-button-group-tab-excel"
        SEL_MORE_OPTIONS = "[data-testid='export-more-options-btn']"
        SEL_REMOVE_LINE_BREAKS = "[data-testid='export-replace-newline-checkbox']"
        SEL_DOWNLOAD_BTN = "._a2bbw._M5ck0" # Botão 'Baixar' fornecido

        # --- FASE 1: LOGIN ---
        print("Abrindo o Qualtrics FGV para LOGIN...")
        driver.get(LOGIN_URL)
        input("Após fazer o LOGIN e estar na tela inicial, pressione ENTER aqui para iniciar a automação completa...")
        
        wait = WebDriverWait(driver, 30)
        page = 1

        while True:
            print(f"\n>>> PROCESSANDO PÁGINA {page} <<<")
            driver.get(PROJECTS_URL)
            time.sleep(5)
            
            # Navega para a página correta se não for a primeira
            if page > 1:
                for p in range(1, page):
                    try:
                        next_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, SEL_NEXT_PAGE)))
                        next_btn.click()
                        time.sleep(3)
                    except:
                        print(f"Não foi possível avançar para a página {page}. Encerrando.")
                        return

            # Identifica todos os formulários na página atual
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, SEL_THREE_DOTS_ALL)))
            survey_menus = driver.find_elements(By.CSS_SELECTOR, SEL_THREE_DOTS_ALL)
            total_on_page = len(survey_menus)
            print(f"Detectados {total_on_page} formulários nesta página.")

            for i in range(total_on_page):
                # Recarrega a lista para evitar elementos obsoletos (stale elements)
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, SEL_THREE_DOTS_ALL)))
                current_menus = driver.find_elements(By.CSS_SELECTOR, SEL_THREE_DOTS_ALL)
                target_menu = current_menus[i]
                survey_id = target_menu.get_attribute("id")
                
                print(f"\n[{i+1}/{total_on_page}] Processando Formulário: {survey_id}")

                try:
                    # 1. Clicar nos 3 pontinhos
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_menu)
                    time.sleep(1)
                    target_menu.click()
                    
                    # 2. Ir para "Dados e Análise"
                    btn_analysis = wait.until(EC.element_to_be_clickable((By.XPATH, SEL_DATA_ANALYSIS_XPATH)))
                    btn_analysis.click()
                    
                    # Aguarda carregamento da página de dados (pode demorar em surveys grandes)
                    time.sleep(8) 

                    # 3. Clicar em Exportar Dados dropdown
                    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, SEL_EXPORT_DROPDOWN))).click()

                    # 4. Clicar no botão Exportar dentro do dropdown
                    wait.until(EC.element_to_be_clickable((By.XPATH, SEL_EXPORT_BTN_INNER))).click()

                    # 5. Selecionar Formato Excel
                    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, SEL_EXCEL_TAB))).click()

                    # 6. Clicar em 'Mais Opções'
                    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, SEL_MORE_OPTIONS))).click()
                    # 7. Marcar 'Remover quebra de linhas'
                    print("Verificando checkbox 'Remover quebra de linhas'...")
                    try:
                        # Tenta encontrar o input real dentro do label para verificar o estado
                        checkbox_input = driver.find_element(By.CSS_SELECTOR, SEL_REMOVE_LINE_BREAKS + " input")
                        if not checkbox_input.is_selected():
                            driver.execute_script("arguments[0].click();", checkbox_input)
                    except:
                        # Fallback: clica no label se não achar o input
                        checkbox_label = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, SEL_REMOVE_LINE_BREAKS)))
                        driver.execute_script("arguments[0].click();", checkbox_label)

                    # 8. Botão Baixar Final
                    print(f"Iniciando download de {survey_id}...")
                    # Usando XPath para garantir que clicamos no botão que diz 'Baixar' e ignorar outros
                    SEL_DOWNLOAD_FINAL_XPATH = "//button[contains(., 'Baixar')] | //button[text()='Baixar']"
                    try:
                        download_btn = wait.until(EC.presence_of_element_located((By.XPATH, SEL_DOWNLOAD_FINAL_XPATH)))
                        driver.execute_script("arguments[0].click();", download_btn)
                    except:
                        # Fallback para a classe anterior se o XPath falhar
                        download_btn = driver.find_element(By.CSS_SELECTOR, SEL_DOWNLOAD_BTN)
                        driver.execute_script("arguments[0].click();", download_btn)

                    # Espera o download ser processado pelo navegador
                    print("Aguardando início do download...")
                    time.sleep(10) 

                except Exception as e:
                    print(f"Erro ao processar {survey_id}: {e}")
                
                # Volta para a lista de projetos para o próximo
                driver.get(PROJECTS_URL)
                time.sleep(3)
                # Re-navega para a página atual
                if page > 1:
                    for p in range(1, page):
                        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, SEL_NEXT_PAGE))).click()
                        time.sleep(2)

            # Verifica se há uma próxima página (Lógica estrita)
            try:
                next_buttons = driver.find_elements(By.CSS_SELECTOR, SEL_NEXT_PAGE)
                if next_buttons:
                    btn = next_buttons[0]
                    # Verifica se o botão está desabilitado por classe, atributo ou aria-disabled
                    is_disabled = (
                        not btn.is_enabled() or 
                        "disabled" in btn.get_attribute("class").lower() or 
                        btn.get_attribute("aria-disabled") == "true"
                    )
                    
                    if not is_disabled:
                        print("\nPróxima página detectada. Avançando...")
                        page += 1
                    else:
                        print("\n[FIM] Botão de próxima página está desabilitado. Fim da lista.")
                        break
                else:
                    print("\n[FIM] Botão de próxima página não encontrado. Fim da lista.")
                    break
            except Exception as e:
                print(f"Erro ao verificar paginação: {e}")
                break

    except Exception as e:
        print(f"\nOcorreu um erro crítico: {e}")
    finally:
        print("\nProcesso finalizado. O navegador permanecerá aberto.")

if __name__ == "__main__":
    main()
