from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import openpyxl
import time


def open_page_and_click_account():
    
    options = Options()
    options.add_argument('--ignore-certificate-errors')  
    driver = webdriver.Chrome(options=options)  

    # Abre o arquivo Excel e seleciona a primeira planilha
    workbook = openpyxl.load_workbook('Vld.xlsx')
    worksheet = workbook.worksheets[0]

    print('Digite o número da célula de início:')
    start_cell = int(get_user_input()) 

    
    print('Digite o número da célula de término:')
    end_cell = int(get_user_input())  

    # Define as variáveis de controle para o loop
    inicio = start_cell
    fim = end_cell

   
    driver.get('https://web.whatsapp.com/')
    time.sleep(50)  

    
    row = 1  # Usada para navegação nas mensagens ou contatos dentro do WhatsApp Web

    # Localiza e clica em uma conversa específica
    conversa = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div[3]/div[3]/div/div[3]/div[2]/div/div/div[1]/div/div/div/div[2]")
    conversa.click()
    time.sleep(3)  # Pausa para garantir que a ação seja realizada
    print("clicou na conversa")

    # Loop para percorrer as linhas de início a fim definidas pelo usuário
    while inicio <= fim:
        try:
            # Tenta localizar e clicar em um número específico (XPath fornecido)
            number = driver.find_element(By.XPATH, f"/html/body/div[1]/div/div/div[3]/div[4]/div/div[3]/div/div[2]/div[3]/div[14]/div/div/div[1]/div[1]/div[1]/div/div[1]/div/span[1]/span[{row}]/a")
            number.click()  # Clica no número encontrado
            print("clicou no numero")
            row += 1  # Incrementa a variável para passar para o próximo número

            time.sleep(4)  # Pausa para aguardar carregamento da página ou elemento

            # Tenta validar se existe um elemento específico indicando status no WhatsApp
            try:
                # Adiciona espera até que o elemento esteja presente (ou até 10 segundos)
                validar = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/span[5]/div/ul/div/li[1]/div"))
                )
                validarreturn = validar.text  # Obtém o texto do elemento após ser encontrado
                print(f"O texto obtido foi {validarreturn}")
                worksheet['B' + str(inicio)].value = "Whatsapp ok"  # Atualiza a célula na coluna B com o status
                worksheet['A' + str(inicio)].value = validarreturn  # Atualiza a célula na coluna A com o texto encontrado
                workbook.save('Vld.xlsx')  # Salva as alterações no arquivo Excel

            except:
                # Caso não encontre o elemento de validação, marca como "Sem whatsapp"
                worksheet['B' + str(inicio)].value = "Sem whatsapp"
                workbook.save('Vld.xlsx')  # Salva as alterações no arquivo Excel

            inicio += 1  # Incrementa para passar para a próxima linha
            workbook.save('Vld.xlsx')  # Salva as alterações no arquivo Excel
            print("planilha salva")

        except:
            # Tratamento de exceção genérico para quaisquer erros que possam ocorrer
            pass

# Função para obter entrada do usuário
def get_user_input():
    return input().strip()  # Lê a entrada do usuário e remove espaços em branco ao redor

# Chama a função principal para execução
open_page_and_click_account()
