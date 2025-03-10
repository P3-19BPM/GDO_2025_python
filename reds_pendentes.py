from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import pyautogui
import os
import pyperclip
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium.webdriver.common.keys import Keys
import time
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import matplotlib.pyplot as plt


# Caminho para a pasta "User Data" do Chrome
user_data_dir = r"C:\Users\P3-19BPM\AppData\Local\Google\Chrome\User Data"

# Nome do perfil que você deseja usar (por exemplo, "Profile 1" ou "Default")
profile_directory = "Default"

# Configurações do Chrome
options = webdriver.ChromeOptions()
options.add_argument(f"--user-data-dir={user_data_dir}")  # Especifica o diretório do perfil
options.add_argument(f"--profile-directory={profile_directory}")  # Especifica o perfil

# Inicializa o Chrome com o perfil selecionado
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Acessa o site
driver.get("https://intranet.policiamilitar.mg.gov.br/autenticacaosso/login.jsf#b")

# Verifica se o perfil correto foi carregado
print("Perfil carregado:", profile_directory)

# Aguarda alguns segundos para a página carregar completamente
time.sleep(5)

# Pressiona a tecla "Enter" usando pyautogui
pyautogui.press('enter')

time.sleep(5)

# CLICA NA FOTO PARA ESCOLHER A CAIXA DE MENSAGEM
try:
    imagem = driver.find_element("xpath", '//*[@id="u1"]/img')
    imagem.click()
    print("Imagem clicada com sucesso!")
except Exception as e:
    print(f"Erro ao clicar na imagem: {e}")

time.sleep(5)

# CLICA NA CAIXA DE ENTRADA
try:
    elemento1 = driver.find_element("xpath", '//*[@id="detalheUsr"]/div/div/a[2]')
    elemento1.click()
    print("Primeiro elemento clicado com sucesso!")
except Exception as e:
    print(f"Erro ao clicar no primeiro elemento: {e}")

time.sleep(5)

# CLICA NA CAIXA DA P3/19 BPM
try:
    elemento2 = driver.find_element("xpath", '//*[@id="c-25866"]')
    elemento2.click()
    print("Segundo elemento clicado com sucesso!")
except Exception as e:
    print(f"Erro ao clicar no segundo elemento: {e}")

time.sleep(5)

# CLICA NO ÍCONE ESCREVER
try:
    elemento3 = driver.find_element("xpath", '//*[@id="botao-escrever"]')
    elemento3.click()
    print("Terceiro elemento clicado com sucesso!")
except Exception as e:
    print(f"Erro ao clicar no terceiro elemento: {e}")

# Aguarda mais alguns segundos para garantir que a página esteja completamente carregada
time.sleep(8)

# Definindo o caminho do arquivo CSV
caminho_arquivo = r"C:\Users\PC 10066066-5\Downloads\data.csv"  # Substitua pelo caminho correto do seu arquivo

# Lê o arquivo CSV e verifica se a coluna E tem células não vazias
df = pd.read_csv(caminho_arquivo, encoding='utf-8', delimiter=',')  # Ajuste o delimitador se necessário

# Exibir as primeiras linhas do dataframe para conferência
print(df.head())

# Verificando o tipo
print(df.dtypes)  # Confirma se virou int

# Verifica o número de linhas na coluna E (índice 4) que não são NaN
num_linhas_digitador = df.shape[0]  # Número total de linhas no DataFrame

# Inicializa variáveis para armazenar os últimos números copiados
numeros_copiados_digitador = set()
numeros_copiados_comandante = set()

# Verifica se ainda há células não vazias para o comandante de pelotão
linha_digitador = 0  # Começa na linha 0, supondo que seja a mesma coluna ou a próxima
while linha_digitador < num_linhas_digitador and pd.notna(df.iloc[linha_digitador, 4]):
    try:
        numero_policia_digitador = str(df.iloc[linha_digitador, 4])

        # COPIANDO NUMERO DE POLÍCIA DO DIGITADOR
         # Verifica se o número já foi copiado anteriormente
        if numero_policia_digitador not in numeros_copiados_digitador:
            pyperclip.copy(numero_policia_digitador)
            ultimo_numero_digitador = numero_policia_digitador
            print(f"Número de Polícia do digitador copiado: {numero_policia_digitador}")

            # Adiciona o número ao conjunto para evitar duplicatas
            numeros_copiados_digitador.add(numero_policia_digitador)
        

       # Espera até que o elemento esteja disponível e clicável
        campo_colar_digitador = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="cdk-drop-list-0"]'))
    )

        # Rola a página até o elemento para garantir que ele está visível
        driver.execute_script("arguments[0].scrollIntoView();", campo_colar_digitador)

        # Clica no elemento usando JavaScript para evitar problemas de interceptação
        driver.execute_script("arguments[0].click();", campo_colar_digitador)
        
        time.sleep(2)  # Aguarda um pouco para o campo responder

        # Copia o número para a área de transferência (já deve ter sido feito antes)
        pyperclip.copy(numero_policia_digitador)

        # Cola o conteúdo da área de transferência no campo ativo
        driver.switch_to.active_element.send_keys(Keys.CONTROL, 'v')
        print("Número do comandante colado com sucesso usando Selenium.")
        
        # Aguarda um pouco antes de pressionar down e enter
        time.sleep(2)
        driver.switch_to.active_element.send_keys(Keys.ARROW_DOWN)
        driver.switch_to.active_element.send_keys(Keys.ENTER)
        time.sleep(4)  # Aguarda a ação ser processada
    
        # Se for o último valor válido, pressiona Tab duas vezes
        if linha_digitador == num_linhas_digitador - 1:
            pyautogui.press('tab')
            pyautogui.press('tab')

            print("Pressionado Tab duas vezes após o último digitador.")

        # Incrementa a linha para a próxima iteração
        linha_digitador += 1

    except Exception as e:
        print(f"Erro durante o processo de automação para digitador: {e}")
        break  # Sai do loop se ocorrer um erro


# Verifica o número de linhas na coluna G (índice 6) que não são NaN
num_linhas_comandante = df.shape[0]  # Supondo que a coluna do comandante tenha a mesma quantidade de linhas

# Verifica se ainda há células não vazias para o comandante de pelotão
linha_comandante = 0  # Começa na linha 0, supondo que seja a mesma coluna ou a próxima
while linha_comandante < num_linhas_comandante and pd.notna(df.iloc[linha_comandante, 7]):
    try:
        numero_policia_comandante = str(df.iloc[linha_comandante, 7])
        if numero_policia_comandante in numeros_copiados_comandante:
            linha_comandante+=1
            continue
        
        # COPIANDO NUMERO DE POLÍCIA DO COMANDANTE DE PELOTÃO
        # Verifica se o número já foi copiado anteriormente
        pyperclip.copy(numero_policia_comandante)
        ultimo_numero_comandante = numero_policia_comandante
        print(f"Número de Polícia do comandante copiado: {numero_policia_comandante}")

        # Adiciona o número ao conjunto para evitar duplicatas
        numeros_copiados_comandante.add(numero_policia_comandante)

       # Espera até que o elemento esteja disponível e clicável
        campo_colar_comandante = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="cdk-drop-list-1"]'))
    )

        # Rola a página até o elemento para garantir que ele está visível
        driver.execute_script("arguments[0].scrollIntoView();", campo_colar_comandante)

        # Clica no elemento usando JavaScript para evitar problemas de interceptação
        driver.execute_script("arguments[0].click();", campo_colar_comandante)
        
        time.sleep(2)  # Aguarda um pouco para o campo responder

        # Copia o número para a área de transferência (já deve ter sido feito antes)
        pyperclip.copy(numero_policia_comandante)

        # Cola o conteúdo da área de transferência no campo ativo
        driver.switch_to.active_element.send_keys(Keys.CONTROL, 'v')
        print("Número do comandante colado com sucesso usando Selenium.")
        
        # Aguarda um pouco antes de pressionar down e enter
        time.sleep(2)
        driver.switch_to.active_element.send_keys(Keys.ARROW_DOWN)
        driver.switch_to.active_element.send_keys(Keys.ENTER)
        time.sleep(4)  # Aguarda a ação ser processada
    
        # Se for o último valor válido, pressiona Tab duas vezes
        if linha_comandante == num_linhas_comandante - 1:
            pyautogui.press('tab')
            pyautogui.press('tab')

            print("Pressionado Tab duas vezes após o último digitador.")

        # Incrementa a linha para a próxima iteração
        linha_comandante += 1

    except TimeoutException as toe:
        print(f"O elemento demorou demais para ficar disponível. Erro: {toe}")
        break  # Sai do loop se ocorrer um erro
    except Exception as e:
        print(f"Erro durante o processo de automação para comandante: {e}")
        break  # Sai do loop se ocorrer um erro

# Etapa final para colar o "Título da Mensagem" e o conteúdo da mensagem
try:
    # COLANDO O "TÍTULO DE MENSAGEM"
    texto_titulo = "REGISTROS EM ABERTO - ENCERRAMENTO DE REDS UU"
    pyperclip.copy(texto_titulo)
    print("Texto do título copiado para a área de transferência com sucesso.")

    # Espera até que o campo do título esteja disponível e clicável
    campo_titulo = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="assunto-txt"]'))
    )
    # Rola a página até o campo e utiliza clique via JavaScript para evitar interferências
    driver.execute_script("arguments[0].scrollIntoView(true);", campo_titulo)
    driver.execute_script("arguments[0].click();", campo_titulo)
    time.sleep(1)  # Aguarda um pouco antes de colar
    driver.switch_to.active_element.send_keys(texto_titulo)
    print("Texto do título colado com sucesso no campo desejado.")
    time.sleep(4)  # Aguarda a página processar a ação

    # COLANDO A MENSAGEM
    texto_conteudo = """Caro Digitador da Ocorrência

Existe REDS em aberto sob sua responsabilidade. Por isso, incumbiu-me o Sr. Chefe da P3 / 27 BPM de encaminhar mensagem recomendando ENCERRAMENTO UU"!!!

Observação: O Comando do 27o BPM recomenda que os Comandos de Companhias inste os relatores ao cumprimento das diretrizes constantes no BOLETIM TÉCNICO No 02 / 2016, BEM COMO RESPONSABILIZAR ADMINISTRATIVAMENTE OS DIGITADORES contumazes.

===================================================================================
BOLETIM TÉCNICO No 02 / 2016 – DAOp/Cinds (BGPM No 12 de 16 de Fevereiro de 2016)

"Do Encerramento de REDS

Diante disso, a recomendação técnica é para que o policial militar registre o REDS de todas as atendidas ou integradas para o militar, durante o turno de serviço, salvo em casos justificadamente comprovados e autorizado pelo CPU/CPCia."

===================================================================================
"""
    pyperclip.copy(texto_conteudo)

    # Espera até que o campo de conteúdo esteja disponível e clicável
    campo_conteudo = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="conteudo-txt"]'))
    )
    campo_conteudo.click()
    time.sleep(1)  # Aguarda um pouco antes de colar
    driver.switch_to.active_element.send_keys(texto_conteudo)
    print("Texto da mensagem colado com sucesso no campo de conteúdo.")
    time.sleep(4)  # Aguarda a página processar a ação

# ETAPA DE CARREGAR IMAGEM

except Exception as e:
    print(f"Erro ao colar o título e a mensagem: {e}")

# Caminho do arquivo CSV
csv_path = r'C:\Users\PC 10066066-5\Downloads\data.csv'

# Carregar os dados
df = pd.read_csv(csv_path)

# Criar a figura e os eixos
fig, ax = plt.subplots(figsize=(10, 6))
ax.axis('tight')
ax.axis('off')

# Criar a tabela
table = ax.table(cellText=df.values, colLabels=df.columns, cellLoc='center', loc='center', colColours=["#003366"] * len(df.columns))

# Ajustar estilo do texto do cabeçalho
for key, cell in table.get_celld().items():
    if key[0] == 0:  # Cabeçalho (linha 0)
        cell.set_text_props(color="white", weight="bold")  # Texto branco e negrito
        cell.set_fontsize(12)  # Ajusta o tamanho da fonte

# Ajustar estilo geral da tabela
table.auto_set_font_size(False)
table.set_fontsize(10)
table.auto_set_column_width(col=list(range(len(df.columns))))

# Salvar como PNG
output_path = r'C:\Users\PC 10066066-5\Downloads\tabela.png'
plt.savefig(output_path, bbox_inches='tight', dpi=300)


# CLICA NO ÍCONE CARREGAR IMAGEM
try:
    elemento3 = driver.find_element("xpath", '//*[@id="iconeAdicionarImagem"]')
    elemento3.click()
    print("Terceiro elemento clicado com sucesso!")
except Exception as e:
    print(f"Erro ao clicar no terceiro elemento: {e}")

# Aguarda mais alguns segundos para garantir que a página esteja completamente carregada
time.sleep(4)

# Abrir a pasta Downloads e carregar imagem
pyautogui.write(r'C:\Users\PC 10066066-5\Downloads\tabela.png')
time.sleep(2)
pyautogui.press('enter')
time.sleep(3)

# Fechar caixa de mensagem aberta após carregamento da imagem
try:
        botao_enviar = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="abaMens"]/app-ferramentas-markdown/app-visualizar-markdown/div[2]/h1/a'))
    )
        botao_enviar.click()
        print("Botão de enviar clicado com sucesso.")
        time.sleep(2)  # Aguarda a ação de envio ser processada

except Exception as e:
    print(f"Erro ao clicar no terceiro elemento: {e}")

    # ENVIANDO MENSAGEM
try:
        botao_enviar = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="botaoEnviar"]'))
    )
        botao_enviar.click()
        print("Botão de enviar clicado com sucesso.")
        time.sleep(4)  # Aguarda a ação de envio ser processada

except TimeoutException as toe:
    print(f"Timeout ao esperar pelo elemento. Erro: {toe}")
except NoSuchElementException as nsee:
    print(f"Elemento não encontrado. Erro: {nsee}")
except Exception as e:
    print(f"Erro inesperado durante o processo de automação final: {e}")

# Fecha o navegador ao finalizar todos os loops
driver.quit()