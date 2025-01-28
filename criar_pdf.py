from fpdf import FPDF

# Criando o PDF
pdf = FPDF()
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_page()
pdf.set_font("Arial", size=12)

# Título
pdf.set_font("Arial", style='B', size=14)
pdf.cell(200, 10, txt="Automatização de Envio de Mensagens na Intranet", ln=True, align='C')
pdf.ln(10)

# Conteúdo
conteudo = """
## Automação com Python no Jupyter Notebook

### 1. Instalar Dependências (Se Necessário)
!pip install selenium pandas openpyxl

### 2. Configurar o Código no Jupyter Notebook
#### 1. Importar Bibliotecas
```python
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time

# Carregar os dados da tabela Excel
caminho_excel = "/mnt/data/image.png"  # Substitua pelo caminho correto do seu arquivo Excel
df = pd.read_excel(caminho_excel)

# Verificar os dados carregados
df.head()

# Configuração do Selenium WebDriver
driver = webdriver.Chrome()  # Certifique-se de ter o driver do Chrome configurado
driver.get("https://pa.policiamilitar.mg.gov.br/#/escrever")

# Realizar login na intranet
username_field = driver.find_element(By.ID, "id_do_campo_usuario")  # Substitua pelo ID do campo de usuário
password_field = driver.find_element(By.ID, "id_do_campo_senha")  # Substitua pelo ID do campo de senha
login_button = driver.find_element(By.ID, "id_do_botao_login")  # Substitua pelo ID do botão de login

username_field.send_keys("seu_usuario")
password_field.send_keys("sua_senha")
login_button.click()

time.sleep(3)  # Aguarde o carregamento

# Realizar login na intranet
username_field = driver.find_element(By.ID, "id_do_campo_usuario")  # Substitua pelo ID do campo de usuário
password_field = driver.find_element(By.ID, "id_do_campo_senha")  # Substitua pelo ID do campo de senha
login_button = driver.find_element(By.ID, "id_do_botao_login")  # Substitua pelo ID do botão de login

username_field.send_keys("seu_usuario")
password_field.send_keys("sua_senha")
login_button.click()

time.sleep(3)  # Aguarde o carregamento


# Iterar pelos funcionários na tabela
for _, row in df.iterrows():
    masp = row['MASP']  # Substitua pelo nome correto da coluna
    doc_num = row['NÚMERO_DOCUMENTO']  # Substitua pelo nome correto da coluna

    # Preencher o campo "Para"
    to_field = driver.find_element(By.ID, "id_campo_para")  # Substitua pelo ID correto
    to_field.send_keys(str(masp))  # Certifique-se de enviar o MASP como string
    time.sleep(1)
    to_field.send_keys(Keys.ARROW_DOWN)
    to_field.send_keys(Keys.ENTER)

    # Preencher o campo "Assunto"
    subject_field = driver.find_element(By.ID, "id_campo_assunto")  # Substitua pelo ID correto
    subject_field.send_keys("Assunto Padrão")

    # Preencher o corpo da mensagem
    message_field = driver.find_element(By.ID, "id_campo_mensagem")  # Substitua pelo ID correto
    message_text = f"Texto padrão.\nPor favor, altere o documento número: {doc_num}"
    message_field.send_keys(message_text)

    # Enviar a mensagem
    send_button = driver.find_element(By.ID, "id_botao_enviar")  # Substitua pelo ID correto
    send_button.click()

    time.sleep(2)  # Aguarde para evitar problemas de sobrecarga

    # Finalizar o navegador
driver.quit()
"""