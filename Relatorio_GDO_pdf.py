import win32com.client as win32  # Biblioteca para automação do Microsoft Excel
import datetime  # Biblioteca para manipulação de datas
import shutil  # Biblioteca para copiar e mover arquivos
import os  # Biblioteca para manipulação de arquivos e diretórios
import time  # Biblioteca para controle de tempo (pausas, delays)
import psutil  # Biblioteca para gerenciar processos do sistema

# ---------------------------- CONFIGURAÇÕES INICIAIS ----------------------------

# Caminho do arquivo Excel original (fonte de dados)
excel_path = r"E:\GDO\GDO_2025\Monitoramento_GDO_2025.xlsx"

# Gerar o nome do arquivo PDF dinamicamente com a data do dia
data_hoje = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")  # Formato: AAAA-MM-DD_HH-MM
pdf_output_path = f"E:\\GDO\\GDO_2025\\Monitoramento_GDO_{data_hoje}.pdf"

# Lista de abas que devem ser exportadas para o PDF na sequência desejada
abas_especificas = ["IMV", "ICVPe", "ICVPa", "POG", "IRTD", "PÓS_DELITO", "PVD", "EGRESSO", "CAVALO_ACO"]

# Criar um caminho temporário para evitar bloqueios no arquivo original
temp_excel_path = r"E:\GDO\GDO_2025\TEMP_Monitoramento_GDO_2025.xlsx"

# ---------------------------- FUNÇÃO PARA FECHAR O EXCEL ----------------------------
def fechar_excel():
    """
    Fecha qualquer instância do Microsoft Excel que esteja rodando no sistema.
    Isso evita que o arquivo esteja bloqueado para leitura/escrita antes de ser aberto no script.
    """
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        if proc.info['name'] and 'EXCEL.EXE' in proc.info['name'].upper():
            try:
                proc.kill()  # Mata o processo do Excel
            except psutil.NoSuchProcess:
                pass  # O processo já foi encerrado

# Fechar o Excel antes de rodar para garantir que o arquivo não esteja bloqueado
fechar_excel()

# Criar uma cópia do arquivo original para evitar problemas de bloqueio
shutil.copy(excel_path, temp_excel_path)

# ---------------------------- INICIAR O MICROSOFT EXCEL ----------------------------

# Inicializa o Excel via COM Object
excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False  # Mantém o Excel invisível durante a execução
excel.DisplayAlerts = False  # Desativa pop-ups e alertas do Excel para evitar interrupções

try:
    # ---------------------------- ABRIR O ARQUIVO EXCEL ----------------------------
    # Tenta abrir a cópia do arquivo em modo SOMENTE LEITURA para evitar conflitos
    workbook = excel.Workbooks.Open(temp_excel_path, ReadOnly=True)

    try:
        # ---------------------------- CRIAR NOVO ARQUIVO EXCEL TEMPORÁRIO ----------------------------
        # Criamos um novo arquivo no Excel para armazenar apenas as abas desejadas
        temp_workbook = excel.Workbooks.Add()

        # Verifica se o novo arquivo foi criado corretamente
        if temp_workbook is None:
            raise Exception("Falha ao criar um novo arquivo no Excel.")

        # ---------------------------- REMOVER ABAS PADRÃO DO NOVO ARQUIVO ----------------------------
        # O Excel cria um novo arquivo com abas padrão, então devemos excluí-las
        while temp_workbook.Sheets.Count > 1:
            temp_workbook.Sheets(1).Delete()  # Deleta cada aba extra até restar apenas uma

        # ---------------------------- COPIAR AS ABAS SELECIONADAS ----------------------------
        # Copiamos as abas desejadas do arquivo original para o novo arquivo
        for aba in abas_especificas:
            worksheet = workbook.Sheets(aba)  # Seleciona a aba específica
            worksheet.Copy(After=temp_workbook.Sheets(temp_workbook.Sheets.Count))  # Copia a aba para o final

        # ---------------------------- EXPORTAR COMO PDF ----------------------------
        # Converte as abas copiadas para um arquivo PDF no caminho definido
        temp_workbook.ExportAsFixedFormat(0, pdf_output_path)  # 0 representa o formato PDF

        # Mensagem de sucesso
        print(f"Relatório gerado com sucesso: {pdf_output_path}")

    finally:
        # ---------------------------- FECHAR OS ARQUIVOS ----------------------------
        # Fecha o arquivo temporário do Excel sem salvar alterações
        if 'temp_workbook' in locals():
            temp_workbook.Close(SaveChanges=False)
        
        # Fecha a cópia do arquivo original sem salvar alterações
        workbook.Close(SaveChanges=False)

except Exception as e:
    # Se houver erro, exibe a mensagem no terminal
    print(f"Erro ao gerar o relatório: {e}")

finally:
    # ---------------------------- FINALIZAR O EXCEL ----------------------------
    excel.Quit()  # Fecha o Excel completamente

    # ---------------------------- REMOVER ARQUIVO TEMPORÁRIO ----------------------------
    if os.path.exists(temp_excel_path):
        time.sleep(2)  # Pequena pausa para garantir que o arquivo não esteja em uso
        os.remove(temp_excel_path)  # Exclui o arquivo temporário após o uso
