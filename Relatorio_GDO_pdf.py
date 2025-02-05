import win32com.client as win32  # Biblioteca para automação do Microsoft Excel
import datetime  # Biblioteca para manipulação de datas
import shutil  # Biblioteca para copiar e mover arquivos
import os  # Biblioteca para manipulação de arquivos e diretórios
import time  # Biblioteca para controle de tempo (pausas, delays)
import psutil  # Biblioteca para gerenciar processos do sistema
from fpdf import FPDF  # Biblioteca para criar PDFs com a capa
from PyPDF2 import PdfMerger  # Biblioteca para mesclar PDFs

# ---------------------------- CONFIGURAÇÕES INICIAIS ----------------------------

# Caminho do arquivo Excel original (fonte de dados)
excel_path = r"E:\GDO\GDO_2025\Monitoramento_GDO_2025.xlsx"

# Gerar o nome do arquivo PDF dinamicamente com a data do dia
data_hoje = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")  # Formato: AAAA-MM-DD_HH-MM
pdf_output_path = f"E:\\GDO\\GDO_2025\\GDO_{data_hoje}.pdf"
pdf_capa_path = f"E:\\GDO\\GDO_2025\\Capa_{data_hoje}.pdf"
pdf_final_path = f"E:\\GDO\\GDO_2025\\Monitoramento_GDO_{data_hoje}.pdf"

# Caminho da imagem de capa
imagem_capa_path = r"C:\Users\P3-19BPM\Downloads\Capa_Escura.png"

# Lista de abas que devem ser exportadas para o PDF na sequência desejada
abas_especificas = ["IMV", "ICVPe", "ICVPa", "IRTD", "PÓS_DELITO", "PVD", "EGRESSO", "CAVALO_ACO", "POG", "PP", "PP_Especifica"]

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

# ---------------------------- GERAR A CAPA EM PDF ----------------------------
def gerar_capa_pdf(imagem_path, pdf_output):
    """
    Gera um PDF contendo apenas a imagem de capa.
    """
    pdf = FPDF()
    pdf.add_page()
    pdf.image(imagem_path, x=0, y=0, w=210, h=297)  # A4 em mm (210x297)
    pdf.output(pdf_output, "F")
    print(f"Capa gerada com sucesso: {pdf_output}")

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

# ---------------------------- MESCLAR A CAPA E O RELATÓRIO ----------------------------
try:
    # Gerar a capa em PDF
    gerar_capa_pdf(imagem_capa_path, pdf_capa_path)

    # Criar um objeto para mesclar os PDFs
    merger = PdfMerger()
    
    # Adicionar a capa primeiro
    merger.append(pdf_capa_path)
    
    # Adicionar o relatório gerado pelo Excel
    merger.append(pdf_output_path)
    
    # Salvar o PDF final com a capa
    merger.write(pdf_final_path)
    merger.close()
    
    print(f"Relatório final gerado com sucesso: {pdf_final_path}")

    # Remover arquivos temporários (capa e PDF original)
    os.remove(pdf_capa_path)
    os.remove(pdf_output_path)

except Exception as e:
    print(f"Erro ao mesclar PDFs: {e}")
