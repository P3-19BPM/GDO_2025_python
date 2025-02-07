import win32com.client as win32  # Automação do Microsoft Excel
import datetime  # Manipulação de datas
import shutil  # Copiar e mover arquivos
import os  # Manipulação de arquivos e diretórios
import time  # Controle de tempo
import psutil  # Gerenciar processos do sistema
import webbrowser  # Abrir o PDF no navegador
from fpdf import FPDF  # Criar PDFs
from PyPDF2 import PdfMerger  # Mesclar PDFs
import subprocess

# ---------------------------- CONFIGURAÇÕES INICIAIS ----------------------------

# Caminho do arquivo Excel original
excel_path = r"E:\GDO\GDO_2025\Monitoramento_GDO_2025.xlsx"

# Gerar nomes dinâmicos para os arquivos
data_hoje = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
pdf_output_path = f"E:\\GDO\\GDO_2025\\GDO_{data_hoje}.pdf"
pdf_capa_path = f"E:\\GDO\\GDO_2025\\Capa_{data_hoje}.pdf"
pdf_final_path = f"E:\\GDO\\GDO_2025\\Monitoramento_GDO_{data_hoje}.pdf"

# Caminho da imagem de capa
imagem_capa_path = r"C:\Users\P3-19BPM\Downloads\Capa_Escura.png"

# Lista de abas que devem ser exportadas para o PDF
abas_especificas = ["IMV", "ICVPe", "ICVPa", "IRTD", "PÓS_DELITO",
                    "PVD", "EGRESSO", "CAVALO_ACO", "POG", "PP", "PP_Especifica"]

# Criar um caminho temporário para evitar bloqueios no arquivo original
temp_excel_path = r"E:\GDO\GDO_2025\TEMP_Monitoramento_GDO_2025.xlsx"

# ---------------------------- FECHAR EXCEL SE ESTIVER ABERTO ----------------------------


def fechar_excel():
    """ Fecha qualquer instância do Microsoft Excel rodando no sistema. """
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        if proc.info['name'] and 'EXCEL.EXE' in proc.info['name'].upper():
            try:
                proc.kill()
            except psutil.NoSuchProcess:
                pass  # O processo já foi encerrado

# ---------------------------- GERAR CAPA EM PDF ----------------------------


def gerar_capa_pdf(imagem_path, pdf_output):
    """ Cria um PDF contendo a imagem de capa. """
    pdf = FPDF()
    pdf.add_page()
    pdf.image(imagem_path, x=0, y=0, w=210, h=297)  # A4 (210x297 mm)
    pdf.output(pdf_output, "F")

# ---------------------------- EXPORTAR PLANILHA COMO PDF ----------------------------


def exportar_planilha_para_pdf():
    """ Usa o Excel para exportar as abas desejadas para um PDF. """
    fechar_excel()  # Garante que o Excel não esteja travado

    # Criar uma cópia do arquivo para evitar bloqueios
    shutil.copy(excel_path, temp_excel_path)

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        workbook = excel.Workbooks.Open(temp_excel_path, ReadOnly=True)

        # Criar um novo arquivo Excel temporário
        temp_workbook = excel.Workbooks.Add()

        # Remover abas padrão do novo arquivo
        while temp_workbook.Sheets.Count > 1:
            temp_workbook.Sheets(1).Delete()

        # Copiar apenas as abas desejadas
        for aba in abas_especificas:
            if aba in [sheet.Name for sheet in workbook.Sheets]:
                workbook.Sheets(aba).Copy(
                    After=temp_workbook.Sheets(temp_workbook.Sheets.Count))
            else:
                print(
                    f"⚠️  Aviso: A aba '{aba}' não existe no arquivo original.")

        # Exportar para PDF
        temp_workbook.ExportAsFixedFormat(0, pdf_output_path)

        # Fechar arquivos
        temp_workbook.Close(SaveChanges=False)
        workbook.Close(SaveChanges=False)

        print(f"✅ Relatório gerado com sucesso: {pdf_output_path}")

    except Exception as e:
        print(f"❌ Erro ao gerar o relatório: {e}")

    finally:
        excel.Quit()
        if os.path.exists(temp_excel_path):
            time.sleep(1)
            os.remove(temp_excel_path)

# ---------------------------- MESCLAR CAPA COM O PDF ----------------------------


def mesclar_pdfs():
    """ Mescla a capa com o PDF gerado e remove arquivos temporários. """
    try:
        gerar_capa_pdf(imagem_capa_path, pdf_capa_path)

        merger = PdfMerger()
        merger.append(pdf_capa_path)
        merger.append(pdf_output_path)
        merger.write(pdf_final_path)
        merger.close()

        print(f"📄 Relatório final gerado: {pdf_final_path}")

        os.remove(pdf_capa_path)
        os.remove(pdf_output_path)

    except Exception as e:
        print(f"❌ Erro ao mesclar PDFs: {e}")

# ---------------------------- ABRIR O PDF NO CHROME ----------------------------


def abrir_pdf_no_chrome():
    """ Abre o PDF gerado no navegador Google Chrome. """
    try:
        chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        
        if os.path.exists(chrome_path):
            subprocess.Popen([chrome_path, pdf_final_path], shell=True)
        else:
            webbrowser.open(pdf_final_path)  # Abre no navegador padrão

        print("🌍 PDF aberto no Chrome.")

    except Exception as e:
        print(f"❌ Erro ao abrir o PDF: {e}")


# ---------------------------- EXECUÇÃO PRINCIPAL ----------------------------
if __name__ == "__main__":
    exportar_planilha_para_pdf()
    mesclar_pdfs()
    abrir_pdf_no_chrome()
