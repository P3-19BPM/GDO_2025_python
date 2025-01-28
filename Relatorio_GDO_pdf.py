import win32com.client as win32

# Caminho do arquivo Excel
excel_path = r"E:\GDO\GDO_2025\Monitoramento_GDO_2025.xlsx"
pdf_output_path = r"E:\Painel_PowerBI\BASE_GDO\BD_2025\Relatorio_GDO_Abas_Especificas.pdf"

# Lista de abas específicas que você deseja incluir no PDF
abas_especificas = ["IMV", "ICVPe", "ICVPa"]

# Abrir o Excel
excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False  # Mantenha o Excel invisível

# Abrir a pasta de trabalho
workbook = excel.Workbooks.Open(excel_path)

try:
    # Criar um novo arquivo temporário para as abas selecionadas
    temp_workbook = excel.Workbooks.Add()

    # Copiar as abas específicas para o novo arquivo
    for aba in abas_especificas:
        worksheet = workbook.Sheets(aba)
        worksheet.Copy(Before=temp_workbook.Sheets(1))  # Copia a aba para o novo arquivo
    
    # Remover a aba padrão criada no novo arquivo (se existir)
    if temp_workbook.Sheets.Count > len(abas_especificas):
        temp_workbook.Sheets(1).Delete()

    # Exportar o novo arquivo como PDF
    temp_workbook.ExportAsFixedFormat(0, pdf_output_path)  # 0 = PDF
    print(f"Relatório gerado em: {pdf_output_path}")
finally:
    # Fechar os arquivos
    temp_workbook.Close(SaveChanges=False)
    workbook.Close(SaveChanges=False)
    excel.Quit()
