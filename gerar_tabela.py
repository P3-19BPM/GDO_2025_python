import pandas as pd
import openpyxl

# Estruturar os dados para cada tópico como uma tabela separada
data_tables = {
    "Armas de Fogo Apreendidas": {
        "Descrição": ["Armas de fogo apreendidas", "Presos", "Resultado em 2024", "Aumento (%)"],
        "Valores": [18, 18, 8, 125]
    },
    "Cumprimento de Mandados de Prisão": {
        "Descrição": ["Mandados cumpridos", "Resultado em 2024", "Aumento (%)"],
        "Valores": [16, 10, 60]
    },
    "Tráfico de Drogas": {
        "Descrição": ["Registros de tráfico", "Presos por tráfico", "Resultado em 2024", "Aumento (%)"],
        "Valores": [32, 28, 25, 28]
    },
    "Material Apreendido": {
        "Descrição": [
            "Pedras de crack", "Buchas de maconha (incluindo porções grandes)",
            "Papelotes de cocaína (+ barra prensada)", "Comprimidos de ecstasy",
            "Dinheiro apreendido (R$)", "Celulares apreendidos"
        ],
        "Valores": [409, 458, 725, 180, 3432.45, 24]
    }
}

# Sugestões de gráficos para cada tópico
chart_suggestions = {
    "Armas de Fogo Apreendidas": "Gráfico de colunas: Comparação entre anos.",
    "Cumprimento de Mandados de Prisão": "Gráfico de barras: Comparação de mandados cumpridos.",
    "Tráfico de Drogas": "Gráfico de colunas: Comparação de registros e presos.",
    "Material Apreendido": "Gráfico de pizza: Distribuição do tipo de material apreendido."
}

# Caminho para o arquivo Excel
file_path = "Produtividade_Jan_2025_Tabelas_e_Graficos.xlsx"

# Criar e salvar tabelas e sugestões de gráficos no Excel
with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
    # Salvar cada tabela em uma aba separada
    for topic, table_data in data_tables.items():
        df_topic = pd.DataFrame(table_data)
        # Limitar o nome da aba para 31 caracteres devido à restrição do Excel
        sheet_name = topic[:31]
        df_topic.to_excel(writer, sheet_name=sheet_name, index=False)

    # Criar DataFrame para sugestões de gráficos
    chart_df = pd.DataFrame(chart_suggestions.items(), columns=[
                            "Tópico", "Sugestão de Gráfico"])
    # Salvar sugestões de gráficos em uma aba separada
    chart_df.to_excel(writer, sheet_name="Sugestões de Gráficos", index=False)

print(f"Arquivo Excel gerado: {file_path}")
