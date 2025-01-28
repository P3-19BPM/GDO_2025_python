import pandas as pd
import openpyxl

# Dados estruturados para exportar ao Excel
data = {
    "Tópico": [
        "1. Armas de Fogo Apreendidas",
        "1.1 19º BPM",
        "1.2 Teófilo Otoni",
        "1.3 Resultado em Teófilo Otoni no mesmo período em 2024",
        "2. Cumprimento de Mandados de Prisão",
        "2.1 19º BPM",
        "2.2 Teófilo Otoni",
        "2.3 Resultado em Teófilo Otoni no mesmo período em 2024",
        "3. Tráfico de Drogas",
        "3.1 19º BPM",
        "3.1.1 Presos por Tráfico",
        "3.2 Teófilo Otoni",
        "3.2.1 Presos por Tráfico",
        "3.3 Resultado em Teófilo Otoni no mesmo período em 2024",
        "3.4 Total de Material Apreendido em Teófilo Otoni"
    ],
    "Descrição": [
        None, 
        "37", 
        "18 (Presos: 18)", 
        "8 (Aumento de 125%)", 
        None, 
        "28", 
        "16", 
        "10 (Aumento de 60%)", 
        None, 
        "35", 
        "37", 
        "32", 
        "28", 
        "25 (Aumento de 28%)", 
        "409 Pedras de Crack; 458 Buchas de Maconha (+10 porções grandes); "
        "725 Papelotes de Cocaína (+1 barra prensada); 180 comprimidos de Ecstasy; "
        "R$ 3432,45; 24 Celulares"
    ]
}

# Criar DataFrame
df = pd.DataFrame(data)

# Exportar para Excel
file_path = "Produtividade_Jan_2025.xlsx"
df.to_excel(file_path, index=False, engine='openpyxl')

# Sugestões de gráficos
chart_suggestions = {
    "1. Armas de Fogo Apreendidas": "Gráfico de colunas: Comparação entre as unidades e períodos.",
    "2. Cumprimento de Mandados de Prisão": "Gráfico de colunas: Comparação de mandados cumpridos em 2025 e 2024.",
    "3. Tráfico de Drogas": "Gráfico de barras: Comparação entre registros e presos por tráfico.",
    "3.4 Total de Material Apreendido em Teófilo Otoni": "Gráfico de pizza: Distribuição do tipo de material apreendido."
}

# Salvar sugestão de gráficos em um DataFrame
chart_df = pd.DataFrame(chart_suggestions.items(), columns=["Tópico", "Sugestão de Gráfico"])

# Exportar para o mesmo arquivo Excel em uma nova aba
with pd.ExcelWriter(file_path, engine="openpyxl", mode="a") as writer:
    chart_df.to_excel(writer, sheet_name="Sugestões de Gráficos", index=False)

file_path
