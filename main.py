import geopandas as gpd
import pandas as pd
from shapely.geometry import Point
from openpyxl import load_workbook

# --- Configurações dos Caminhos ---
geojson_path = r"E:\QGis\csv\Mapas_Tratados\SubSetores_19BPM_GeoJSON.json"
output_excel_path = r"E:\Painel_PowerBI\BASE_GDO\BD_2025\Query.xlsx"

# Arquivos CSV e os nomes das abas correspondentes
csv_files = {
    "BD_IMV": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\GDO_2025_1.csv",
    "BD_ICVPe": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\GDO_2025_2.csv",
    "BD_ICVPa": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\GDO_2025_3.csv",
    "BD_POG": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\GDO_2025_4.csv"
}

# --- Função para Processar Cada Arquivo ---
def process_csv(file_path, geojson_path):
    # Carregar o CSV
    df = pd.read_csv(file_path)

    # Criar GeoDataFrame
    df['geometry'] = df.apply(lambda row: Point(row['numero_longitude'], row['numero_latitude']), axis=1)
    points_gdf = gpd.GeoDataFrame(df, geometry='geometry', crs='EPSG:4326')

    # Carregar GeoJSON
    polygons_gdf = gpd.read_file(geojson_path)

    # Garantir CRS Compatível
    if polygons_gdf.crs != 'EPSG:4326':
        polygons_gdf = polygons_gdf.to_crs('EPSG:4326')

    # Realizar Interseção Geoespacial
    result_gdf = gpd.sjoin(points_gdf, polygons_gdf, how='left', predicate='within')

    # Selecionar Colunas Desejadas
    output_columns = list(df.columns) + ['name', 'PELOTAO', 'CIA_PM']
    result = result_gdf[output_columns]

    # Renomear Colunas
    result = result.rename(columns={
        'name': 'poligono_intersectado',
        'PELOTAO': 'geojson_pelotao',
        'CIA_PM': 'geojson_cia_pm'
    })

    return result

# --- Atualizar ou Criar o Arquivo Excel ---
try:
    # Se o arquivo já existir, carregar as abas existentes
    with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        for sheet_name, file_path in csv_files.items():
            # Processar o arquivo CSV
            processed_data = process_csv(file_path, geojson_path)
            
            # Salvar ou substituir a aba no Excel
            processed_data.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"Arquivo atualizado em: {output_excel_path}")
except FileNotFoundError:
    # Se o arquivo não existir, criar um novo
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        for sheet_name, file_path in csv_files.items():
            # Processar o arquivo CSV
            processed_data = process_csv(file_path, geojson_path)
            
            # Salvar no Excel
            processed_data.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"Novo arquivo criado em: {output_excel_path}")
