# 'dataset' tem os dados de entrada para este script
import geopandas as gpd
import pandas as pd
from shapely.geometry import Point

# --- Configurações do Caminho ---
# Caminho para o arquivo GeoJSON
geojson_path = r"E:\QGis\UrbanoRural\Rural_19BPM.json"

# --- Carregar Dados do Power BI ---
# A tabela 'f_CVPe' será carregada automaticamente como 'dataset' no Power BI
df = dataset

# --- Criar GeoDataFrame para a Tabela 'f_CVPe' ---
# Criar a coluna de geometria com pontos baseados em latitude e longitude
df['geometry'] = df.apply(lambda row: Point(row['numero_longitude'], row['numero_latitude']), axis=1)
points_gdf = gpd.GeoDataFrame(df, geometry='geometry', crs="EPSG:4326")  # Define EPSG:4326 diretamente

# --- Carregar o GeoJSON ---
polygons_gdf = gpd.read_file(geojson_path)

# Garantir que o GeoJSON está no CRS EPSG:4326
if polygons_gdf.crs != "EPSG:4326":
    polygons_gdf = polygons_gdf.to_crs("EPSG:4326")

# --- Realizar Interseção Geoespacial ---
# Usar um 'spatial join' para encontrar qual ponto está dentro de qual polígono
result_gdf = gpd.sjoin(points_gdf, polygons_gdf, how="left", predicate="within")

# --- Preparar Resultado ---
# Selecionar todas as colunas originais de `f_CVPe` mais as colunas específicas do GeoJSON
output_columns = list(df.columns) + ['Tipo']
output = result_gdf[output_columns]

# Renomear as colunas do GeoJSON para deixar mais claro no contexto do Power BI
output = output.rename(columns={
    'name': 'Urbano_Rural'
})

# Ajustar o formato dos números para usar vírgulas como separadores decimais
output['numero_latitude'] = output['numero_latitude'].apply(lambda x: f"{x:.6f}".replace('.', ','))
output['numero_longitude'] = output['numero_longitude'].apply(lambda x: f"{x:.6f}".replace('.', ','))


# Converter para DataFrame para exportação ao Power BI
output = pd.DataFrame(output)
