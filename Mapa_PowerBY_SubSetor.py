# Importação das bibliotecas necessárias
import geopandas as gpd  # Biblioteca especializada em manipulação de dados geoespaciais.
import pandas as pd  # Usada para manipulação de tabelas (DataFrames).
from shapely.geometry import Point  # Usada para criar objetos geográficos, como pontos no mapa.
from datetime import datetime, timedelta  # Usada para manipulação de datas e intervalos de tempo.

# --- Configurações do Caminho ---
# Caminho do arquivo GeoJSON contendo informações geográficas de polígonos (como setores, bairros, etc.).
geojson_path = r"C:\Users\valfr\Downloads\csv\Mapas_Tratados\SubSetores_19BPM_GeoJSON.json"

# --- Carregar Dados do Power BI ---
# O dataset é um DataFrame fornecido pelo Power BI. Ele contém informações tabulares (linhas e colunas)
# como Latitude, Longitude, e outros dados relacionados. Aqui, simplesmente estamos atribuindo o dataset à variável 'df'.
df = dataset

# --- Verificar e Filtrar Linhas Inválidas ---
# Removemos quaisquer linhas do DataFrame onde os valores de Latitude ou Longitude estão ausentes (NaN).
df = df.dropna(subset=['Latitude', 'Longitude'])

# Mantemos apenas as linhas onde Latitude e Longitude têm valores válidos (não nulos).
df = df[(df['Latitude'].notnull()) & (df['Longitude'].notnull())]

# --- Criar GeoDataFrame ---
# Aqui criamos uma nova coluna chamada 'geometry', que converte cada par de Latitude e Longitude em um ponto geográfico.
if not df.empty:  # Verifica se o DataFrame não está vazio após a filtragem.
    df['geometry'] = df.apply(lambda row: Point(row['Longitude'], row['Latitude']), axis=1)

    # Converte o DataFrame do pandas em um GeoDataFrame do GeoPandas.
    # Um GeoDataFrame é como um DataFrame comum, mas com suporte para manipulação de dados geoespaciais.
    points_gdf = gpd.GeoDataFrame(df, geometry='geometry', crs="EPSG:4326")
else:
    # Caso o DataFrame esteja vazio (não há dados válidos), lançamos um erro.
    raise ValueError("O DataFrame está vazio após filtrar as linhas inválidas.")

# --- Carregar o GeoJSON ---
# Lê o arquivo GeoJSON, que contém informações geográficas (como polígonos representando áreas no mapa).
polygons_gdf = gpd.read_file(geojson_path)

# Garantir que o sistema de coordenadas do GeoJSON seja compatível com o EPSG:4326.
# EPSG:4326 é um padrão usado para coordenadas geográficas (latitude e longitude).
if polygons_gdf.crs != "EPSG:4326":
    polygons_gdf = polygons_gdf.to_crs("EPSG:4326")

# --- Realizar Interseção Geoespacial ---
# Realizamos um "Spatial Join" (junção espacial) para combinar os pontos (latitude/longitude) com os polígonos.
# Isso associa cada ponto ao polígono em que ele está contido.
result_gdf = gpd.sjoin(points_gdf, polygons_gdf, how="left", predicate="within")

# --- Garantir a Preservação da Coluna DATA ---
# Aqui asseguramos que todas as colunas originais do DataFrame, incluindo a coluna 'DATA', sejam preservadas.
# Adicionamos também colunas do GeoJSON que nos interessam, como 'name', 'PELOTAO', 'CIA_PM', e 'municipio'.
output_columns = list(df.columns) + ['name', 'PELOTAO', 'CIA_PM', 'municipio']
output = result_gdf[output_columns]

# Renomear as colunas do GeoJSON
# Para facilitar a leitura e uso dos dados no Power BI, renomeamos algumas colunas do GeoJSON.
output = output.rename(columns={
    'name': 'poligono_intersectado',  # Nome do polígono em que o ponto está contido.
    'PELOTAO': 'geojson_pelotao',  # Pelotão relacionado ao polígono.
    'CIA_PM': 'geojson_cia_pm',  # Companhia PM relacionada ao polígono.
    'municipio': 'geojson_municipio'  # Município relacionado ao polígono.
})

# --- Ajustar o formato dos números ---
# Convertemos os valores de Latitude e Longitude para usar vírgulas no lugar de pontos decimais,
# porque alguns formatos ou configurações locais (por exemplo, no Brasil) preferem vírgulas.
output['Latitude'] = output['Latitude'].apply(lambda x: f"{x:.6f}".replace('.', ','))
output['Longitude'] = output['Longitude'].apply(lambda x: f"{x:.6f}".replace('.', ','))

# --- Converter para DataFrame ---
# Finalmente, convertemos o GeoDataFrame de volta para um DataFrame do pandas, que é o formato esperado pelo Power BI.
output = pd.DataFrame(output)