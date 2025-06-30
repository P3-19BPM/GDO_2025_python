import pandas as pd
import geopandas as gpd
import pyodbc
import os
from shapely.geometry import Point

# 🔹 Configuração de conexão ODBC
dsn_name = 'Sample Cloudera Impala DSN'
username = 'SEUCPF'
password = 'SENHABISP'

# 🔹 Caminhos
geojson_path = r"C:\GDO 2025 - 21 CIA PM IND\ZonaRural.json"
output_folder = r"C:\GDO 2025 - 21 CIA PM IND\export_zonarural"
output_file = os.path.join(output_folder, "dados_ocorrencias_zonarural.csv")

# 🔹 Garante existência da pasta de saída
os.makedirs(output_folder, exist_ok=True)
print(f"📂 Diretório de exportação: {output_folder}")

# 🔹 Conectar ao banco de dados
connection_string = f'DSN={dsn_name};UID={username};PWD={password};'
conn = pyodbc.connect(connection_string, autocommit=True)
cursor = conn.cursor()

try:
    print("✅ Conectado ao banco de dados.")

    # 🔹 Consulta SQL
    query = """
    SELECT
        oco.numero_ocorrencia,
        oco.numero_latitude,
        oco.numero_longitude
    FROM db_bisp_reds_reporting.tb_ocorrencia AS oco
    WHERE
        YEAR(oco.data_hora_fato) >= 2025
        AND codigo_municipio IN (310040,310250,310570,312820,313550,314585,315020,315210,315490,315500,315740,316010,316400,317050)
        AND oco.descricao_estado NOT IN ('INUTILIZADO', 'NOVO PENDENTE DE ELABORACAO')
    """
    print("⏳ Executando consulta SQL...")
    cursor.execute(query)

    # 🔹 DataFrame com resultados
    columns = [desc[0] for desc in cursor.description]
    rows = cursor.fetchall()
    df = pd.DataFrame.from_records(rows, columns=columns)
    print(f"✅ Consulta finalizada: {df.shape[0]} registros obtidos.")

    # 🔹 Remover registros com coordenadas ausentes
    df.dropna(subset=['numero_latitude', 'numero_longitude'], inplace=True)

    # 🔹 GeoDataFrame dos pontos
    df['geometry'] = df.apply(lambda row: Point(
        row['numero_longitude'], row['numero_latitude']), axis=1)
    points_gdf = gpd.GeoDataFrame(df, geometry='geometry', crs="EPSG:4326")

    # 🔹 Carregar e alinhar o GeoJSON
    polygons_gdf = gpd.read_file(geojson_path)
    if polygons_gdf.crs != "EPSG:4326":
        polygons_gdf = polygons_gdf.to_crs("EPSG:4326")

    # 🔹 Junção espacial
    print("📍 Realizando junção espacial...")
    result_gdf = gpd.sjoin(points_gdf, polygons_gdf,
                           how="left", predicate="within")

    # 🔹 Limpar colunas desnecessárias (como visuais do GeoJSON)
    colunas_a_remover = ['index_right',
        'fill-opacity', 'stroke-opacity', 'stroke']
    result_gdf.drop(columns=[
                    col for col in colunas_a_remover if col in result_gdf.columns], inplace=True)

    # 🔹 Exportar CSV
    result_gdf.to_csv(output_file, index=False, encoding='utf-8-sig')
    print(f"✅ Arquivo CSV salvo com sucesso: {output_file}")

except Exception as e:
    print(f"❌ Erro durante a execução: {e}")

finally:
    cursor.close()
    conn.close()
    print("🔌 Conexão encerrada."
