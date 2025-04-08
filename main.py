import geopandas as gpd
import pandas as pd
from shapely.geometry import Point
from openpyxl import load_workbook

# --- Configurações dos Caminhos ---
geojson_path = r"E:\QGis\csv\Mapas_Tratados\SubSetores_19BPM_GeoJSON.json"
# "E:\Painel_PowerBI\BASE_GDO\BD_2025\Query.xlsx"
output_excel_path = r"E:\GDO\GDO_2025\Monitoramento_GDO_2025.xlsx"


# Arquivos CSV e os nomes das abas correspondentes
csv_files = {
    "BD_IMV2025": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\GDO_2025_1.csv",
    "BD_ICVPe": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\GDO_2025_2.csv",
    "BD_ICVPa": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\GDO_2025_3.csv",
    "BD_POG": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\GDO_2025_4.csv",
    "BD_PL": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\GDO_2025_5.csv",
    "BD_IRTD": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\GDO_2025_6.csv",
    "BD_INTERACAO_COMUNITARIA": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\GDO_2025_7.csv",
    "BD_MRPP_INT_CUM": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\INT_COM_2025_1.csv",
    "BD_RC_INT_CUM": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\INT_COM_2025_2.csv",
    "BD_VCP_INT_CUM": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\INT_COM_2025_3.csv",
    "BD_VTC_DENOMINADOR_INT_CUM": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\INT_COM_2025_4.csv",
    "BD_VTC_NUMERADOR_INT_CUM": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\INT_COM_2025_5.csv",
    "BD_VT_DENOMINADOR_INT_CUM": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\INT_COM_2025_6.csv",
    "BD_VT_NUMERADOR_INT_CUM": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\INT_COM_2025_7.csv"
}

# --- Função para Processar Cada Arquivo ---


def process_csv(file_path, geojson_path):
    try:
        # Carregar o CSV
        df = pd.read_csv(file_path)
        
        # Verificar se as colunas necessárias existem
        required_columns = ['numero_longitude', 'numero_latitude']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"\n[ERRO] Arquivo CSV: {file_path}")
            print(f"Colunas esperadas: {required_columns}")
            print(f"Colunas faltando: {missing_columns}")
            print(f"Colunas disponíveis: {df.columns.tolist()}\n")
            raise ValueError(f"Colunas faltando: {missing_columns}")
        
        # Criar GeoDataFrame
        df['geometry'] = df.apply(lambda row: Point(
            row['numero_longitude'], row['numero_latitude']), axis=1)
        points_gdf = gpd.GeoDataFrame(df, geometry='geometry', crs='EPSG:4326')

        # Restante do seu código...
        polygons_gdf = gpd.read_file(geojson_path)

        if polygons_gdf.crs != 'EPSG:4326':
            polygons_gdf = polygons_gdf.to_crs('EPSG:4326')

        result_gdf = gpd.sjoin(points_gdf, polygons_gdf,
                               how='left', predicate='within')

        output_columns = list(df.columns) + ['name', 'PELOTAO', 'CIA_PM']
        result = result_gdf[output_columns]

        result = result.rename(columns={
            'name': 'SubSetor',
            'PELOTAO': 'Pelotao',
            'CIA_PM': 'CIA_PM'
        })

        if 'data_hora_fato' in result.columns:
            result['data_fato'] = pd.to_datetime(
                result['data_hora_fato'], errors='coerce').dt.date

        return result

    except Exception as e:
        print(f"\n[ERRO] Falha ao processar o arquivo: {file_path}")
        print(f"Detalhes do erro: {str(e)}\n")
        raise  # Re-lança o erro para interromper a execução


# --- Atualizar ou Criar o Arquivo Excel ---
try:
    # Se o arquivo já existir, carregar as abas existentes
    with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        for sheet_name, file_path in csv_files.items():
            # Processar o arquivo CSV e criar a nova coluna 'data_fato'
            processed_data = process_csv(file_path, geojson_path)

            # Salvar ou substituir a aba no Excel
            processed_data.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"Arquivo atualizado em: {output_excel_path}")
except FileNotFoundError:
    # Se o arquivo não existir, criar um novo
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        for sheet_name, file_path in csv_files.items():
            # Processar o arquivo CSV e criar a nova coluna 'data_fato'
            processed_data = process_csv(file_path, geojson_path)

            # Salvar no Excel
            processed_data.to_excel(writer, sheet_name=sheet_name, index=False)
            print("Caminho absoluto do arquivo salvo:",
                  os.path.abspath(output_excel_path))
    print(f"Novo arquivo criado em: {output_excel_path}")
