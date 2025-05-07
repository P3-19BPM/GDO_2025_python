import geopandas as gpd
import pandas as pd
from shapely.geometry import Point
import os

# --- Configurações dos Caminhos ---
geojson_path = r"E:\QGis\UrbanoRural\Rural_19BPM.json"
output_excel_path = r"E:\QGis\UrbanoRural\BD_2020-2024_memorando_018-15RPM.xlsx"

# Dicionário com os arquivos e abas correspondentes
excel_sheets = {
    "BD_tratado": {
        "file_path": output_excel_path,
        "sheet_name": "bd_2020-2024_memorando_018"
    }
}

def process_data(input_path, sheet_name, geojson_path):
    try:
        # 1. Carregar os dados do Excel
        print(f"Carregando dados do Excel: {input_path} - Aba: {sheet_name}")
        df = pd.read_excel(input_path, sheet_name=sheet_name)
        
        # Verificar colunas obrigatórias
        required_cols = ['LONGITUDE', 'LATITUDE']
        if not all(col in df.columns for col in required_cols):
            missing = [col for col in required_cols if col not in df.columns]
            raise ValueError(f"Colunas obrigatórias não encontradas: {missing}")

        # 2. Criar GeoDataFrame com os pontos
        print("Criando pontos geográficos...")
        df['geometry'] = df.apply(lambda row: Point(row['LONGITUDE'], row['LATITUDE']), axis=1)
        points_gdf = gpd.GeoDataFrame(df, geometry='geometry', crs='EPSG:4326')

        # 3. Carregar o GeoJSON
        print(f"Carregando GeoJSON: {geojson_path}")
        polygons_gdf = gpd.read_file(geojson_path)
        
        # Verificar estrutura do GeoJSON
        print("\nInformações do GeoJSON:")
        print(f"Colunas disponíveis: {polygons_gdf.columns.tolist()}")
        print(f"Valores únicos em 'Tipo': {polygons_gdf['Tipo'].unique()}")
        print(f"Número de polígonos: {len(polygons_gdf)}")
        
        # Garantir CRS compatível
        if polygons_gdf.crs != 'EPSG:4326':
            polygons_gdf = polygons_gdf.to_crs('EPSG:4326')

        # 4. Realizar a junção espacial
        print("Realizando junção espacial...")
        result_gdf = gpd.sjoin(points_gdf, polygons_gdf, how='left', predicate='within')
        
        # 5. Padronizar a coluna URB_RURAL
        print("Processando classificação urbano/rural...")
        if 'Tipo' in result_gdf.columns:
            # Simplificar a classificação
            result_gdf['URB_RURAL'] = result_gdf['Tipo'].str.upper().apply(
                lambda x: 'URBANO' if 'URBAN' in str(x) else 'RURAL')
        else:
            available_cols = result_gdf.columns.tolist()
            raise ValueError(f"Coluna 'Tipo' não encontrada. Colunas disponíveis: {available_cols}")

        # 6. Selecionar colunas para o resultado final
        print("Preparando resultado final...")
        original_cols = [col for col in df.columns if col != 'geometry']
        result = result_gdf[original_cols + ['URB_RURAL', 'NM_MUN_2', 'Area_KM2']]
        
        # Renomear colunas para padronização
        result = result.rename(columns={
            'NM_MUN_2': 'MUNICIPIO',
            'Area_KM2': 'AREA_MUN_KM2'
        })

        # 7. Processar datas se existirem
        if 'data_hora_fato' in result.columns:
            result['data_fato'] = pd.to_datetime(result['data_hora_fato'], errors='coerce').dt.date

        return result

    except Exception as e:
        print(f"Erro durante o processamento: {str(e)}")
        raise

# --- Processamento principal ---
try:
    print("Iniciando processamento...")
    
    with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        for sheet_name, sheet_info in excel_sheets.items():
            print(f"\nProcessando aba: {sheet_info['sheet_name']}")
            
            processed_data = process_data(
                input_path=sheet_info["file_path"],
                sheet_name=sheet_info["sheet_name"],
                geojson_path=geojson_path
            )
            
            # Salvar no Excel
            processed_data.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"✔ Dados processados salvos na aba '{sheet_name}'")
    
    print(f"\n✅ Processamento concluído com sucesso!")
    print(f"Arquivo salvo em: {output_excel_path}")

except Exception as e:
    print(f"\n❌ ERRO durante a execução: {str(e)}")
    print("Possíveis causas:")
    print("- Caminhos dos arquivos incorretos")
    print("- Colunas obrigatórias faltando nos dados de entrada")
    print("- Problemas com o formato do GeoJSON")