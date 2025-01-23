import geopandas as gpd
import pandas as pd
import os
import shutil
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def generate_files():
    # Configuração inicial para selecionar o arquivo
    Tk().withdraw()  # Esconde a janela principal do Tkinter
    input_file = askopenfilename(title="Selecione o arquivo de entrada")  # Abre o seletor de arquivo
    
    if not input_file:
        print("Nenhum arquivo selecionado.")
        return
    
    # Identificar extensão do arquivo
    file_extension = os.path.splitext(input_file)[1].lower()
    print(f"Arquivo selecionado: {input_file}")
    print(f"Extensão identificada: {file_extension}")
    
    # Caminho de saída
    download_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    base_output_dir = os.path.join(download_dir, "Mapas")
    os.makedirs(base_output_dir, exist_ok=True)
    
    # Carregar o arquivo dependendo da extensão
    try:
        if file_extension in ['.geojson', '.json', '.shp', '.kml']:
            gdf = gpd.read_file(input_file)
        elif file_extension in ['.csv']:
            df = pd.read_csv(input_file)
            gdf = gpd.GeoDataFrame(df)  # Tenta transformar em GeoDataFrame
        else:
            print(f"Extensão não suportada: {file_extension}")
            return
    except Exception as e:
        print(f"Erro ao carregar o arquivo: {e}")
        return
    
    # Lista de formatos para exportação
    formats = {
        "GeoJSON": "geojson",
        "TopoJSON": "topojson",
        "KML": "kml",
        "Shapefile": "shp",
        "JSON": "json",
        "SVG": "svg",
        "CSV": "csv",
        "Snapshot": "snapshot"  # Nome fictício; pode ser implementado como PDF ou imagem
    }

    # Criar subpastas e salvar os arquivos
    for folder_name, extension in formats.items():
        output_folder = os.path.join(base_output_dir, folder_name)
        os.makedirs(output_folder, exist_ok=True)

        try:
            # Definir o caminho do arquivo de saída
            output_file = os.path.join(output_folder, f"mapa.{extension}")

            # Salvar conforme o formato
            if extension == "geojson":
                gdf.to_file(output_file, driver="GeoJSON")
            elif extension == "topojson":
                # O GeoPandas não suporta diretamente TopoJSON; necessita de outras bibliotecas
                print(f"Conversão para {extension} não implementada diretamente.")
            elif extension == "kml":
                gdf.to_file(output_file, driver="KML")
            elif extension == "shp":
                gdf.to_file(output_file, driver="ESRI Shapefile")
            elif extension == "json":
                gdf.to_file(output_file, driver="GeoJSON")
            elif extension == "svg":
                print(f"Exportação para {extension} não implementada diretamente.")
            elif extension == "csv":
                gdf.to_csv(output_file, index=False)
            elif extension == "snapshot":
                print(f"Exportação para {extension} precisa ser definida pelo usuário.")
            
            print(f"Arquivo {folder_name} salvo com sucesso em {output_file}")
        except Exception as e:
            print(f"Erro ao salvar {folder_name}: {e}")
    
    print(f"Todos os arquivos foram gerados na pasta: {base_output_dir}")

if __name__ == "__main__":
    generate_files()
