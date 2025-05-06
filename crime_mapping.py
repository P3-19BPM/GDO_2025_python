import geopandas as gpd
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import folium
import contextily as ctx
from sklearn.neighbors import KernelDensity
import numpy as np
from scipy import stats
import uuid
import os

# --- Configurações dos Caminhos ---
geojson_path = r"E:\QGis\UrbanoRural\Uniao_Urb_Rural_MG.json"
excel_path = r"E:\QGis\UrbanoRural\BD_2020-2024_memorando_018-15RPM.xlsx"
output_dir = r"E:\QGis\UrbanoRural\Mapas"
os.makedirs(output_dir, exist_ok=True)

# --- Funções Auxiliares ---
def load_data():
    """Carrega os dados do Excel e GeoJSON."""
    df = pd.read_excel(excel_path, sheet_name="bd_2020-2024_memorando_018")
    gdf = gpd.read_file(geojson_path)
    if gdf.crs != "EPSG:4326":
        gdf = gdf.to_crs("EPSG:4326")
    return df, gdf

def calculate_percentages(df):
    """Calcula percentuais de crimes por município e tipo (urbano/rural)."""
    total_crimes = len(df)
    crime_counts = df.groupby(['MUNICIPIO', 'URB_RURAL']).size().reset_index(name='counts')
    crime_counts['percent'] = (crime_counts['counts'] / total_crimes * 100).round(2)
    return crime_counts

def statistical_analysis(df):
    """Realiza análise estatística básica e teste de proporções."""
    urban_crimes = df[df['URB_RURAL'] == 'URBANO'].shape[0]
    rural_crimes = df[df['URB_RURAL'] == 'RURAL'].shape[0]
    total = urban_crimes + rural_crimes
    
    # Estatísticas descritivas
    counts_by_municipio = df.groupby('MUNICIPIO').size()
    stats_summary = {
        'Média': counts_by_municipio.mean(),
        'Mediana': counts_by_municipio.median(),
        'Desvio Padrão': counts_by_municipio.std(),
        'Total Urbano': urban_crimes,
        'Total Rural': rural_crimes
    }
    
    # Teste de proporções
    prop_urban = urban_crimes / total
    prop_rural = rural_crimes / total
    z_stat, p_value = stats.proportions_ztest(
        [urban_crimes, rural_crimes], [total, total]
    )
    stats_summary['Z-test p-value'] = p_value
    
    return stats_summary

def plot_choropleth(gdf, crime_counts, output_path):
    """Gera mapa cloroplético com percentuais por município."""
    urban_data = crime_counts[crime_counts['URB_RURAL'] == 'URBANO'][['MUNICIPIO', 'percent']]
    rural_data = crime_counts[crime_counts['URB_RURAL'] == 'RURAL'][['MUNICIPIO', 'percent']]
    
    gdf = gdf.merge(urban_data, left_on='NM_MUN_2', right_on='MUNICIPIO', how='left')
    gdf = gdf.merge(rural_data, left_on='NM_MUN_2', right_on='MUNICIPIO', how='left', suffixes=('_urban', '_rural'))
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(20, 10))
    
    gdf.plot(column='percent_urban', cmap='Reds', legend=True, ax=ax1,
             missing_kwds={'color': 'lightgrey'}, edgecolor='black')
    ax1.set_title('Percentual de Crimes - Urbano')
    ctx.add_basemap(ax1, crs=gdf.crs.to_string(), source=ctx.providers.OpenStreetMap.Mapnik)
    
    gdf.plot(column='percent_rural', cmap='Blues', legend=True, ax=ax2,
             missing_kwds={'color': 'lightgrey'}, edgecolor='black')
    ax2.set_title('Percentual de Crimes - Rural')
    ctx.add_basemap(ax2, crs=gdf.crs.to_string(), source=ctx.providers.OpenStreetMap.Mapnik)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()

def plot_kernel_density(df, gdf, output_path):
    """Gera mapa de densidade de kernel."""
    coords = df[['LONGITUDE', 'LATITUDE']].values
    kde = KernelDensity(bandwidth=0.01, metric='haversine', kernel='gaussian')
    kde.fit(np.radians(coords))
    
    xmin, ymin, xmax, ymax = gdf.total_bounds
    x, y = np.mgrid[xmin:xmax:100j, ymin:ymax:100j]
    xy_sample = np.vstack([x.ravel(), y.ravel()]).T
    z = np.exp(kde.score_samples(np.radians(xy_sample)))
    z = z.reshape(x.shape)
    
    fig, ax = plt.subplots(figsize=(12, 10))
    gdf.plot(ax=ax, color='lightgrey', edgecolor='black', alpha=0.5)
    plt.imshow(z, extent=[xmin, xmax, ymin, ymax], origin='lower', cmap='hot', alpha=0.6)
    plt.colorbar(label='Densidade de Crimes')
    plt.title('Mapa de Densidade de Kernel - Crimes')
    ctx.add_basemap(ax, crs=gdf.crs.to_string(), source=ctx.providers.OpenStreetMap.Mapnik)
    
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()

def plot_heatmap_folium(df, output_path):
    """Gera mapa de calor com fundo do Google Maps."""
    m = folium.Map(location=[-19.9, -43.9], zoom_start=7, tiles='cartodbpositron')
    
    # Adicionar fundo do Google Maps
    google_tiles = 'https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}'
    folium.TileLayer(tiles=google_tiles, attr='Google', name='Google Maps').add_to(m)
    
    # Dados para o mapa de calor
    heat_data = [[row['LATITUDE'], row['LONGITUDE']] for _, row in df.iterrows()]
    from folium.plugins import HeatMap
    HeatMap(heat_data).add_to(m)
    
    folium.LayerControl().add_to(m)
    m.save(output_path)

def plot_google_map(df, output_path):
    """Gera mapa com pontos categorizados e fundo do Google Maps."""
    m = folium.Map(location=[-19.9, -43.9], zoom_start=7, tiles='cartodbpositron')
    
    google_tiles = 'https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}'
    folium.TileLayer(tiles=google_tiles, attr='Google', name='Google Maps').add_to(m)
    
    # Adicionar pontos
    for _, row in df.iterrows():
        color = 'red' if row['URB_RURAL'] == 'URBANO' else 'blue'
        folium.CircleMarker(
            location=[row['LATITUDE'], row['LONGITUDE']],
            radius=5,
            color=color,
            fill=True,
            fill_color=color,
            fill_opacity=0.7,
            popup=f"{row['MUNICIPIO']} - {row['URB_RURAL']}"
        ).add_to(m)
    
    folium.LayerControl().add_to(m)
    m.save(output_path)

# --- Processamento Principal ---
def main():
    print("Iniciando geração de mapas...")
    
    # Carregar dados
    df, gdf = load_data()
    
    # Calcular percentuais
    crime_counts = calculate_percentages(df)
    
    # Análise estatística
    stats_summary = statistical_analysis(df)
    print("\nResumo Estatístico:")
    for key, value in stats_summary.items():
        print(f"{key}: {value}")
    
    # Gerar mapas
    plot_choropleth(gdf, crime_counts, os.path.join(output_dir, "choropleth_map.png"))
    print("✔ Mapa cloroplético gerado.")
    
    plot_kernel_density(df, gdf, os.path.join(output_dir, "kernel_density_map.png"))
    print("✔ Mapa de densidade de kernel gerado.")
    
    plot_heatmap_folium(df, os.path.join(output_dir, "heatmap.html"))
    print("✔ Mapa de calor gerado.")
    
    plot_google_map(df, os.path.join(output_dir, "google_map_points.html"))
    print("✔ Mapa com fundo Google gerado.")
    
    print(f"\n✅ Mapas salvos em: {output_dir}")

if __name__ == "__main__":
    main()