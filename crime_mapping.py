import geopandas as gpd
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import folium
import contextily as ctx
from sklearn.neighbors import KernelDensity
import numpy as np
from scipy.stats import norm
import os

# --- Configurações dos Caminhos ---
geojson_path = r"E:\QGis\UrbanoRural\Rural_19BPM.json"
excel_path = r"E:\QGis\UrbanoRural\BD_2020-2024_memorando_018-15RPM.xlsx"
output_dir = r"E:\QGis\UrbanoRural\Mapas"
os.makedirs(output_dir, exist_ok=True)

# Média fornecida
MEAN_THRESHOLD = 1428.76

# --- Funções Auxiliares ---


def load_data():
    """Carrega os dados do Excel e GeoJSON."""
    try:
        df = pd.read_excel(excel_path, sheet_name="BD_tratado")
        required_cols = ['MUNICIPIO', 'URB_RURAL', 'LONGITUDE', 'LATITUDE',
                         'IMV_TOTAL', 'ICVPE_TOTAL', 'ICVPA_TOTAL', 'ANO_FATO']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(
                f"Colunas obrigatórias não encontradas no Excel: {missing_cols}")

        gdf = gpd.read_file(geojson_path)
        if gdf.crs != "EPSG:4326":
            gdf = gdf.to_crs("EPSG:4326")
        return df, gdf
    except Exception as e:
        print(f"Erro ao carregar dados: {str(e)}")
        raise


def calculate_percentages(df, column='counts'):
    """Calcula percentuais por município e tipo (urbano/rural)."""
    try:
        total = df[column].sum()
        crime_counts = df.groupby(['MUNICIPIO', 'URB_RURAL'])[
            column].sum().reset_index(name='counts')
        crime_counts['percent'] = (
            crime_counts['counts'] / total * 100).round(2)
        return crime_counts
    except Exception as e:
        print(f"Erro ao calcular percentuais: {str(e)}")
        raise


def statistical_analysis(df, column='counts'):
    """Realiza análise estatística básica e teste de proporções."""
    try:
        urban_crimes = df[df['URB_RURAL'] == 'URBANO'][column].sum()
        rural_crimes = df[df['URB_RURAL'] == 'RURAL'][column].sum()
        total = urban_crimes + rural_crimes

        counts_by_municipio = df.groupby('MUNICIPIO')[column].sum()
        stats_summary = {
            'Média': counts_by_municipio.mean(),
            'Mediana': counts_by_municipio.median(),
            'Desvio Padrão': counts_by_municipio.std(),
            'Total Urbano': urban_crimes,
            'Total Rural': rural_crimes
        }

        if total > 0:
            prop_urban = urban_crimes / total
            prop_rural = rural_crimes / total
            pooled_prop = (urban_crimes + rural_crimes) / (total + total)
            se = np.sqrt(pooled_prop * (1 - pooled_prop)
                         * (1 / total + 1 / total))
            if se > 0:
                z_stat = (prop_urban - prop_rural) / se
                p_value = 2 * (1 - norm.cdf(abs(z_stat)))
                stats_summary['Z-test p-value'] = p_value
            else:
                stats_summary['Z-test p-value'] = None
        else:
            stats_summary['Z-test p-value'] = None

        return stats_summary
    except Exception as e:
        print(f"Erro na análise estatística: {str(e)}")
        raise


def categorize_by_mean(counts):
    """Categoriza percentuais em faixas com base na média."""
    conditions = [
        (counts <= MEAN_THRESHOLD / 3),
        (counts > MEAN_THRESHOLD / 3) & (counts <= 2 * MEAN_THRESHOLD / 3),
        (counts > 2 * MEAN_THRESHOLD / 3)
    ]
    choices = ['low', 'medium', 'high']
    return pd.cut(counts, bins=[0, MEAN_THRESHOLD / 3, 2 * MEAN_THRESHOLD / 3, np.inf], labels=choices, include_lowest=True)


def plot_choropleth(gdf, crime_counts, output_path, title_prefix):
    """Gera mapa cloroplético com faixas de cores baseadas na média."""
    try:
        urban_data = crime_counts[crime_counts['URB_RURAL']
                                  == 'URBANO'][['MUNICIPIO', 'percent']]
        rural_data = crime_counts[crime_counts['URB_RURAL']
                                  == 'RURAL'][['MUNICIPIO', 'percent']]

        gdf = gdf.merge(urban_data, left_on='NM_MUN_2',
                        right_on='MUNICIPIO', how='left')
        gdf = gdf.merge(rural_data, left_on='NM_MUN_2', right_on='MUNICIPIO',
                        how='left', suffixes=('_urban', '_rural'))

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(20, 10))

        gdf['category_urban'] = categorize_by_mean(
            gdf['percent_urban'].fillna(0))
        gdf['category_rural'] = categorize_by_mean(
            gdf['percent_rural'].fillna(0))

        gdf.plot(column='category_urban', cmap='RdYlGn', legend=True, ax=ax1,
                 categorical=True, legend_kwds={'loc': 'center left', 'bbox_to_anchor': (1, 0.5)},
                 missing_kwds={'color': 'lightgrey'}, edgecolor='black')
        ax1.set_title(f'{title_prefix} - Percentual de Crimes - Urbano')
        ctx.add_basemap(ax1, crs=gdf.crs.to_string(),
                        source=ctx.providers.OpenStreetMap.Mapnik)

        gdf.plot(column='category_rural', cmap='RdYlGn', legend=True, ax=ax2,
                 categorical=True, legend_kwds={'loc': 'center left', 'bbox_to_anchor': (1, 0.5)},
                 missing_kwds={'color': 'lightgrey'}, edgecolor='black')
        ax2.set_title(f'{title_prefix} - Percentual de Crimes - Rural')
        ctx.add_basemap(ax2, crs=gdf.crs.to_string(),
                        source=ctx.providers.OpenStreetMap.Mapnik)

        plt.tight_layout()
        plt.savefig(output_path, dpi=300, bbox_inches='tight')
        plt.close()
    except Exception as e:
        print(f"Erro ao gerar mapa cloroplético: {str(e)}")
        raise


def plot_kernel_density(df, gdf, output_path, title_prefix):
    """Gera mapa de densidade de kernel."""
    try:
        # Calcular coordenadas médias por município
        coords = df.groupby('MUNICIPIO')[
            ['LONGITUDE', 'LATITUDE']].mean().reset_index()
        coords_gdf = gpd.GeoDataFrame(coords, geometry=gpd.points_from_xy(
            coords['LONGITUDE'], coords['LATITUDE']), crs="EPSG:4326")

        kde = KernelDensity(
            bandwidth=0.01, metric='haversine', kernel='gaussian')
        kde.fit(np.radians(coords_gdf[['LONGITUDE', 'LATITUDE']].values))

        xmin, ymin, xmax, ymax = gdf.total_bounds
        x, y = np.mgrid[xmin:xmax:100j, ymin:ymax:100j]
        xy_sample = np.vstack([x.ravel(), y.ravel()]).T
        z = np.exp(kde.score_samples(np.radians(xy_sample)))
        z = z.reshape(x.shape)

        fig, ax = plt.subplots(figsize=(12, 10))
        gdf.plot(ax=ax, color='lightgrey', edgecolor='black', alpha=0.5)
        plt.imshow(z, extent=[xmin, xmax, ymin, ymax],
                   origin='lower', cmap='hot', alpha=0.6)
        plt.colorbar(label='Densidade de Crimes')
        plt.title(f'{title_prefix} - Mapa de Densidade de Kernel - Crimes')
        ctx.add_basemap(ax, crs=gdf.crs.to_string(),
                        source=ctx.providers.OpenStreetMap.Mapnik)

        plt.savefig(output_path, dpi=300, bbox_inches='tight')
        plt.close()
    except Exception as e:
        print(f"Erro ao gerar mapa de densidade de kernel: {str(e)}")
        raise


def plot_heatmap_folium(df, output_path, title_prefix):
    """Gera mapa de calor com fundo do Google Maps."""
    try:
        # Calcular coordenadas médias por município
        coords = df.groupby('MUNICIPIO')[
            ['LONGITUDE', 'LATITUDE']].mean().reset_index()

        m = folium.Map(location=[-19.9, -43.9],
                       zoom_start=7, tiles='cartodbpositron')

        google_tiles = 'https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}'
        folium.TileLayer(tiles=google_tiles, attr='Google',
                         name='Google Maps').add_to(m)

        heat_data = [[row['LATITUDE'], row['LONGITUDE']]
                     for _, row in coords.iterrows()]
        from folium.plugins import HeatMap
        HeatMap(heat_data).add_to(m)

        folium.LayerControl().add_to(m)
        m.save(output_path)
    except Exception as e:
        print(f"Erro ao gerar mapa de calor: {str(e)}")
        raise


def plot_google_map_with_popups(df, gdf, output_path, title_prefix):
    """Gera mapa com pontos e balões de dados por município."""
    try:
        m = folium.Map(location=[-19.9, -43.9],
                       zoom_start=7, tiles='cartodbpositron')

        google_tiles = 'https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}'
        folium.TileLayer(tiles=google_tiles, attr='Google',
                         name='Google Maps').add_to(m)

        # Agregar dados por município
        crime_counts = df.groupby(['MUNICIPIO', 'URB_RURAL'])[
            df.columns[0]].sum().reset_index(name='counts')
        urban_counts = crime_counts[crime_counts['URB_RURAL'] == 'URBANO'].set_index(
            'MUNICIPIO')['counts'].reindex(gdf['NM_MUN_2'], fill_value=0)
        rural_counts = crime_counts[crime_counts['URB_RURAL'] == 'RURAL'].set_index(
            'MUNICIPIO')['counts'].reindex(gdf['NM_MUN_2'], fill_value=0)

        # Adicionar polígonos com popups
        for idx, row in gdf.iterrows():
            total = urban_counts[row['NM_MUN_2']] + \
                rural_counts[row['NM_MUN_2']]
            popup_text = f"Município: {row['NM_MUN_2']}<br>Crimes Urbano: {urban_counts[row['NM_MUN_2']]}<br>Crimes Rural: {rural_counts[row['NM_MUN_2']]}<br>Total: {total}"
            folium.GeoJson(row['geometry'], popup=folium.Popup(
                popup_text, max_width=300)).add_to(m)

        folium.LayerControl().add_to(m)
        m.save(output_path)
    except Exception as e:
        print(f"Erro ao gerar mapa com popups: {str(e)}")
        raise


def plot_by_category(df, gdf, category, output_dir):
    """Gera análises e mapas para uma categoria específica."""
    print(f"\nProcessando {category}...")
    # Calcular coordenadas médias por município
    coords = df.groupby('MUNICIPIO')[
        ['LONGITUDE', 'LATITUDE']].mean().reset_index()
    df_category = df.groupby(['ANO_FATO', 'MUNICIPIO', 'URB_RURAL'])[
        category].sum().reset_index(name=category)
    df_category = df_category.merge(coords, on='MUNICIPIO', how='left')
    crime_counts = calculate_percentages(df_category, category)
    stats_summary = statistical_analysis(df_category, category)
    print(f"Resumo Estatístico ({category}):")
    for key, value in stats_summary.items():
        print(f"{key}: {value}")

    plot_choropleth(gdf, crime_counts, os.path.join(
        output_dir, f"{category}_choropleth_map.png"), category)
    print(f"✔ Mapa cloroplético para {category} gerado.")

    plot_kernel_density(df_category, gdf, os.path.join(
        output_dir, f"{category}_kernel_density_map.png"), category)
    print(f"✔ Mapa de densidade de kernel para {category} gerado.")

    plot_heatmap_folium(df_category, os.path.join(
        output_dir, f"{category}_heatmap.html"), category)
    print(f"✔ Mapa de calor para {category} gerado.")

    plot_google_map_with_popups(df_category, gdf, os.path.join(
        output_dir, f"{category}_map_with_popups.html"), category)
    print(f"✔ Mapa com popups para {category} gerado.")

# --- Processamento Principal ---


def main():
    print("Iniciando geração de mapas...")
    try:
        df, gdf = load_data()
        print(f"Colunas disponíveis no Excel: {df.columns.tolist()}")

        # Processar categorias
        categories = ['IMV_TOTAL', 'ICVPE_TOTAL', 'ICVPA_TOTAL']
        for category in categories:
            plot_by_category(df, gdf, category, output_dir)

        # Processar total de registros
        df_total = df.groupby(
            ['ANO_FATO', 'MUNICIPIO', 'URB_RURAL']).size().reset_index(name='counts')
        coords = df.groupby('MUNICIPIO')[
            ['LONGITUDE', 'LATITUDE']].mean().reset_index()
        df_total = df_total.merge(coords, on='MUNICIPIO', how='left')
        crime_counts = calculate_percentages(df_total, 'counts')
        stats_summary = statistical_analysis(df_total, 'counts')
        print("\nResumo Estatístico (Total de Registros):")
        for key, value in stats_summary.items():
            print(f"{key}: {value}")

        plot_choropleth(gdf, crime_counts, os.path.join(
            output_dir, "total_choropleth_map.png"), "Total de Registros")
        print("✔ Mapa cloroplético para total de registros gerado.")

        plot_kernel_density(df_total, gdf, os.path.join(
            output_dir, "total_kernel_density_map.png"), "Total de Registros")
        print("✔ Mapa de densidade de kernel para total de registros gerado.")

        plot_heatmap_folium(df_total, os.path.join(
            output_dir, "total_heatmap.html"), "Total de Registros")
        print("✔ Mapa de calor para total de registros gerado.")

        plot_google_map_with_popups(df_total, gdf, os.path.join(
            output_dir, "total_map_with_popups.html"), "Total de Registros")
        print("✔ Mapa com popups para total de registros gerado.")

        print(f"\n✅ Mapas salvos em: {output_dir}")
    except Exception as e:
        print(f"\n❌ Erro durante a execução: {str(e)}")
        print("Possíveis causas:")
        print("- Caminhos dos arquivos incorretos")
        print("- Aba do Excel não encontrada")
        print("- Colunas obrigatórias faltando nos dados de entrada")
        print("- Problemas com o formato do GeoJSON")


if __name__ == "__main__":
    main()
