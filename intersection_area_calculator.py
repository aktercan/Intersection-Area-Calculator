import geopandas as gpd
import pandas as pd
from shapely.geometry import Polygon


def load_shapefile(file_path):
    """Loads a shapefile and ensures its geometries are valid."""
    gdf = gpd.read_file(file_path)
    gdf['geometry'] = gdf['geometry'].apply(lambda geom: geom.buffer(0) if not geom.is_valid else geom)
    return gdf


def calculate_areas(gdf, area_column_name):
    """Adds an area column to a GeoDataFrame."""
    gdf[area_column_name] = gdf.geometry.area
    return gdf


def calculate_intersections(mahalle, parsel):
    """Calculates intersections and returns detailed and summary results."""
    # Initialize spatial index for performance
    sindex = parsel.sindex

    result_list = []
    summary_list = []

    for i, row_mahalle in mahalle.iterrows():
        # Find potential matches using spatial index
        possible_matches_index = list(sindex.intersection(row_mahalle.geometry.bounds))
        possible_matches = parsel.iloc[possible_matches_index]

        # Filter precise matches
        precise_matches = possible_matches[possible_matches.intersects(row_mahalle.geometry)]

        intersection_geometries = []
        for j, row_parsel in precise_matches.iterrows():
            intersection = row_mahalle.geometry.intersection(row_parsel.geometry)
            if not intersection.is_empty:
                intersected_area = intersection.area
                intersection_geometries.append(intersection)
                result_list.append({
                    'mahalle_ID': row_mahalle['ID'],
                    'mahalle_ADI': row_mahalle['ADI'],
                    'mahalle_ILCE_ID': row_mahalle['ILCE_ID'],
                    'mahalle_TIP_ID': row_mahalle['TIP_ID'],
                    'mahalle_UAVT_KODU': row_mahalle['UAVT_KODU'],
                    'mahalle_NUFUS': row_mahalle['NUFUS'],
                    'parsel_KAD_PARSEL': row_parsel['KAD_PARSEL'],
                    'parsel_MI_PRINX': row_parsel['MI_PRINX'],
                    'area_mahalle': row_mahalle['area_mahalle'],
                    'area_parsel': row_parsel['area_parsel'],
                    'intersected_area': intersected_area
                })

        # Calculate the sum of intersected areas
        if len(intersection_geometries) > 1:
            merged_intersection = intersection_geometries[0]
            for geom in intersection_geometries[1:]:
                merged_intersection = merged_intersection.union(geom)
            sum_of_intersected_area = merged_intersection.area
        else:
            sum_of_intersected_area = intersection_geometries[0].area if intersection_geometries else 0

        summary_list.append({
            'mahalle_ID': row_mahalle['ID'],
            'mahalle_ADI': row_mahalle['ADI'],
            'sum_of_intersected_area': sum_of_intersected_area
        })

    # Convert results to DataFrames
    detailed_results = pd.DataFrame(result_list)
    summary_results = pd.DataFrame(summary_list)

    return detailed_results, summary_results


def save_to_excel(detailed_results, summary_results, output_file):
    """Saves the detailed and summary results to an Excel file."""
    with pd.ExcelWriter(output_file) as writer:
        detailed_results.to_excel(writer, sheet_name='Detaylar', index=False)
        summary_results.to_excel(writer, sheet_name='Özet', index=False)
    print(f"Results saved to {output_file}")


def main(mahalle_path, parsel_path, output_file):
    """Main function to execute the workflow."""
    # Load shapefiles
    mahalle = load_shapefile(mahalle_path)
    parsel = load_shapefile(parsel_path)

    # Calculate areas
    mahalle = calculate_areas(mahalle, 'area_mahalle')
    parsel = calculate_areas(parsel, 'area_parsel')

    # Calculate intersections
    detailed_results, summary_results = calculate_intersections(mahalle, parsel)

    # Save results to Excel
    save_to_excel(detailed_results, summary_results, output_file)


# Example usage
if __name__ == '__main__':
    # Fill in the file paths
    mahalle_file = ''  # Provide the path to the mahalle shapefile
    parsel_file = ''   # Provide the path to the parsel shapefile
    output_excel = ''  # Provide the output Excel file path

    main(mahalle_file, parsel_file, output_excel)
