# Intersection-Area-Calculator

This Python tool calculates intersection areas between two shapefiles (e.g., mahalle and parsel). The results are saved in an Excel file with both detailed and summary information.

## Features

    •    Geometric Validity Check: Ensures all geometries in the shapefiles are valid.
    •    Area Calculation: Computes areas for each geometry and intersection area between overlapping geometries.
    •    Intersection Details: Provides detailed information for each intersected geometry pair.
    •    Summary Results: Aggregates intersected areas per mahalle geometry.
    •    Excel Export: Saves the results in a structured Excel file with separate sheets for detailed and summary data.

## Requirements

To run the project, ensure you have the following installed:
    •    Python 3.8+
    •    Libraries:
    •    geopandas
    •    pandas
    •    shapely
    •    openpyxl

## Example Output

    •    Detaylar (Detailed Results): Contains columns such as mahalle ID, parsel ID, and intersected area.
    •    Özet (Summary Results): Aggregates intersected areas for each mahalle geometry.

# Note: Ensure that both shapefiles contain valid geometries and required attributes like IDs. 
