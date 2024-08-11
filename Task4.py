# Importing necessary libraries
import pandas as pd
from sklearn.impute import KNNImputer
from sklearn.cluster import KMeans
import folium
from folium.plugins import MarkerCluster

# Loading outlet information from the Excel file
# 'sheet_name' specifies the sheet from which to read data
outlet_info = pd.read_excel('Worksheet in Assessment question (006).xlsx',
                            sheet_name='outlet info')

# Filling the missing latitude and longitude values
imputer = KNNImputer(n_neighbors=5)
# Extracting latitude and longitude columns
coords = outlet_info[['latitude', 'longitude']]
# Imputing missing values
coords_imputed = imputer.fit_transform(coords)

# Updating the outlet_info DataFrame with the imputed latitude and longitude
outlet_info[['latitude', 'longitude']] = coords_imputed

# Performing clustering using KMeans with 50 clusters
kmeans = KMeans(n_clusters=50, random_state=0)
outlet_info['cluster'] = kmeans.fit_predict(coords_imputed)

# Assigning representative names based on the cluster number
# The representative names are generated based on cluster number, starting from 'rep#1'
outlet_info['representative'] = 'rep#' + (outlet_info['cluster'] + 1).astype(str)

# Defining the center location for the map as the average latitude and longitude
center_lat = outlet_info['latitude'].mean()
center_lon = outlet_info['longitude'].mean()

# Creating a folium map centered at the average location with a zoom level of 10
# 'MarkerCluster' will help in grouping nearby markers
map_clusters = folium.Map(location=[center_lat,
                                    center_lon],
                          zoom_start=10)

# Adding marker clustering to the map
marker_cluster = MarkerCluster().add_to(map_clusters)

# Plotting each outlet on the map with a marker
for _, row in outlet_info.iterrows():
    folium.Marker(
        location=[row['latitude'], row['longitude']],  # Marker location
        popup=f"Outlet: {row['outlet']}<br>Cluster: {row['cluster']}<br>Representative: {row['representative']}",
        # Popup displays outlet ID, cluster, and representative
        icon=folium.Icon(color='blue' if row['cluster'] % 2 == 0 else 'green')  # Marker color based on cluster
    ).add_to(marker_cluster)    # 'add_to(marker_cluster)' adds the marker to the marker cluster

# Saving the map to an HTML file
map_clusters.save('outlet_clusters_map.html')

# Exporting the outlet-representative mapping to an Excel file
outlet_info[['outlet', 'representative']].to_excel('outlet_representative_mapping.xlsx', index=False)


"""
Justification for the methodology used ->

For task 4, the methodology involves using KMeans clustering and KNN imputation to optimize the assignment of company 
representatives to mobile financial service outlets. KNN imputation is employed to handle missing latitude and 
longitude values, ensuring a complete dataset for clustering. The KMeans algorithm is then used to segment the 
outlets into 50 distinct clusters, each representing a group of outlets with similar geographical locations. This 
clustering helps in efficiently assigning each outlet to a representative, balancing the workload among the available 
representatives. By visualizing the clustered outlets on a map using Folium's `MarkerCluster`, the approach not only 
facilitates clear spatial distribution but also enhances the efficiency of routing and assignment by grouping nearby 
outlets together. This method ensures that each representative's area of service is optimized, aiding in effective 
coverage and resource allocation."""
