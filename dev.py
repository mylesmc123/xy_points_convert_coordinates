# %%
import pandas as pd
import geopandas as gpd
import contextily as cx
# %%
# Load the data from the Excel file for each sheet into a DataFrame.
# The data is in the form of x and y coordinates for each structure and has no header.
xslx = pd.ExcelFile(r"V:\projects\p00832_ocd_2023_latz_hr\01_processing\SLaMM_Structure_XY_Data\SLaMM_XY_Albers.xlsx")
# read the number of sheets in the Excel file
len(xslx.sheet_names)
# %%
df= pd.read_excel(xslx, sheet_name=0, header=None)
df.columns = ['x', 'y']
df
# %%
# Get sheet name
sheet_name = xslx.sheet_names[0]
sheet_name

# %%
old_crs = 'PROJCS["USA_Contiguous_Albers_Equal_Area_Conic_USGS_version",GEOGCS["GCS_North_American_1983",DATUM["D_North_American_1983",SPHEROID["GRS_1980",6378137.0,298.257222101]],PRIMEM["Greenwich",0.0],UNIT["Degree",0.0174532925199433]],PROJECTION["Albers"],PARAMETER["false_easting",0.0],PARAMETER["false_northing",0.0],PARAMETER["central_meridian",-96.0],PARAMETER["standard_parallel_1",29.5],PARAMETER["standard_parallel_2",45.5],PARAMETER["latitude_of_origin",23.0],UNIT["Foot_US",0.3048006096012192]]'
new_crs = 'EPSG:6479'
# %%
# Convert the DataFrame to a GeoDataFrame with crs of old_crs.
gdf = gpd.GeoDataFrame(df, geometry=gpd.points_from_xy(df.x, df.y), crs=old_crs)
# Reproject the GeoDataFrame to new_crs.
gdf = gdf.to_crs(new_crs)
gdf
# %%
plot_gdf = gdf.to_crs('EPSG:4326')
# plot to folium map wtih the default basemap
import folium
m = folium.Map(location=[plot_gdf.geometry.y.mean(), plot_gdf.geometry.x.mean()], zoom_start=10, control_scale=True)
folium.GeoJson(plot_gdf).add_to(m)
m
# %%
# buid function to convert the xlsx sheet to gdf with conversion to new_crs
def xlsx_to_gdf(xlsx, sheet_name, old_crs, new_crs):
    df = pd.read_excel(xlsx, sheet_name=sheet_name, header=None)
    df.columns = ['old_x', 'old_y']
    gdf = gpd.GeoDataFrame(df, geometry=gpd.points_from_xy(df.old_x, df.old_y), crs=old_crs)
    gdf = gdf.to_crs(new_crs)
    return gdf
# %%
# create a blank excel file
with pd.ExcelWriter('slamm_xy_6479.xlsx') as writer:
    for sheet_name in xslx.sheet_names:
        gdf = xlsx_to_gdf(xslx, sheet_name, old_crs, new_crs)
        # move geoemtry to x and y columns
        gdf['new_x'] = gdf.geometry.x
        gdf['new_y'] = gdf.geometry.y
        # write the gdf to the excel file with the sheet name = sheet_name.
        gdf.to_excel(writer, sheet_name=sheet_name, index=False)
    # save the excel file as slamm_xy_6479.xlsx
    writer.save()

# %%
