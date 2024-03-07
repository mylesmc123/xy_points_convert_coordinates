import pandas as pd
import geopandas as gpd

xslx = pd.ExcelFile(r"V:\projects\p00832_ocd_2023_latz_hr\01_processing\SLaMM_Structure_XY_Data\SLaMM_XY_Albers.xlsx")
old_crs = 'PROJCS["USA_Contiguous_Albers_Equal_Area_Conic_USGS_version",GEOGCS["GCS_North_American_1983",DATUM["D_North_American_1983",SPHEROID["GRS_1980",6378137.0,298.257222101]],PRIMEM["Greenwich",0.0],UNIT["Degree",0.0174532925199433]],PROJECTION["Albers"],PARAMETER["false_easting",0.0],PARAMETER["false_northing",0.0],PARAMETER["central_meridian",-96.0],PARAMETER["standard_parallel_1",29.5],PARAMETER["standard_parallel_2",45.5],PARAMETER["latitude_of_origin",23.0],UNIT["Foot_US",0.3048006096012192]]'
new_crs = 'EPSG:6479'
output_xlsx = 'slamm_xy_6479.xlsx'

def xlsx_to_gdf(xlsx, sheet_name, old_crs, new_crs):
    df = pd.read_excel(xlsx, sheet_name=sheet_name, header=None)
    df.columns = ['old_x', 'old_y']
    gdf = gpd.GeoDataFrame(df, geometry=gpd.points_from_xy(df.old_x, df.old_y), crs=old_crs)
    gdf = gdf.to_crs(new_crs)
    return gdf

with pd.ExcelWriter(output_xlsx) as writer:
    for sheet_name in xslx.sheet_names:
        gdf = xlsx_to_gdf(xslx, sheet_name, old_crs, new_crs)
        # move geoemtry to x and y columns
        gdf['new_x'] = gdf.geometry.x
        gdf['new_y'] = gdf.geometry.y
        # write the gdf to the excel file with the sheet name = sheet_name.
        gdf.to_excel(writer, sheet_name=sheet_name, index=False)
    # save the excel file as slamm_xy_6479.xlsx
    writer.save()