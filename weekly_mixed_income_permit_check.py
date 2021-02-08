import pandas as pd
import os
import geopandas as gpd
from datetime import datetime

# file paths and variables
todays_date = datetime.today().strftime('%m-%d-%y')
project_tracker_onedrive = os.environ.get('ProjectTrackerOneDrive')
project_tracker_input = project_tracker_onedrive + '/Project Tracker.xlsx'
project_tracker_output = project_tracker_onedrive + '/Project Tracker Permits/Project Tracker Permits - Week of ' + todays_date + '.xlsx'

# read the shared project tracker document
project_tracker_shared = pd.read_excel(project_tracker_input, 'Sheet1', engine='openpyxl')

# rename the ZP App # column in the shared project tracker
project_tracker_shared.rename(columns={"ZP App #": "ZP_App_Num"}, inplace=True)

# get comm & res building permits AND zoning permits that do not have a null ZONINGPERMITS value and were submitted in the past week
past_week_permits_url = 'https://services.arcgis.com/fLeGjb7u4uXqeF9q/arcgis/rest/services/PermitAppStatusEclipse/FeatureServer/0/query?where=APPLICATIONDATE+%3E+%28CURRENT_TIMESTAMP-INTERVAL%277%27DAY%29+AND+ZONINGPERMITS+is+not+null+AND+APPLICATIONDESCRIPTION+IN%28%27COMMERCIAL+BUILDING+PERMIT%27%2C%27RESIDENTIAL+BUILDING+PERMIT%27%2C%27ZONING+PERMIT%27%29&objectIds=&time=&geometry=&geometryType=esriGeometryEnvelope&inSR=&spatialRel=esriSpatialRelIntersects&resultType=none&distance=0.0&units=esriSRUnit_Meter&returnGeodetic=false&outFields=*&returnGeometry=true&featureEncoding=esriDefault&multipatchOption=xyFootprint&maxAllowableOffset=&geometryPrecision=&outSR=&datumTransformation=&applyVCSProjection=false&returnIdsOnly=false&returnUniqueIdsOnly=false&returnCountOnly=false&returnExtentOnly=false&returnQueryGeometry=false&returnDistinctValues=false&cacheHint=false&orderByFields=&groupByFieldsForStatistics=&outStatistics=&having=&resultOffset=&resultRecordCount=&returnZ=false&returnM=false&returnExceededLimitFeatures=true&quantizationParameters=&sqlFormat=none&f=pgeojson'
past_week_permits_gdf = gpd.read_file(past_week_permits_url)

# convert dates and sort by date
past_week_permits_gdf['APPLICATIONDATE']=(pd.to_datetime(past_week_permits_gdf['APPLICATIONDATE'],unit='ms')) 
past_week_permits_gdf['LATESTREVIEWDUEDATE']=(pd.to_datetime(past_week_permits_gdf['LATESTREVIEWDUEDATE'],unit='ms')) 
past_week_permits_gdf['LATESTREVIEWCOMPLETEDDATE']=(pd.to_datetime(past_week_permits_gdf['LATESTREVIEWCOMPLETEDDATE'],unit='ms')) 

# merge project tracker projects and permits from the past week using the zoning permit number
merged_df = project_tracker_shared.merge(past_week_permits_gdf, left_on='ZP_App_Num', right_on='ZONINGPERMITS')

# export the past week's permits
merged_df.to_excel(project_tracker_output, index=False)
