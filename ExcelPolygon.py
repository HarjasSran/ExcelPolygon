import pandas as pd


#Function using to check if x,y point is inside polygon
def point_in_polygon(x, y, poly):
    """
    Check if a point (x, y) is inside a polygon defined by a list of (x, y) vertices.
    """
    n = len(poly)
    inside = False

    # Loop through all edges of the polygon
    for i in range(n):
        x1, y1 = poly[i]
        x2, y2 = poly[(i + 1) % n]

        # Check if the point is on the right side of the edge
        if y > min(y1, y2):
            if y <= max(y1, y2):
                if x <= max(x1, x2):
                    if y1 != y2:
                        x_inters = (y - y1) * (x2 - x1) / (y2 - y1) + x1
                        if x1 == x2 or x <= x_inters:
                            inside = not inside

    return inside


# Read the Excel files
df = pd.read_excel(r'C:\Users\HarjasSran\OneDrive - AMICO\Documents\Programs\GeofenceTracing\input_ouput.xlsx', sheet_name = 'Sheet1')
geofence = pd.read_excel(r'C:\Users\HarjasSran\OneDrive - AMICO\Documents\Programs\GeofenceTracing\Geofences.xlsx', sheet_name = 'Sheet1')


# xfile = openpyxl.load_workbook(r'C:\Users\HarjasSran\OneDrive - AMICO\Documents\Programs\GeofenceTracing\input_ouput.xlsx')

# sheet = xfile.get_sheet_by_name('Sheet1')

# Iterate over each row of the Excel file
for i, row in df.iterrows():

    # Extract the point coordinates and polygon vertices from the row
    x, y = row['Point X'], row['Point Y']

    # Iterate over each polygon in the row and check if the point is inside, J in range of number of polygons + 1
    is_inside = 'None'
    for j in range(1,5):
        poly = [(geofence.loc[j-1,f'Vertex {k} X'], geofence.loc[j-1,f'Vertex {k} Y']) for k in range (1,6)] # All verticies of the polygon until k which is max number of coordintaes +1
        
        #using the function, check if the current x, y points are inside the current polygon
        if point_in_polygon(x, y, poly):
            # if x, y points are a part of the current polygon, save the name of the polygon
            is_inside = geofence.loc[j-1, f'Name']
            
            # sheet[f'C{j+1}'] = is_inside
            # df.loc[(j, 'Inside Polygon')] = is_inside
            break

    #Save the name of the polygone into the correct row
    df.loc[i, 'Inside Polygon'] = is_inside

# #write new updates into the excel file
with pd.ExcelWriter(r'C:\Users\HarjasSran\OneDrive - AMICO\Documents\Programs\GeofenceTracing\input_ouput.xlsx') as writer:
    df.to_excel(writer, sheet_name='Sheet1', index = False)

# df.save(r'C:\Users\HarjasSran\OneDrive - AMICO\Documents\Programs\GeofenceTracing\input_ouput.xlsx')