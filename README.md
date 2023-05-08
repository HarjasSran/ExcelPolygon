# ExcelPolygon
The idea of this is to read from 2 different Excel files to determine if a point is within a Polygon.
# How it works
One Excel file will contain a bunch of points, the other Excel file will contain all the Polygons which has multiple different vertices
The Python script will read through both of these Excel files and compare the information to check if any of the points are in any of the Polygons.
Each Polygon has a name. The name of the Polygon will be stored into the Excel file with all the points. Under the section named "Insde Polygon"
# Note
Both Excel files must be closed before running the code.
