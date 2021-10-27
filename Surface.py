# import Workbook from openpyxl
from openpyxl import Workbook

# import SurfaceChart, Reference, Series from openpyxl.chart sub_module .
from openpyxl.chart import SurfaceChart, Reference, Series

# Call a Workbook() function of openpyxl
# to create a new blank Workbook object
wb = Workbook()

# Get workbook active sheet
# from the active attribute.
ws = wb.active

# given data
data = [
	[None, 10, 20, 30, 40, 50, ],
	[0.1, 15, 65, 105, 65, 15, ],
	[0.2, 35, 105, 170, 105, 35, ],
	[0.3, 55, 135, 215, 135, 55, ],
	[0.4, 75, 155, 240, 155, 75, ],
	[0.5, 80, 190, 245, 190, 80, ],
	[0.6, 75, 155, 240, 155, 75, ],
	[0.7, 55, 135, 215, 135, 55, ],
	[0.8, 35, 105, 170, 105, 35, ],
	[0.9, 15, 65, 105, 65, 15],
]

# write content of each row in 1st and 2nd
# column of the active sheet respectively .
for row in data:
	ws.append(row)

# Create object of SurfaceChart class
chart = SurfaceChart()

# create data for plotting
labels = Reference(ws, min_col = 1, min_row = 2, max_row = 10)
data = Reference(ws, min_col = 2, max_col = 6, min_row = 1, max_row = 10)

# adding data to the Surface chart object
chart.add_data(data, titles_from_data = True)

# set labels in the chart object
chart.set_categories(labels)

# set the title of the chart
chart.title = "Surface Chart"

# set style of the chart
chart.style = 26

# add chart to the sheet
# the top-left corner of a chart
# is anchored to cell H2 .
ws.add_chart(chart, "H2")

# save the file
wb.save("Surface.xlsx")
