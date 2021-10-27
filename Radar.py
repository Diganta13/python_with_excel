# import Workbook from openpyxl
from openpyxl import Workbook

# import RadarChart, Reference from openpyxl.chart sub_module .
from openpyxl.chart import RadarChart, Reference

# Call a Workbook() function of openpyxl
# to create a new blank Workbook object
wb = Workbook()

# Get workbook active sheet
# from the active attribute.
ws = wb.active

# data given
data = [
	['Month', "Bulbs", "Seeds", "Flowers", "Trees & shrubs"],
	['Jan', 0, 2500, 500, 0, ],
	['Feb', 0, 5500, 750, 1500],
	['Mar', 0, 9000, 1500, 2500],
	['Apr', 0, 6500, 2000, 4000],
	['May', 0, 3500, 5500, 3500],
	['Jun', 0, 0, 7500, 1500],
	['Jul', 0, 0, 8500, 800],
	['Aug', 1500, 0, 7000, 550],
	['Sep', 5000, 0, 3500, 2500],
	['Oct', 8500, 0, 2500, 6000],
	['Nov', 3500, 0, 500, 5500],
	['Dec', 500, 0, 100, 3000 ],
]

# write content of each row in 1st and 2nd
# column of the active sheet respectively .
for row in data:
	ws.append(row)

# Create object of RadarChart class
chart = RadarChart()

# filled type of radar chart
chart.type = "filled"

# create data for plotting
labels = Reference(ws, min_col = 1, min_row = 2, max_row = 13)
data = Reference(ws, min_col = 2, max_col = 5, min_row = 2, max_row = 13)

# adding data to the Radar chart object
chart.add_data(data, titles_from_data = True)

# set labels in the chart object
chart.set_categories(labels)

# set the title of the chart
chart.title = "Radar Chart"

# set style of the chart
chart.style = 26

# delete y axis from the chart
chart.y_axis.delete = True

# add chart to the sheet
# the top-left corner of a chart
# is anchored to cell G2 .
ws.add_chart(chart, "G2")

# save the file
wb.save("Radar.xlsx")
