extends Node2D


var excel = ExcelFile.open_file("C:\\Users\\z\\Desktop\\role_data.xlsx")


func _ready():
	var workbook = excel.get_workbook()
	var sheet = workbook.get_sheet(0) as ExcelSheet
	
	var table_data = sheet.get_table_data()
	print(JSON.stringify(table_data, "\t"))
	

