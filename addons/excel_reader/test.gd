#============================================================
#    Test
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 19:06:07
# - version: 4.0
#============================================================
@tool
extends EditorScript


var excel = ExcelFile.open_file("C:\\Users\\z\\Desktop\\example.xlsx")


func _run():
	var workbook = excel.get_workbook()
	print(workbook.get_sheet_files())
	print(workbook.get_sheet_name_list())
	
	
	var sheet = workbook.get_sheet(0) as ExcelSheet
	var table_data = sheet.get_table_data()
	print(JSON.stringify(table_data, "\t"))
	
	return
	
	
	for row in table_data:
		var column_data = table_data[row]
		for column in column_data:
			print(column_data[column])
	

