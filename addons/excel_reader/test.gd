#============================================================
#    Test
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 19:06:07
# - version: 4.0
#============================================================
@tool
extends EditorScript


func _run():
	
	#print(Time.get_unix_time_from_datetime_string("1899-12-30T00:00:00"))
	#print( Time.get_datetime_string_from_unix_time(1036800 + 1121221) )
	#print(Time.get_datetime_string_from_unix_time(4529200000))
	var days = ((45292+1 - 70 * 365 - 19) * 86400 - 8 * 3600)
	print( Time.get_datetime_string_from_unix_time(days) )
	
	return
	
	var excel = ExcelFile.open_file("D:\\Downloads\\test_.xlsx")
	
	#var workbook = excel.get_workbook()
	#var cell_images_xml_file = workbook.get_xml_file("xl/cellimages.xml")
	#var etc_cell_images_node = cell_images_xml_file.get_root()
	#var xdr_pic = etc_cell_images_node.find_first_node_by_path("etc:cellImage/xdr:pic")
	#for node in xdr_pic.find_nodes_by_path("././a:./(a:.)"):
		#print(node.get_type())
	#return
	
	var workbook = excel.get_workbook()
	var sheet = workbook.get_sheet("Sheet1") as ExcelSheet
	#print( JSON.stringify(sheet.get_table_data(), "\t") )
	
	sheet.alter(5, 5, 10)
	#print( JSON.stringify(sheet.get_table_data(), "\t") )
	
	print(sheet.get_xml_root().to_xml())
	excel.save("D:\\Downloads\\test_2.xlsx")

