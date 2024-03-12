#============================================================
#    Test
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 19:06:07
# - version: 4.0
#============================================================
@tool
extends EditorScript


var excel = ExcelFile.open_file("C:\\Users\\Administrator\\Downloads\\user.xlsx")


func _run():

	var workbook = excel.get_workbook()

	#var cell_images_xml_file = workbook.get_xml_file("xl/cellimages.xml")
	#var etc_cell_images_node = cell_images_xml_file.get_root()
	#var xdr_pic = etc_cell_images_node.find_first_node_by_path("etc:cellImage/xdr:pic")
	#for node in xdr_pic.find_nodes_by_path("././a:./(a:.)"):
		#print(node.get_type())
	#return

	var sheet = workbook.get_sheet(0) as ExcelSheet
	sheet.alter(3, 1, 5555)
	print( sheet.get_xml_root().to_xml() )
	#sheet.save()

