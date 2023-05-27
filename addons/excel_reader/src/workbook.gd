#============================================================
#    Workbook
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:52:01
# - version: 4.0
#============================================================
class_name ExcelWorkbook


var zip_reader : ZIPReader
var xml_data : ExcelXMLData
var sheets : Dictionary = {}
var sheet_path_list : Array[String] = []
var data_value_list : Array = []


#============================================================
#  内置
#============================================================
func _init(zip_reader: ZIPReader, sheet_xml_path: String):
	self.zip_reader = zip_reader
	self.xml_data = ExcelXMLData.new(zip_reader, sheet_xml_path)
	
	for file in zip_reader.get_files():
		if file.begins_with("xl/worksheets/"):
			sheet_path_list.append(file)
	self.sheet_path_list.erase("xl/worksheets/")
	
	var sharedStrings = ExcelXMLData.new(zip_reader, "xl/sharedStrings.xml")
	for si_node in sharedStrings.get_root().get_children():
		if si_node.get_child_count() > 0:
			var t_node = si_node.get_child(0)
			data_value_list.append(t_node.get_value())


func _to_string():
	return "<%s:%s>" % ["Workbook", get_instance_id()]



#============================================================
#  自定义
#============================================================
func _create_sheet(xml_path: String) -> ExcelSheet:
	return ExcelSheet.new(zip_reader, xml_path, data_value_list)


func get_sheet_file_list() -> Array[String]:
	return sheet_path_list


func get_sheet_name_list() -> Array[String]:
	return Array(xml_data.get_first_node("workbook/sheets").get_children().map(
			func(item: ExcelXMLNode): return item.get_attr("name")
		)
		, TYPE_STRING
		, ""
		, null
	)


func get_sheets() -> Array[ExcelSheet]:
	if sheets.is_empty():
		for xml_path in sheet_path_list:
			sheets[xml_path] = _create_sheet(xml_path)
	return Array(sheets.values(), TYPE_OBJECT, "RefCounted", ExcelSheet)


func get_sheet(idx_or_name) -> ExcelSheet:
	var xml_path : String = get_sheet_file_list()[idx_or_name] \
		if idx_or_name is int \
		else ("xl/worksheets/" + idx_or_name)
	
	if not xml_path.ends_with(".xml"):
		xml_path += ".xml"
	
	# 没有这个 sheet 路径
	if not sheet_path_list.has(xml_path):
		printerr("没有这个文件：", xml_path)
		return null
	
	# 还没加载这个数据则进行加载
	if not sheets.has(xml_path):
		sheets[xml_path] = _create_sheet(xml_path)
	
	return sheets[xml_path]

