#============================================================
#    Workbook
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:52:01
# - version: 4.0
#============================================================
class_name ExcelWorkbook


var xml_data : ExcelXMLData:
	set(v):
		assert(xml_data == null)
		xml_data = v

var _zip_reader : ZIPReader
var _sheets : Dictionary = {}
var _sheet_files : Array[String] = []
var _data_value_list : Array = []


#============================================================
#  内置
#============================================================
func _init(zip_reader: ZIPReader, sheet_xml_path: String):
	self._zip_reader = zip_reader
	self.xml_data = ExcelXMLData.new(_zip_reader, sheet_xml_path)
	
	# 表单文件列表
	for file in _zip_reader.get_files():
		if file.begins_with("xl/worksheets/") and file.ends_with(".xml"):
			_sheet_files.append(file)
	self._sheet_files.erase("xl/worksheets/")
	
	# 获取值列表，string 类型单元格数据的缓存
	var sharedStrings = ExcelXMLData.new(_zip_reader, "xl/sharedStrings.xml")
	for si_node in sharedStrings.get_root().get_children():
		if si_node.get_child_count() > 0:
			var t_node = si_node.get_child(0)
			_data_value_list.append(t_node.get_value())


func _to_string():
	return "<%s#%s>" % ["Workbook", get_instance_id()]



#============================================================
#  自定义
#============================================================
func _create_sheet(xml_path: String) -> ExcelSheet:
	return ExcelSheet.new(_zip_reader, xml_path, _data_value_list)


func get_sheet_files() -> Array[String]:
	return _sheet_files


func get_sheet_name_list() -> Array[String]:
	return Array(xml_data.get_first_node("workbook/_sheets").get_children().map(
			func(item: ExcelXMLNode): return item.get_attr("name")
		)
		, TYPE_STRING
		, ""
		, null
	)


func get_sheets() -> Array[ExcelSheet]:
	if _sheets.is_empty():
		for xml_path in _sheet_files:
			_sheets[xml_path] = _create_sheet(xml_path)
	return Array(_sheets.values(), TYPE_OBJECT, "RefCounted", ExcelSheet)


func get_sheet(idx_or_name) -> ExcelSheet:
	var xml_path : String = get_sheet_files()[idx_or_name] \
		if idx_or_name is int \
		else ("xl/worksheets/" + idx_or_name)
	
	if not xml_path.ends_with(".xml"):
		xml_path += ".xml"
	
	# 没有这个 sheet 路径
	if not _sheet_files.has(xml_path):
		printerr("没有这个文件：", xml_path)
		return null
	
	# 还没加载这个数据则进行加载
	if not _sheets.has(xml_path):
		_sheets[xml_path] = _create_sheet(xml_path)
	
	return _sheets[xml_path]

