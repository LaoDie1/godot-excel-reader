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
var _data_value_list : Array = []
var _sheet_data_list : Array[Dictionary] = []

var _rels : ExcelXMLData  # rid data
var _rid_to_path_map : Dictionary = {}


#============================================================
#  内置
#============================================================
func _init(zip_reader: ZIPReader):
	self._zip_reader = zip_reader
	self.xml_data = ExcelXMLData.new(_zip_reader, "xl/workbook.xml")
	self._rels = ExcelXMLData.new(_zip_reader, "xl/_rels/workbook.xml.rels")
	
	# RID file path
	for child in _rels.get_root().get_children():
		var id = child.get_attr("Id")
		var target_path = child.get_attr("Target") # Files in the xl directory
		self._rid_to_path_map[id] = target_path
	
	# Sheets 
	var sheets = xml_data.get_root().get_first_node("sheets")
	var rid : String
	var sheet_name : String
	for child in sheets.get_children():
		rid = child.get_attr("r:id")
		sheet_name = child.get_attr("name")
		_sheet_data_list.append({
			"rid": rid,
			"sheet_name": sheet_name,
			"path": _rid_to_path_map[rid],
		})
	
	# 获取值列表，string 类型单元格数据的缓存（共享的字符串）
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
	return ExcelSheet.new(_zip_reader, "xl".path_join(xml_path), _data_value_list)


func get_sheet_files() -> Array[String]:
	return Array(_sheet_data_list.map(func(item): 
		return item["path"]
	), TYPE_STRING, &"", null)


func get_sheet_name_list() -> Array[String]:
	return Array(_sheet_data_list.map(func(item): 
		return item["sheet_name"]
	), TYPE_STRING, &"", null)


func get_sheets() -> Array[ExcelSheet]:
	if _sheets.is_empty():
		for xml_path in get_sheet_files():
			_sheets[xml_path] = _create_sheet(xml_path)
	return Array(_sheets.values(), TYPE_OBJECT, "RefCounted", ExcelSheet)


func get_path_by_sheet_name(sheet_name: String) -> String:
	for data in _sheet_data_list:
		if data["sheet_name"] == sheet_name:
			return data["path"]
	return ""


func get_sheet(idx_or_sheet_name) -> ExcelSheet:
	var xml_path : String = get_sheet_files()[idx_or_sheet_name] \
		if idx_or_sheet_name is int \
		else get_path_by_sheet_name(idx_or_sheet_name)
	
	if not xml_path.ends_with(".xml"):
		xml_path += ".xml"
	
	# 没有这个 sheet 路径
	if not get_sheet_files().has(xml_path):
		printerr("没有这个文件：", xml_path)
		return null
	
	# 还没加载这个数据则进行加载
	if not _sheets.has(xml_path):
		_sheets[xml_path] = _create_sheet(xml_path)
	
	return _sheets[xml_path]

