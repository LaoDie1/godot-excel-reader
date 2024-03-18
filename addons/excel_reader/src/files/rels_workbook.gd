#============================================================
#    Rels Workbook
#============================================================
# - author: zhangxuetu
# - datetime: 2024-03-18 16:56:46
# - version: 4.0
#============================================================
class_name ExcelRelsWorkbook
extends ExcelXlFileBase


## 关系类型
const Type = {
	WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
	THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
	STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
	SHARED_STRINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
	CELL_IMAGE = "http://www.wps.cn/officeDocument/2020/cellImage",
}

var _ids = {}
var _max_id = 0


#============================================================
#  内置
#============================================================
func _init(workbook: ExcelWorkbook):
	super._init(workbook)
	
	for child in xml_file.get_root().get_children():
		var id = int(child.get_attr("Id"))
		_ids[id] = null
		if _max_id < id:
			_max_id = id


#============================================================
#  自定义
#============================================================
#@override
func _get_xl_file_path() -> String:
	return "xl/_rels/workbook.xml.rels"


## 添加关系。type 为上方的 Type 中的其中一项。返回这个关系的 ID
##[br]这个 file_path 必须是完整的路径
func add_relationship(type: String, file_path: String) -> String:
	if file_path.begins_with("xl/"):
		file_path = file_path.substr(3)
	elif file_path.begins_with("/xl/"):
		file_path = file_path.substr(4)
	else:
		assert(false)
	
	while _ids.has(_max_id):
		_max_id += 1
	
	var id : String = "rId" + str(_max_id)
	var relationship = ExcelXMLNode.create("Relationship", true, {
		"Id": id,
		"Type": type,
		"Target": file_path,
	})
	xml_file.get_root().add_child(relationship)
	workbook.add_changed_file(ExcelWorkbook.FilePaths.RELS_WORKBOOK)
	return id


## 获取这个 ID 的文件的路径
func get_file_path_by_id(id: String) -> String:
	for child in xml_file.get_root().get_children():
		if child.get_attr("Id") == id:
			return child.get_attr("Target")
	return ""


## 获取 ID 对应的路径的字典
func get_id_to_path_dict() -> Dictionary:
	var dict : Dictionary = {}
	for child in xml_file.get_root().get_children():
		var id = child.get_attr("Id")
		var path = child.get_attr("Target")
		dict[id] = "xl/" + path
	return dict

