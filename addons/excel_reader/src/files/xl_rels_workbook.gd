#============================================================
#    Rels Workbook
#============================================================
# - author: zhangxuetu
# - datetime: 2024-03-18 16:56:46
# - version: 4.2.1
#============================================================
## workbook 中的文件关系。用于获取 rId 对应的文件路径
class_name ExcelXlRelsWorkbook
extends ExcelXlBase


var _data_list : Array[Dictionary] = []
var _rid_to_path_dict : Dictionary = {}


func _init_data():
	for child in xml_file.get_root().get_children():
		_data_list.append(child.get_attr_map())
		
		var rid : String = child.get_attr("Id")
		_rid_to_path_dict[rid] = "xl".path_join(child.get_attr("Target"))


func _get_xl_path() -> String:
	return "xl/_rels/workbook.xml.rels"


func record_path(rid: String, path: String) -> void:
	# 文件是有顺序的，末尾的必须都是 sheet 类型的文件，否则有的文件读取不出来
	if _rid_to_path_dict.has(rid):
		var new_rid_to_path_dict : Dictionary = {}
		# 数据向后偏移
		var id : int = int(rid)
		for k in _rid_to_path_dict:
			var k_id : int = int(k)
			if k_id >= id: # 这个 key 比参数 rid 大
				new_rid_to_path_dict["rId%d" % (k_id + 1)] = _rid_to_path_dict[k]
		_rid_to_path_dict = new_rid_to_path_dict
	_rid_to_path_dict[rid] = "xl".path_join(path)


## 添加关系。返回这个关系的 ID
##[br]这个 file_path 必须是完整的路径
func add_relationship(type: String, rid: String, file_path: String) -> String:
	if file_path.begins_with("xl/"):
		file_path = file_path.substr(3)
	elif file_path.begins_with("/xl/"):
		file_path = file_path.substr(4)
	else:
		assert(false)
	
	# 添加节点
	var relationship = ExcelXMLNode.create("Relationship", true, {
		"Id": rid,
		"Type": type,
		"Target": file_path,
	})
	xml_file.get_root().add_child_to(relationship, 0)
	
	# 记录
	record_path(rid, file_path)
	
	notify_change()
	return rid


## 获取这个 ID 的文件的路径
func get_path_by_id(id: String) -> String:
	for child in xml_file.get_root().get_children():
		if child.get_attr("Id") == id:
			return "xl/".path_join(child.get_attr("Target"))
	return ""


## 获取 ID 对应的路径的字典
func get_id_to_path_dict() -> Dictionary:
	var dict : Dictionary = {}
	for child in xml_file.get_root().get_children():
		var id = child.get_attr("Id")
		var path = child.get_attr("Target")
		dict[id] = "xl/" + path
	return dict


## 获取 Sheet 文件列表
func get_sheet_files() -> Array[String]:
	var list : Array[String] = []
	for data in _data_list:
		if data["Type"] == ExcelDataUtil.FileType.WORKSHEET:
			list.append("xl/" + data["Target"])
	return list


