#============================================================
#    Excel Data
#============================================================
# - author: zhangxuetu
# - datetime: 2024-06-16 12:27:57
# - version: 4.2.1
#============================================================
class_name ExcelFileData 
extends Resource


## 表单名称列表
@export var sheet_name_list : Array[String] = []
## 表单名对应的数据
@export var data : Dictionary = {}


static func load_file(path: String) -> ExcelFileData:
	return load(path) as ExcelFileData


func _to_string() -> String:
	return "<ExcelFileData#%d>" % get_instance_id()


func get_sheet_data(sheet_name: String) -> Dictionary:
	return data.get(sheet_name, {})
