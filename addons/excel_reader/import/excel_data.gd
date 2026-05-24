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


## 以 [head_line] 行中的数据作为 [code]key[/code] 记录到数据列表中方便使用 
func get_format_array(sheet_name: String, head_line: int = 1) -> Array[Dictionary]:
	var sheet_data : Dictionary = get_sheet_data(sheet_name)
	var head_data : Dictionary = sheet_data[head_line]
	var new_data: Array[Dictionary] = []
	var item : Dictionary
	for line in sheet_data:
		if line != head_line:
			item = {}
			# 格式化
			for column in sheet_data[line]:
				item[ head_data.get(column) ] = sheet_data[line][column]
			new_data.append(item)
	return new_data
