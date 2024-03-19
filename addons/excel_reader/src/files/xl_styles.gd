#============================================================
#    Xl Styles
#============================================================
# - author: zhangxuetu
# - datetime: 2024-03-19 13:07:02
# - version: 4.2.1
#============================================================
class_name ExcelXlStyle
extends ExcelXlBase


# 单元格数字格式化代码数据
var _num_to_format_code_dict : Dictionary = {}
# ID 对应的格式化代码
var _id_to_format_code : Dictionary = {}

# 单元格样式的列表
var _cell_style_xfs : Array = []
# 已应用的单元格样式
var _cell_xfs : Array = []
# 应用的数字对应的样式数据
var _num_to_style_dict : Dictionary = {}



func _get_xl_path():
	return "xl/styles.xml"


func _init_data():
	var num_fmts_node = xml_file.get_root().find_first_node("numFmts")
	for child in num_fmts_node.get_children():
		var numFmtId = child.get_attr("numFmtId")
		_num_to_format_code_dict[int(numFmtId)] = child.get_attr("formatCode")
	
	# 单元格样式
	var cellStyleXfs = xml_file.get_root().find_first_node("cellStyleXfs")
	for child in cellStyleXfs.get_children():
		_cell_style_xfs.append(child.get_attr_map())
	
	# 已应用的单元格
	var cellXfs = xml_file.get_root().find_first_node("cellXfs")
	for child in cellXfs.get_children():
		_cell_xfs.append(child.get_attr_map())
		# 应用的数字格式的数
		var applyNumberFormat = child.get_attr("applyNumberFormat", null)
		if applyNumberFormat != null:
			_num_to_style_dict[ int(applyNumberFormat) ] = child.get_attr_map()


# 获取这个 num 的格式
func get_format_by_num(num: int) -> String:
	var cell_xfs_data : Dictionary = _num_to_style_dict[num]
	var num_fmt_id : int = int(cell_xfs_data["numFmtId"])
	return _num_to_format_code_dict[num_fmt_id]

