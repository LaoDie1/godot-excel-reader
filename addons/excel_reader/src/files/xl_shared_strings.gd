#============================================================
#    Xl Shared Strings
#============================================================
# - author: zhangxuetu
# - datetime: 2024-03-19 00:22:14
# - version: 4.2.1
#============================================================
class_name ExcelXlSharedStrings
extends ExcelXlBase


var _shared_strings : Array[String] = []


func _init_data():
	# 获取值列表，string 类型单元格数据的缓存
	for si_node in get_xml_file().get_root().get_children():
		_shared_strings.append(si_node.get_full_value())


func _get_xl_path():
	return "xl/sharedStrings.xml"


## 更新共享的文字。返回这个字符串的索引
func update_shared_string_xml(text: String) -> int:
	var idx = _shared_strings.find(text)
	if idx > -1:
		return idx
	else:
		idx = _shared_strings.size()
		_shared_strings.append(text)
		
		var si_node = ExcelXMLNode.create("si", false)
		var t_node = ExcelXMLNode.create("t", false)
		t_node.value = text
		si_node.add_child(t_node)
		
		# 更新属性
		var sst_node = get_xml_file().get_root()
		sst_node.add_child(si_node)
		sst_node.set_attr("uniqueCount", _shared_strings.size())
		sst_node.set_attr("count", int(sst_node.get_attr("count"))+1 )
		
		notify_change()
		
		return idx


## 获取共享字符串
func get_shared_string(idx: int) -> String:
	return _shared_strings[idx]
