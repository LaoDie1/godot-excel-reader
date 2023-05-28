#============================================================
#    Sheet
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:51:59
# - version: 4.0
#============================================================
class_name ExcelSheet


var reg : RegEx = RegEx.new()
var xml_data : ExcelXMLData:
	set(v):
		assert(xml_data == null)
		xml_data = v

var _data_value_list : Array = []


#============================================================
#  内置 
#============================================================
func _init(zip_reader: ZIPReader, sheet_xml_path: String, data_value_list: Array):
	self.xml_data = ExcelXMLData.new(zip_reader, sheet_xml_path)
	self._data_value_list = data_value_list
	self.reg.compile("([A-Z]+)([0-9]+)")


func _to_string():
	return "<%s#%s>" % ["Sheet", get_instance_id()]


#============================================================
#  自定义
#============================================================
func get_xml_root() -> ExcelXMLNode:
	return xml_data.get_root()


## Return Data Format Example
##[codeblock]
##var table_data = sheet.get_table_data()
##var value = table_data[row][column]
##[/codeblock]
func get_table_data() -> Dictionary:
	if not has_meta("table_data"):
		
		var row_to_column_data = {}
		var sheet_data = get_xml_root().get_first_node("sheetData")
		
		for row_node in sheet_data.get_children():
	#		var spans = row_node.get_attr("spans")
	#		var from_column = spans.split(":")[0]
	#		var to_column = spans.split(":")[1]
	#		prints(from_column, to_column)
			
			var column_to_data : Dictionary = {}
			for column_node in row_node.get_children():
				var value_node = column_node.get_child(0)
				if value_node == null:
					continue
				var value = value_node.get_value()
				# 所在行列坐标
				var coords = to_coords(column_node.get_attr("r"))
				# 判断数据类型
				var data_type = column_node.get_attr("t")
				if data_type == "s":
					var value_idx = int(value)
					# 如果是字符串，则进行转换
					column_to_data[coords.x] = _data_value_list[value_idx]
				
				else:
					var json = JSON.new()
					if json.parse(value) == OK:
						column_to_data[coords.x] = json.data
					else:
						column_to_data[coords.x] = value
			
			var row = row_node.get_attr("r")
			row_to_column_data[int(row)] = column_to_data
		
		set_meta("table_data", row_to_column_data)
	return get_meta("table_data")


func to_coords(r: String) -> Vector2i:
	var result = reg.search(r)
	var column_str = result.get_string(1)
	var row_str = result.get_string(2)
	
	var x : int = 0
	var column_length : int = column_str.length()
	for i in column_length:
		var num = (column_str.unicode_at(i) - 64)
		x += num * pow(26, column_length - 1 - i) 
	return Vector2i(x, row_str.to_int())

