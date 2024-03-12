#============================================================
#    Sheet
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:51:59
# - version: 4.0
#============================================================
class_name ExcelSheet


var workbook : ExcelWorkbook
var xml_file : ExcelXMLFile:
	set(v):
		assert(xml_file == null)
		xml_file = v

var _string_value_cache : Array:
	get: return workbook._string_value_cache
var _col_row_regex : RegEx = RegEx.new()
var _image_regex : RegEx = RegEx.new()


#============================================================
#  内置 
#============================================================
func _init(workbook: ExcelWorkbook, sheet_xml_path: String):
	self.workbook = workbook
	self.xml_file = ExcelXMLFile.new(workbook, sheet_xml_path)
	self._col_row_regex.compile("([A-Z]+)([0-9]+)")
	self._image_regex.compile("=DISPIMG\\(\"(?<rid>\\w+)\",\\d+\\)")


func _to_string():
	return "<%s#%s>" % ["Sheet", get_instance_id()]


#============================================================
#  自定义
#============================================================
func get_xml_root() -> ExcelXMLNode:
	return xml_file.get_root()


## 获取表单中的数据。示例：
##[codeblock]
##var table_data = sheet.get_table_data()
##var value = table_data[row][column]
##print(value)
##[/codeblock]
func get_table_data() -> Dictionary:
	const META_KEY = "_table_data"
	if not has_meta(META_KEY):
		var row_to_column_data = {}
		var sheet_data_node = get_xml_root().find_first_node("sheetData")
		
		for row_node in sheet_data_node.get_children():
			var column_to_data : Dictionary = {}
			for column_node in row_node.get_children():
				var value_node = column_node.get_child(0)
				if value_node == null:
					continue
				var value = value_node.get_value()
				# 所在行列坐标
				var coords = _to_coords(column_node.get_attr("r"))
				# 判断数据类型
				var data_type = column_node.get_attr("t")
				if data_type == "s":
					var value_idx = int(value)
					# 如果是字符串，则进行转换
					column_to_data[coords.x] = workbook.shared_strings[value_idx]
				elif data_type == "str":
					column_to_data[coords.x] = convert_image(value)
				
				else:
					var json = JSON.new()
					if json.parse(value) == OK:
						column_to_data[coords.x] = json.data
					else:
						column_to_data[coords.x] = value
			
			var row = int(row_node.get_attr("r"))
			row_to_column_data[row] = column_to_data
		
		set_meta(META_KEY, row_to_column_data)
	return get_meta(META_KEY)


func convert_image(value):
	# 嵌入单元格的图片
	var result = _image_regex.search(value)
	if result:
		var rid = result.get_string("rid")
		return workbook.cellimages[rid]
	else:
		return value


func _get_spans(row_node: ExcelXMLNode):
	var spans = row_node.get_attr("spans")
	var from_column = int(spans.split(":")[0])
	var to_column = int(spans.split(":")[1])
	return {
		"from": from_column,
		"to": to_column,
	}

func _to_coords(r: String) -> Vector2i:
	var result = _col_row_regex.search(r)
	var column_str = result.get_string(1)
	var row_str = result.get_string(2)
	
	var column : int = 0
	var column_length : int = column_str.length()
	for i in column_length:
		var num = (column_str.unicode_at(i) - 64)
		column += num * pow(26, column_length - 1 - i)
	return Vector2i(column, row_str.to_int())


# 转为 26 进制
func _to_26_base(num: int) -> String:
	assert(num > 0)
	var value = ""
	for i in range(1, 16):
		var power_value = (26 ** i)
		var result : int = num / power_value
		if result > 0:
			value += char(result + 64)
		else:
			value += char(num + 64)
			break
		num -= power_value
	return value


func _get_row_node_dict() -> Dictionary:
	const META_KEY = "_row_node_dict"
	if not has_meta(META_KEY):
		var xml_node_dict = {}
		var sheet_data_node = get_xml_root().find_first_node("sheetData")
		for row_node in sheet_data_node.get_children():
			var row = int(row_node.get_attr("r"))
			xml_node_dict[row] = row_node
		set_meta(META_KEY, xml_node_dict)
	return get_meta(META_KEY)


func _get_column_node_dict() -> Dictionary:
	const META_KEY = "_column_node_dict"
	if not has_meta(META_KEY):
		var xml_node_dict = {}
		var row_node_dict = _get_row_node_dict()
		for row in row_node_dict:
			var row_node = row_node_dict[row]
			var columns = {}
			for column_node in row_node.get_children():
				var coord = _to_coords( column_node.get_attr("r") )
				var column = coord.x
				columns[column] = column_node
			xml_node_dict[row] = columns
		set_meta(META_KEY, xml_node_dict)
	return get_meta(META_KEY)


func get_value(row: int, column: int):
	var column_dict = _get_column_node_dict()
	if column_dict.has(row):
		var columns = column_dict[row]
		if columns.has(column):
			var column_node = columns[column]
			var value_node = column_node.find_first_node("v")
			return value_node.get_value()
	return null


## 修改表单值
func alter(row: int, column: int, value) -> void:
	assert(row > 0)
	assert(column > 0)
	#assert(value is String or value is Image)
	
	var column_node_dict = _get_column_node_dict()
	if (value is String and value == "") or typeof(value) == TYPE_NIL:
		if column_node_dict.has(row):
			column_node_dict[row].erase(column)
		return
	
	var columns : Dictionary
	if not column_node_dict.has(row):
		column_node_dict[row] = columns
	columns = column_node_dict[row]
	var column_node : ExcelXMLNode
	if columns.has(column):
		column_node = columns[column] as ExcelXMLNode
	else:
		var row_node_dict = _get_row_node_dict()
		var row_node : ExcelXMLNode
		if not row_node_dict.has(row):
			# 没有这一行，则新建
			row_node = ExcelXMLNode.create("row", false)
			row_node.set_attr("r", str(row))
			var sheet_data_node = get_xml_root().find_first_node("sheetData")
			row_node.set_attr("spans", "%s:%s" % [column, column])
			sheet_data_node.add_child(row_node)
			row_node_dict[row] = row_node
		row_node = row_node_dict[row]
		
		# 新数据的单元格
		var new_column_node = ExcelXMLNode.create("c", false)
		var column_r = _to_26_base(column) + row_node.get_attr("r") # 单元格坐标位置
		new_column_node.set_attr("r", column_r)
		row_node.add_child(new_column_node)
		column_node = new_column_node
		
		# 更新 row 的 spans 值
		var spans = _get_spans(row_node)
		if spans.to < column:
			row_node.set_attr("r", "%d:%d" % [ spans.from, column ])
		if spans.from > column:
			row_node.set_attr("r", "%d:%d" % [ column, spans.to ])
		
		# 更新 workbook 的维度
		var dimension_node = get_xml_root().find_first_node("dimension")
		var refs = dimension_node.get_attr("ref").split(":")
		var from = _to_coords(refs[0])
		var to = _to_coords(refs[1])
		from.x = min(from.x, column)
		from.y = min(from.y, row)
		to.x = max(to.x, column)
		to.y = max(to.y, row)
		dimension_node.set_attr("ref", "%s%s:%s%s" % [ 
			_to_26_base(from.x), from.y,
			_to_26_base(to.x), to.y,
		] )
	
	# 修改值
	match typeof(value):
		TYPE_STRING:
			ExcelDataType.set_text(column_node, value, _string_value_cache)
			#var update_result = ExcelDataType.set_text(column_node, value, _string_value_cache)
			#if update_result:
				# TODO 如果缓存发生改变，则在下一帧进行保存缓存
				#pass
			
		TYPE_INT, TYPE_FLOAT:
			ExcelDataType.set_number(column_node, value)
			
		TYPE_OBJECT:
			assert(value is Image)
			ExcelDataType.set_image(workbook, column_node, value)
			
		_:
			assert(false, "错误的数据类型")


func save():
	var new_file_path = workbook.file_path.get_basename() + "_.xlsx"
	var result = xml_file.update()
	workbook.save()
	print(  error_string(result)  )
