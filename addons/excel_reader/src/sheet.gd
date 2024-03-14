#============================================================
#    Sheet
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:51:59
# - version: 4.0
#============================================================
class_name ExcelSheet


const MetaKey = {
	TABLE_DATA = "_table_data",
	ROW_NODE_DICT = "_row_node_dict",
	COLUMN_NODE_DICT = "_column_node_dict",
}

var workbook : ExcelWorkbook
var xml_file : ExcelXMLFile:
	set(v):
		assert(xml_file == null)
		xml_file = v

var _xml_path : String
var _col_row_regex : RegEx = RegEx.new()
var _image_regex : RegEx = RegEx.new()


#============================================================
#  内置 
#============================================================
func _init(workbook: ExcelWorkbook, sheet_xml_path: String):
	self.workbook = workbook
	self.xml_file = workbook.get_xml_file(sheet_xml_path)
	self._col_row_regex.compile("([A-Z]+)([0-9]+)")
	self._image_regex.compile("=DISPIMG\\(\"(?<rid>\\w+)\",\\d+\\)")


func _to_string():
	return "<%s#%s>" % ["Sheet", get_instance_id()]


#============================================================
#  自定义
#============================================================
func get_xml_root() -> ExcelXMLNode:
	return xml_file.get_root()

func get_xml_path() -> String:
	return _xml_path


## 获取表单中的数据。示例：
##[codeblock]
##var table_data = sheet.get_table_data()
##var value = table_data[row][column]
##print(value)
##[/codeblock]
func get_table_data() -> Dictionary:
	if not has_meta(MetaKey.TABLE_DATA):
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
				match ExcelDataUtil.get_data_type(column_node):
					ExcelDataUtil.DataType.STRING:
						var value_idx = int(value)
						# 如果是字符串，则进行转换
						column_to_data[coords.x] = workbook.get_shared_string(value_idx)
					
					ExcelDataUtil.DataType.EXPRESSION:
						column_to_data[coords.x] = workbook.convert_image(value)
					
					ExcelDataUtil.DataType.NUMBER:
						column_to_data[coords.x] = float(value)
			
			var row : int = int(row_node.get_attr("r"))
			row_to_column_data[row] = column_to_data
		
		set_meta(MetaKey.TABLE_DATA, row_to_column_data)
	return get_meta(MetaKey.TABLE_DATA)


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
	var value : String = ""
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
	if not has_meta(MetaKey.ROW_NODE_DICT):
		var xml_node_dict = {}
		var sheet_data_node = get_xml_root().find_first_node("sheetData")
		for row_node in sheet_data_node.get_children():
			var row = int(row_node.get_attr("r"))
			xml_node_dict[row] = row_node
		set_meta(MetaKey.ROW_NODE_DICT, xml_node_dict)
	return get_meta(MetaKey.ROW_NODE_DICT)


func _get_column_node_dict() -> Dictionary:
	if not has_meta(MetaKey.COLUMN_NODE_DICT):
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
		set_meta(MetaKey.COLUMN_NODE_DICT, xml_node_dict)
	return get_meta(MetaKey.COLUMN_NODE_DICT)


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
		column_node = ExcelXMLNode.create("c", false)
		var column_r = _to_26_base(column) + row_node.get_attr("r") # 单元格坐标位置
		column_node.set_attr("r", column_r)
		row_node.add_child(column_node)
		
		# 更新 row 的 spans 值
		var spans = _get_spans(row_node)
		if spans.from > column or spans.to < column:
			row_node.set_attr("spans", "%d:%d" % [ 
				min(spans.from, column), 
				max(spans.to, column) 
			])
		
		# 更新 sheet 的维度
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
			ExcelDataUtil.set_text(workbook, column_node, value)
			
		TYPE_INT, TYPE_FLOAT:
			ExcelDataUtil.set_number(column_node, value)
			
		TYPE_OBJECT:
			assert(value is Image)
			ExcelDataUtil.set_image(workbook, column_node, value)
			
		_:
			assert(false, "错误的数据类型")
	
	remove_meta(MetaKey.TABLE_DATA)
	
	# 调用过这个方法的 xml 路径都会记录到 workbook 中
	# 保存时自动更新数据
	workbook.add_changed_file(xml_file.get_xml_path())

