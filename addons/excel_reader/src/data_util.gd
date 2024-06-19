#============================================================
#    Data Util
#============================================================
# - author: zhangxuetu
# - datetime: 2024-03-12 14:28:17
# - version: 4.2.1
#============================================================
class_name ExcelDataUtil


const FileType = {
	WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
	THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
	STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
	SHARED_STRINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
	CELL_IMAGE = "http://www.wps.cn/officeDocument/2020/cellImage",
}

const ContentType = {
	WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
}

## 数据类型
const DataType = {
	NUMBER = "", ## 没有类型的都为数值类型（待定）
	STRING = "s", ## 字符型
	EXPRESSION = "str", ## 表达式
}

## 属性名
const PropertyName = {
	COLUMN_ROW = "r", ## 所在行或列
	DATA_TYPE = "t", ## 数据类型，对应到sharedStrings.xml中的sst元素
	CELL_FORMAT = "s", ## 单元格格式，对应 style.xml
}

static var _COLUMN_ROW_REGEX: RegEx:
	get:
		if _COLUMN_ROW_REGEX == null:
			_COLUMN_ROW_REGEX = RegEx.new()
			_COLUMN_ROW_REGEX.compile("([A-Z]+)([0-9]+)")
		return _COLUMN_ROW_REGEX


## 获取数据类型
static func get_data_type(cell_node: ExcelXMLNode) -> String:
	return cell_node.get_attr(PropertyName.DATA_TYPE)


## 日期时间戳转为字符串类型
static func date_stamp_to_string(date_stamp: int) -> String:
	var days = ((date_stamp + 1 - 70 * 365 - 19) * 86400 - 8 * 3600)
	var datetime = Time.get_datetime_string_from_unix_time(days)
	return datetime.split("T")[0]


## 获取这个 xml 的 bytes 数据
static func get_xml_file_data(xml_file: ExcelXMLFile) -> PackedByteArray:
	# 不使用缩进
	return xml_file \
		.to_xml(false) \
		.to_utf8_buffer()

static func get_xml_node_data(xml_node: ExcelXMLNode) -> PackedByteArray:
	return xml_node \
		.to_xml(0, false) \
		.to_utf8_buffer()

static func set_value(cell_node: ExcelXMLNode, value):
	var v_node = cell_node.find_first_node("v")
	if v_node == null:
		v_node = ExcelXMLNode.create("v", false)
		cell_node.add_child(v_node)
	v_node.value = value


static func set_text(workbook: ExcelWorkbook, cell_node: ExcelXMLNode, value: String):
	cell_node.remove_attr(PropertyName.CELL_FORMAT) # 字符串没有单元格格式，移除掉这个属性
	cell_node.set_attr(PropertyName.DATA_TYPE, DataType.STRING)
	
	# 记录到字符串到缓存值列表
	var string_idx = workbook.xl_shared_string.update_shared_string_xml(value)
	set_value(cell_node, string_idx)


static func set_number(
	cell_node: ExcelXMLNode, 
	value: int,
	cell_format: int = -1,
):
	cell_node.remove_attr(PropertyName.DATA_TYPE) # 移除属性类型。默认为数字类型
	if cell_format > -1:
		cell_node.set_attr(PropertyName.CELL_FORMAT, cell_format)
	set_value(cell_node, value)


#static func set_date_by_string(cell_node: ExcelXMLNode, value: String):
	# TODO 字符串转为时间戳
	# var time_stamp
	#set_number(cell_node, time_stamp, CellFormat.DATE)


## 修改数据
static func alter_value(
	workbook: ExcelWorkbook, 
	column_node: ExcelXMLNode,
	value
):
	match typeof(value):
		TYPE_STRING:
			set_text(workbook, column_node, value)
			
		TYPE_INT, TYPE_FLOAT:
			set_number(column_node, value)
			
		TYPE_OBJECT:
			assert(value is Image)
			set_image(workbook, column_node, value)
			
		_:
			assert(false, "错误的数据类型")


# FIXME 暂时遇到问题，未完成
static func set_image(
	workbook: ExcelWorkbook,
	cell_node: ExcelXMLNode, 
	image: Image, 
	descr: String = "",
):
	cell_node.remove_all_child()
	
	cell_node.remove_attr(PropertyName.CELL_FORMAT)
	cell_node.set_attr(PropertyName.DATA_TYPE, DataType.EXPRESSION)
	
	# 生成一个 IMAGE 的 ID
	var image_id = "ID_" % str(ResourceUID.create_id()).sha256_text().substr(32).to_upper()
	var cell_images_xml_file = workbook.get_xml_file(workbook.xl_cell_images._get_xl_path())
	
	var etc_cell_images_node = cell_images_xml_file.get_root()
	var xdr_pic = etc_cell_images_node.find_first_node_by_path("etc:cellImage/xdr:pic")
	var nv_pic_pr = xdr_pic.find_first_node_by_path("xdr:nvPicPr")
	
	# 最大索引值
	var last_idx = 0
	for child in nv_pic_pr.find_nodes_by_path("xdr:cNvPr"):
		var id = int(child.get_attr("id"))
		if id > last_idx:
			last_idx = id
	
	# 添加图片
	if descr == "":
		descr = image_id
	var xdr_c_nv_pr = ExcelXMLNode.create("xdr:cNvPr", false)
	xdr_c_nv_pr.set_attr("id", last_idx + 1)
	xdr_c_nv_pr.set_attr("name", image_id)
	xdr_c_nv_pr.set_attr("descr", descr)
	nv_pic_pr.add_child(xdr_c_nv_pr)
	
	# 创建节点
	var f_node = ExcelXMLNode.create("f", false)
	f_node.value = '_xlfn.DISPIMG("%s",1)' % image_id
	cell_node.add_child(f_node)
	set_value(cell_node, 'DISPIMG("%s",1)' % image_id)


## r属性转为坐标
static func to_coords(r: String) -> Vector2i:
	var result = _COLUMN_ROW_REGEX.search(r)
	var column_str = result.get_string(1)
	var row_str = result.get_string(2)
	return Vector2i( 
		to_10_base_by_26_base(column_str), 
		row_str.to_int()
	)


## 26 进制转为 10 进制
static func to_10_base_by_26_base(base_26: String) -> int:
	var column : int = 0
	var column_length : int = base_26.length()
	for i in column_length:
		var num = (base_26.unicode_at(i) - 64)
		column += num * pow(26, column_length - 1 - i)
	return column


# 10 进制转为 26 进制
static func to_26_base_by_10_base(dividend: int) -> String:
	const BASE = 26
	if dividend == 0:
		return "@"
	var result = []
	var quotient : int = dividend
	var remainder : int
	while quotient > 0:
		quotient = dividend / BASE
		remainder = dividend % BASE
		if remainder > 0:
			result.append(
				char( (remainder if remainder > 0 else BASE) + 64 )
			)
		else:
			result.append(char(BASE + 64))
			quotient -= 1
			if quotient > 0:
				result.append(char(quotient + 64))
			break
		dividend = quotient
	
	result.reverse()
	return "".join(result)


## 获取 row 的列范围
static func get_spans(spans: String) -> Dictionary:
	var from_column := int(spans.split(":")[0])
	var to_column := int(spans.split(":")[1])
	return {
		"from": from_column,
		"to": to_column,
	}


## 有效范围
static func to_dimension(rect: Rect2i) -> String:
	return "%s%s:%s%s" % [
		to_26_base_by_10_base(rect.position.x), rect.position.y,
		to_26_base_by_10_base(rect.end.x), rect.end.y
	]


## 根据数据添加节点
static func add_node_by_data(
	workbook: ExcelWorkbook, 
	sheet_data_node: ExcelXMLNode, # sheet.xml 文件的 sheetData 节点
	data: Dictionary,
):
	if data.is_empty():
		return
	
	# 原来的“每行对应的列的数据”
	var row_to_column_data : Dictionary = {}
	for row_node in sheet_data_node.get_children():
		var row : int = int(row_node.get_attr(PropertyName.COLUMN_ROW))
		var column_data : Dictionary = {
			"row_node": row_node, # “这一行的每个列”对应的行节点
		}
		var coords : Vector2i
		var column : int 
		for child in row_node.get_children():
			coords = to_coords(child.get_attr(PropertyName.COLUMN_ROW))
			column = coords.x
			column_data[column] = child
		row_to_column_data[row] = column_data
	
	# 添加数据
	var remove_children : Array[ExcelXMLNode] = []
	for row in data:
		var row_node : ExcelXMLNode 
		var min_column = INF
		var max_column : int = 0
		if row_to_column_data.has(row):
			row_node = row_to_column_data[row]["row_node"]
			var spans = get_spans(row_node.get_attr("spans"))
			min_column = spans.from
			max_column = spans.to
			
		else:
			row_node = ExcelXMLNode.create("row", false, {
				PropertyName.COLUMN_ROW: row,
			})
			row_to_column_data[row] = {
				"row_node": row_node,
			}
		
		# 添加列节点
		var column_to_node_dict : Dictionary = row_to_column_data[row]
		var column_data : Dictionary = data[row]
		for column in column_data:
			min_column = min(min_column, column)
			max_column = max(max_column, column)
			# 位置属性
			var r = "%s%s" % [ to_26_base_by_10_base( column ), row ]
			var column_node : ExcelXMLNode = column_to_node_dict.get(column) 
			if column_node == null:
				column_node = ExcelXMLNode.create("c", false, {
					PropertyName.COLUMN_ROW: r,
				})
				row_node.add_child(column_node)
			alter_value(workbook, column_node, column_data[column])
		
		row_node.set_attr("spans", "%s:%s" % [min_column, max_column])
		sheet_data_node.add_child(row_node)
