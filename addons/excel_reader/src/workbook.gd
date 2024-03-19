#============================================================
#    Workbook
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:52:01
# - version: 4.2.1
#============================================================
## 工作簿
##
##统一管理项目中的文件
class_name ExcelWorkbook


var file_path: String
var zip_reader : ZIPReader

# 文件路径对应的文件数据（所有文件的数据）
var _path_to_file_bytes_cache : Dictionary = {}
# 文件路径对应的 XML 文件
var _path_to_xml_file_cache : Dictionary = {}
# 发生了改变的文件路径
var _changed_file_path_dict : Dictionary = {}
# 路径对应的 ExcelSheet 对象
var _path_to_sheet_dict : Dictionary = {}
# 匹配表达式中的 ID 值
var _image_regex : RegEx = null:
	get:
		if _image_regex == null:
			_image_regex = RegEx.new()
			_image_regex.compile("DISPIMG\\(\"(?<rid>\\w+)\",\\d+\\)")
		return _image_regex


# 内容类型
var xl_content_types : ExcelContentTypes

# workbook 中的 rid 与 XML 文件关系
var xl_rels_workbook : ExcelXlRelsWorkbook
# 单元格中的 rid 与图片路径关系
var xl_rels_cell_images : ExcelXlRelsCellImages

# Sheet 的文件关系
var xl_workbook : ExcelXlWorkbook
# 单元格的图片数据关系
var xl_cell_images : ExcelXlCellImages
# 共享的字符串
var xl_shared_string : ExcelXlSharedStrings
# 单元格样式
var xl_styles : ExcelXlStyle



#============================================================
#  内置
#============================================================
func _init(zip_reader: ZIPReader):
	self.zip_reader = zip_reader
	
	# 文件的数据预先加载完
	for path in zip_reader.get_files():
		_path_to_file_bytes_cache[path] = zip_reader.read_file(path)
	
	# 先加载 rels 文件
	xl_rels_workbook = ExcelXlRelsWorkbook.new(self)
	xl_rels_cell_images = ExcelXlRelsCellImages.new(self)
	
	# 加载 xml 文件
	xl_workbook = ExcelXlWorkbook.new(self)
	xl_cell_images = ExcelXlCellImages.new(self)
	xl_shared_string = ExcelXlSharedStrings.new(self)
	xl_styles = ExcelXlStyle.new(self)
	xl_content_types = ExcelContentTypes.new(self)


func _to_string():
	return "<%s#%s>" % ["ExcelWorkbook", get_instance_id()]


#============================================================
#  自定义
#============================================================
## 获取所有文件数据。数据以 
##[codeblock]
##data[file_path] = PackedByteArray()
##[/codeblock]
##格式返回
func get_files_bytes() -> Dictionary:
	if not _changed_file_path_dict.is_empty():
		for path in _changed_file_path_dict:
			var xml_file = get_xml_file(path)
			update_file_data(path, ExcelDataUtil.get_xml_file_data(xml_file))
		print("有 %d 个文件发生了改变：" % _changed_file_path_dict.size(), _changed_file_path_dict.keys())
		_changed_file_path_dict.clear()
	return _path_to_file_bytes_cache


## 获取这个 XML 文件对象
func get_xml_file(path: String) -> ExcelXMLFile:
	if not _path_to_xml_file_cache.has(path):
		# 如果缓存中不存在，则从 ZIPReader 中读取数据
		_path_to_xml_file_cache[path] = ExcelXMLFile.new(self, path, read_file(path))
	return _path_to_xml_file_cache[path]


## 读取 XML 文件 bytes 数据
func read_file(path: String) -> PackedByteArray:
	if _changed_file_path_dict.has(path):
		_changed_file_path_dict.erase(path)
		# 如果这个文件发生了改变，则重新加载数据
		var xml_file = get_xml_file(path)
		update_file_data(xml_file, ExcelDataUtil.get_xml_file_data(xml_file))
	return _path_to_file_bytes_cache[path]

## 更新这个文件的数据
func update_file_data(path: String, data: PackedByteArray):
	_path_to_file_bytes_cache[path] = data

## 记录新建/修改过的文件路径
func add_changed_file(path: String):
	_changed_file_path_dict[path] = null

## 获取这个 Sheet 名称的 xml 文件路径
func get_path_by_sheet_name(sheet_name: String) -> String:
	var sheet_data = xl_workbook.get_sheet_data_by_name(sheet_name)
	var rid = sheet_data["r:id"]
	var path = xl_rels_workbook.get_path_by_id(rid)
	return path


func get_sheets() -> Array[ExcelSheet]:
	if _path_to_sheet_dict.is_empty():
		for data in xl_workbook._sheet_data_list:
			var r_id : String = data["r:id"]
			var xml_path : String = xl_rels_workbook.get_path_by_id(r_id)
			_path_to_sheet_dict[xml_path] = ExcelSheet.new(self, get_xml_file(xml_path))
	return Array(_path_to_sheet_dict.values(), TYPE_OBJECT, "RefCounted", ExcelSheet)


## 获取这个索引或这个名称的 Sheet
func get_sheet(idx_or_name) -> ExcelSheet:
	assert(idx_or_name is int or idx_or_name is String)
	
	# 如果 idx_or_name 为 name，则要注意大小写，因为这里区分大小写
	var xml_path : String = get_sheet_files()[idx_or_name] \
		if idx_or_name is int \
		else get_path_by_sheet_name(idx_or_name)
	if not xml_path.ends_with(".xml"):
		xml_path += ".xml"
	
	# 没有这个 sheet 路径
	if not get_sheet_files().has(xml_path):
		printerr("没有这个文件：", xml_path)
		return null
	
	# 还没加载这个数据则进行加载
	if not _path_to_sheet_dict.has(xml_path):
		_path_to_sheet_dict[xml_path] = ExcelSheet.new(self, get_xml_file(xml_path))
	
	return _path_to_sheet_dict[xml_path]


## 创建新的 Sheet
func create_new_sheet(sheet_name: String, data: Dictionary = {}) -> ExcelSheet:
	# 文件路径
	var sheet_id : int = xl_workbook.get_new_id()
	var xml_path : String = "xl/worksheets/sheet%d.xml" % sheet_id
	assert(not _path_to_sheet_dict.has(xml_path), "不能创建重复的xml")
	
	# workbook 记录这个文件
	var sheet_rid : String = xl_workbook.add_sheet(sheet_id, sheet_name)
	# 创建 Sheet 节点
	var sheet_xml_node = xl_workbook.create_sheet(sheet_id, sheet_rid, xml_path, sheet_name, data)
	# Workbook XML 文件关系
	xl_rels_workbook.add_relationship(ExcelDataUtil.FileType.WORKSHEET, sheet_rid, xml_path)
	
	# 记录发生改变的数据
	var sheet_file_bytes : PackedByteArray = ExcelDataUtil.get_xml_node_data(sheet_xml_node)
	var sheet_xml_file : ExcelXMLFile = ExcelXMLFile.new(self, xml_path, sheet_file_bytes)
	_path_to_xml_file_cache[xml_path] = sheet_xml_file
	add_changed_file(xml_path)
	
	# 内容类型
	xl_content_types.add_file(ExcelDataUtil.ContentType.WORKSHEET, xml_path)
	
	# ExcelSheet 对象
	var sheet = ExcelSheet.new(self, sheet_xml_file)
	_path_to_sheet_dict[xml_path] = sheet
	
	return sheet


## 更新共享的文字。返回这个字符串的索引
func update_shared_string_xml(text: String) -> int:
	return xl_shared_string.update_shared_string_xml(text)

## 获取共享字符串
func get_shared_string(idx: int) -> String:
	return xl_shared_string.get_shared_string(idx)


func get_sheet_files() -> Array[String]:
	return xl_rels_workbook.get_sheet_files()


## 表达式值转为图片。如果没有这张图片，则返回原数据
func convert_image(expression: String):
	# 嵌入单元格的图片表达式转为实际图片数据
	var result = _image_regex.search(expression)
	if result:
		var rid : String = result.get_string("rid")
		return xl_cell_images.get_image_by_id(rid)
	else:
		return expression


## 数值格式化
func format_value(cell_node: ExcelXMLNode):
	var value_node = cell_node.find_first_node("v")
	var value = float(value_node.get_value())
	#var apply_num_format = int(cell_node.get_attr(ExcelDataUtil.PropertyName.CELL_FORMAT))
	#var format_code = xl_styles.get_format_by_num(apply_num_format)
	#print("进行格式化：", " %-20s" % value, " %-4s" % format_idx, " ", format_code)
	# TODO 编写格式化码
	return value

