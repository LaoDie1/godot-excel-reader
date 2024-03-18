#============================================================
#    Workbook
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:52:01
# - version: 4.0
#============================================================
## 工作簿
##
##统一管理项目中的文件数据
class_name ExcelWorkbook


const FilePaths = {
	CONTENT_TYPES = "[Content_Types].xml",
	
	RELS_WORKBOOK = "xl/_rels/workbook.xml.rels",
	RELS_CELL_IMAGES = "xl/_rels/cellimages.xml.rels",
	
	WORKBOOK = "xl/workbook.xml",
	SHARED_STRINGS = "xl/sharedStrings.xml",
	CELL_IMAGES = "xl/cellimages.xml",
	STYLES = "xl/styles.xml",
}

const DirPaths = {
	DOC_PROPS = "docProps",
	XL = "xl",
	
	WORKSHEETS = "xl/worksheets", ## 表单
	MEDIA = "xl/media", ## 媒体资源
	THEME = "xl/theme", ## 主题
	diagrams = "xl/diagrams", ## 形状（SmartArt）
	chartsheets = "xl/diagrams", ## 图表
	charts = "charts", ## 图表设置
	
}


var file_path: String
var zip_reader : ZIPReader
var shared_strings : Array = []
var cellimages : Dictionary = {} # name to image map

var _image_regex : RegEx = null:
	get:
		if _image_regex == null:
			_image_regex = RegEx.new()
			_image_regex.compile("DISPIMG\\(\"(?<rid>\\w+)\",\\d+\\)")
		return _image_regex
var _path_to_file_bytes_cache : Dictionary = {}
var _path_to_sheet_cache : Dictionary = {}
var _path_to_xml_node_cache : Dictionary = {}
var _path_to_changed_file_node_dict : Dictionary = {}
var _sheet_info_data_list : Array[Dictionary] = []

var _rels_workbook : ExcelXMLFile  # rid data file

var _format_code : Array = [] # 数字
var _rid_to_path_dict : Dictionary = {}



#============================================================
#  内置
#============================================================
func _init(zip_reader: ZIPReader):
	self.zip_reader = zip_reader
	
	# Files
	for path in zip_reader.get_files():
		_path_to_file_bytes_cache[path] = zip_reader.read_file(path)
	
	self._rels_workbook = get_xml_file(FilePaths.RELS_WORKBOOK)
	
	# RID file path
	for child in _rels_workbook.get_root().get_children():
		var id = child.get_attr("Id")
		var target_path = "xl/" + child.get_attr("Target") # Files in the xl directory
		self._rid_to_path_dict[id] = target_path
	
	# 读取 rid 对应的图片
	var rels_cellimages_xml_data = get_xml_file(FilePaths.RELS_CELL_IMAGES)
	var rid_to_images = {}
	for child in rels_cellimages_xml_data.get_root().get_children():
		var rid = child.get_attr("Id")
		var path = child.get_attr("Target")
		rid_to_images[rid] = zip_reader.read_file("xl".path_join(path))
	var cellimages_xml_file = get_xml_file(FilePaths.CELL_IMAGES)
	for child in cellimages_xml_file.get_root().get_children():
		# 名称
		var xdr_pic = child.get_child(0)
		var xdr_nv_pic_pr = xdr_pic.find_first_node("xdr:nvPicPr")
		var xdr_cNvPr = xdr_nv_pic_pr.find_first_node("xdr:cNvPr")
		var name = xdr_cNvPr.get_attr("name")
		# rid
		var xdr_blip_fill = xdr_pic.find_first_node("xdr:blipFill")
		var a_blip = xdr_blip_fill.find_first_node("a:blip")
		var rid = a_blip.get_attr("r:embed")
		# 记录
		var image = Image.new()
		image.load_png_from_buffer(PackedByteArray(rid_to_images[rid]))
		cellimages[name] = ImageTexture.create_from_image(image)
	
	# 数值格式化
	var xml_file = get_xml_file(FilePaths.STYLES)
	var num_fmts_node = xml_file.get_root().find_first_node("numFmts")
	for child in num_fmts_node.get_children():
		var format_code = child.get_attr("formatCode")
		_format_code.append(format_code)
	
	# Sheets 
	var sheets = get_xml_file(FilePaths.WORKBOOK) \
		.get_root() \
		.find_first_node("sheets")
	var rid : String
	var sheet_name : String
	for child in sheets.get_children():
		rid = child.get_attr("r:id")
		sheet_name = child.get_attr("name")
		_sheet_info_data_list.append({
			"rid": rid,
			"sheet_name": sheet_name,
			"path": _rid_to_path_dict[rid],
		})
	
	# 获取值列表，string 类型单元格数据的缓存
	var shared_strings_xml_node = get_xml_file(FilePaths.SHARED_STRINGS)
	for si_node in shared_strings_xml_node.get_root().get_children():
		shared_strings.append(si_node.get_full_value())
	


func _to_string():
	return "<%s#%s>" % ["Workbook", get_instance_id()]



#============================================================
#  自定义
#============================================================
## 获取所有文件数据。数据以 
##[codeblock]
##data[file_path] = PackedByteArray()
##[/codeblock]
##格式返回
func get_files_bytes() -> Dictionary:
	if not _path_to_changed_file_node_dict.is_empty():
		for path in _path_to_changed_file_node_dict:
			var xml_file = get_xml_file(path)
			var data = ExcelDataUtil.get_xml_file_data(xml_file)
			_path_to_file_bytes_cache[path] = data
		print("有 %d 个文件发生了改变：" % _path_to_changed_file_node_dict.size(), _path_to_changed_file_node_dict.keys())
		_path_to_changed_file_node_dict.clear()
	return _path_to_file_bytes_cache

## 获取这个 XML 文件对象
func get_xml_file(path: String) -> ExcelXMLFile:
	if not _path_to_xml_node_cache.has(path):
		# 如果缓存中不存在，则从 ZIPReader 中读取数据
		_path_to_xml_node_cache[path] = ExcelXMLFile.new(self, path, zip_reader.read_file(path))
	return _path_to_xml_node_cache[path]

## 读取 XML 文件 bytes 数据
func read_file(path: String) -> PackedByteArray:
	if _path_to_changed_file_node_dict.has(path):
		_path_to_changed_file_node_dict.erase(path)
		# 如果这个文件发生了改变，则重新加载数据
		var xml_file = get_xml_file(path)
		var data = ExcelDataUtil.get_xml_file_data(xml_file)
		_path_to_file_bytes_cache[path] = data
	return _path_to_file_bytes_cache[path]


func get_sheet_files() -> Array[String]:
	return Array(_sheet_info_data_list.map(func(item): 
		return item["path"]
	), TYPE_STRING, &"", null)

func get_sheets() -> Array[ExcelSheet]:
	if _path_to_sheet_cache.is_empty():
		for data in _sheet_info_data_list:
			var xml_path = data["path"]
			_path_to_sheet_cache[xml_path] = ExcelSheet.new(self, xml_path)
	return Array(_path_to_sheet_cache.values(), TYPE_OBJECT, "RefCounted", ExcelSheet)

func get_sheet_count() -> int:
	return _sheet_info_data_list.size()


## 获取这个 Sheet 名称的 xml 文件路径
func get_path_by_sheet_name(sheet_name: String) -> String:
	for data in _sheet_info_data_list:
		if data["sheet_name"] == sheet_name:
			return data["path"]
	return ""


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
	if not _path_to_sheet_cache.has(xml_path):
		_path_to_sheet_cache[xml_path] = ExcelSheet.new(self, xml_path)
	
	return _path_to_sheet_cache[xml_path]


## 记录新建/修改过的文件路径
func add_changed_file(path: String):
	_path_to_changed_file_node_dict[path] = null


## 更新共享的文字
func update_shared_string_xml(text: String) -> int:
	var idx = shared_strings.find(text)
	if idx > -1:
		return idx
	else:
		idx = shared_strings.size()
		shared_strings.append(text)
		
		var shared_string_xml_node = get_xml_file(FilePaths.SHARED_STRINGS)
		var si_node = ExcelXMLNode.create("si", false)
		var t_node = ExcelXMLNode.create("t", false)
		t_node.value = text
		si_node.add_child(t_node)
		var sst_node = shared_string_xml_node.get_root()
		sst_node.add_child(si_node)
		sst_node.set_attr("uniqueCount", shared_strings.size())
		sst_node.set_attr("count", int(sst_node.get_attr("count"))+1 )
		
		add_changed_file(FilePaths.SHARED_STRINGS)
		
		return idx


## 创建新的 Sheet
func create_sheet(sheet_name: String, data: Dictionary = {}) -> ExcelSheet:
	var xml_path = DirPaths.WORKSHEETS.path_join(sheet_name)
	if not xml_path.ends_with(".xml"):
		xml_path += ".xml"
	
	# 根节点
	var worksheet = ExcelXMLNode.create("worksheet", false, {
		"xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
		"xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		"xmlns:xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
		"xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
		"xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
		"xmlns:etc": "http://www.wps.cn/officeDocument/2017/etCustomData",
	})
	
	# 表单分界线
	var sheet_pr = ExcelXMLNode.create("sheetPr", true)
	worksheet.add_child(sheet_pr)
	
	# 有效范围
	var dimension = ExcelXMLNode.create("dimension", true)
	dimension.set_attr("ref", "A1:A1")
	worksheet.add_child(dimension)
	
	# 选中范围
	var sheetViews = ExcelXMLNode.create("sheetViews", false)
	var sheetView = ExcelXMLNode.create("sheetView", false)
	sheetView.set_attr("tabSelected", 1)
	sheetView.set_attr("workbookViewId", 0)
	var selection = ExcelXMLNode.create("selection", true)
	selection.set_attr("activeCell", "A1")
	selection.set_attr("sqref", "A1")
	sheetView.add_child(selection)
	sheetViews.add_child(sheetView)
	worksheet.add_child(sheetViews)
	
	# 表单格式分界线
	var sheetFormatPr = ExcelXMLNode.create("sheetFormatPr", true)
	sheetFormatPr.set_attr("defaultColWidth", 8.88888888888889)
	sheetFormatPr.set_attr("defaultRowHeight", 14.4)
	worksheet.add_child(sheetFormatPr)
	
	# 列宽
	var cols = ExcelXMLNode.create("cols", false)
	worksheet.add_child(cols)
	
	# 表单数据
	var sheetData = ExcelXMLNode.create("sheetData", false)
	var row = ExcelXMLNode.create("row", false, {
		"r": 4,
		"spans": "1:1",
	})
	sheetData.add_child(row)
	
	# 添加这个 Sheet 的数据
	ExcelDataUtil.add_node_by_data(self, worksheet, data)
	
	# 其他
	var pageMargins = ExcelXMLNode.create("pageMargins", true, {
		"left": "0.75",
		"right": "0.75",
		"top": "1",
		"bottom": "1",
		"header": "0.5",
		"footer": "0.5",
	})
	worksheet.add_child(pageMargins)
	
	var headerFooter = ExcelXMLNode.create("headerFooter", true)
	worksheet.add_child(headerFooter)
	
	# Workbook XML文件关系文件
	var rels_workbook = get_xml_file(FilePaths.RELS_WORKBOOK)
	var relationships = rels_workbook.get_root()
	var relationship = relationships.get_child(0)
	var last_id = int(relationship.get_attr("Id"))
	
	var new_relationship = ExcelXMLNode.create("Relationship", true, {
		"Id": "rId" + str(last_id + 1),
		"Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
		"Target": "worksheets/" + sheet_name + ".xml"
	})
	relationships.add_child_to(new_relationship, 0)
	
	# 内容类型
	var content_types = get_xml_file(FilePaths.CONTENT_TYPES)
	var types = content_types.get_root()
	var new_override = ExcelXMLNode.create("Override", true, {
		"PartName": xml_path,
		"ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
	})
	types.add_child(new_override)
	
	# workbook
	var workbook = get_xml_file(FilePaths.WORKBOOK)
	var sheets = workbook.get_root().find_first_node("sheets")
	# FIXME 下面的 ID 需要自增
	var sheet = ExcelXMLNode.create("sheet", true, {
		"name": sheet_name,
		"sheetId": 2,
		"r:id": 2,
	})
	sheets.add_child(sheet)
	
	
	# 记录发生改变的数据
	_path_to_xml_node_cache[xml_path] = ExcelXMLFile.new(self, xml_path, ExcelDataUtil.get_xml_node_data(worksheet))
	add_changed_file(xml_path)
	add_changed_file(FilePaths.RELS_WORKBOOK)
	add_changed_file(FilePaths.CONTENT_TYPES)
	add_changed_file(FilePaths.WORKBOOK)
	
	# TODO _rid_to_path_dict 等内容都需要更新数据
	_rels_workbook
	_rid_to_path_dict
	
	
	return get_sheet(sheet_name)


## 获取共享字符串
func get_shared_string(idx: int) -> String:
	return shared_strings[idx]


## 表达式值转为图片
func convert_image(value: String):
	# 嵌入单元格的图片表达式转为实际图片数据
	var result = _image_regex.search(value)
	if result:
		var rid = result.get_string("rid")
		return cellimages[rid]
	else:
		return value


## 数值格式化
func format_value(cell_node: ExcelXMLNode):
	var value_node = cell_node.find_first_node("v")
	var value = float(value_node.get_value())
	#var format_idx = int(cell_node.get_attr(ExcelDataUtil.PropertyName.CELL_FORMAT))
	#var format_code = _format_code[format_idx - 2]
	#print("进行格式化：", " %-20s" % value, " %-4s" % format_idx, " ", format_code)
	return value

