#============================================================
#    Workbook
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:52:01
# - version: 4.0
#============================================================
class_name ExcelWorkbook


const FilePaths = {
	WORKBOOK = "xl/workbook.xml",
	SHARED_STRINGS = "xl/sharedStrings.xml",
	CELL_IMAGES = "xl/cellimages.xml",
	STYLES = "xl/styles.xml",
	
	RELS_WORKBOOK = "xl/_rels/workbook.xml.rels",
	RELS_CELL_IMAGES = "xl/_rels/cellimages.xml.rels",
}


var file_path: String
var zip_reader : ZIPReader
var shared_strings : Array = []
var cellimages : Dictionary = {}: # name to image map
	get:
		if cellimages.is_empty():
			# 读取 rid 对应的图片
			var rels_cellimages_xml_data = get_xml_file(FilePaths.RELS_CELL_IMAGES)
			var rid_to_images = {}
			for child in rels_cellimages_xml_data.get_root().get_children():
				var rid = child.get_attr("Id")
				var path = child.get_attr("Target")
				rid_to_images[rid] = zip_reader.read_file("xl".path_join(path))
			
			# 读取数据引用
			var cellimages_xml_data = get_xml_file(FilePaths.CELL_IMAGES)
			for child in cellimages_xml_data.get_root().get_children():
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
		return cellimages

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

var _format_code = []

var _rels : ExcelXMLFile  # rid data
var _rid_to_path_dict : Dictionary = {}



#============================================================
#  内置
#============================================================
func _init(zip_reader: ZIPReader):
	self.zip_reader = zip_reader
	
	# Files
	for path in zip_reader.get_files():
		_path_to_file_bytes_cache[path] = zip_reader.read_file(path)
	
	self._rels = get_xml_file(FilePaths.RELS_WORKBOOK)
	
	# RID file path
	for child in _rels.get_root().get_children():
		var id = child.get_attr("Id")
		var target_path = "xl/" + child.get_attr("Target") # Files in the xl directory
		self._rid_to_path_dict[id] = target_path
	
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
func _create_sheet(xml_path: String) -> ExcelSheet:
	return ExcelSheet.new(self, xml_path)

func get_files_bytes() -> Dictionary:
	if not _path_to_changed_file_node_dict.is_empty():
		for path in _path_to_changed_file_node_dict:
			var xml_file = get_xml_file(path)
			var data = ExcelDataUtil.get_xml_file_data(xml_file)
			_path_to_file_bytes_cache[path] = data
		print("有 %d 个文件发生了改变：" % _path_to_changed_file_node_dict.size(), _path_to_changed_file_node_dict.keys())
		_path_to_changed_file_node_dict.clear()
	return _path_to_file_bytes_cache

func read_file(path: String) -> PackedByteArray:
	if _path_to_changed_file_node_dict.has(path):
		var xml_file = get_xml_file(path)
		var data = ExcelDataUtil.get_xml_file_data(xml_file)
		_path_to_file_bytes_cache[path] = data
		_path_to_changed_file_node_dict.erase(path)
		return data
	return _path_to_file_bytes_cache[path]

func get_xml_file(path: String) -> ExcelXMLFile:
	if not _path_to_xml_node_cache.has(path):
		_path_to_xml_node_cache[path] = ExcelXMLFile.new(self, path, zip_reader.read_file(path))
	return _path_to_xml_node_cache[path]

func get_sheet_files() -> Array[String]:
	return Array(_sheet_info_data_list.map(func(item): 
		return item["path"]
	), TYPE_STRING, &"", null)

func get_sheets() -> Array[ExcelSheet]:
	if _path_to_sheet_cache.is_empty():
		for data in _sheet_info_data_list:
			var xml_path = data["path"]
			_path_to_sheet_cache[xml_path] = _create_sheet(xml_path)
	return Array(_path_to_sheet_cache.values(), TYPE_OBJECT, "RefCounted", ExcelSheet)


## 获取这个 Sheet 名称的 xml 文件路径
func get_path_by_sheet_name(sheet_name: String) -> String:
	for data in _sheet_info_data_list:
		if data["sheet_name"] == sheet_name:
			return data["path"]
	return ""


func get_sheet(idx_or_name) -> ExcelSheet:
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
		_path_to_sheet_cache[xml_path] = _create_sheet(xml_path)
	
	return _path_to_sheet_cache[xml_path]


## 记录修改过的文件路径
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

func get_shared_string(idx: int) -> String:
	return shared_strings[idx]


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
	var format_idx = int(cell_node.get_attr(ExcelDataUtil.PropertyName.CELL_FORMAT))
	var format_code = _format_code[format_idx]
	print("进行格式化：", format_code)
	
	return value

