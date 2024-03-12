#============================================================
#    Workbook
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:52:01
# - version: 4.0
#============================================================
class_name ExcelWorkbook


const PATH_XML_SHEET = "xl/worksheets/"


var file_path: String
var zip_reader : ZIPReader
var _workbook_xml_file : ExcelXMLFile:
	set(v):
		assert(_workbook_xml_file == null)
		_workbook_xml_file = v
var shared_strings : Array = []
var cellimages : Dictionary = {}: # name to image map
	get:
		if cellimages.is_empty():
			# 读取 rid 对应的图片
			var rels_cellimages_xml_data = get_xml_file("xl/_rels/cellimages.xml.rels")
			var rid_to_images = {}
			for child in rels_cellimages_xml_data.get_root().get_children():
				var rid = child.get_attr("Id")
				var path = child.get_attr("Target")
				rid_to_images[rid] = zip_reader.read_file("xl".path_join(path))
			
			# 读取数据引用
			var cellimages_xml_data = get_xml_file("xl/cellimages.xml")
			for child in cellimages_xml_data.get_root().get_children():
				# 名称
				var xdr_pic = child.get_child(0)
				var xdr_nv_pic_pr = xdr_pic.find_first_child_node("xdr:nvPicPr")
				var xdr_cNvPr = xdr_nv_pic_pr.find_first_child_node("xdr:cNvPr")
				var name = xdr_cNvPr.get_attr("name")
				# rid
				var xdr_blip_fill = xdr_pic.find_first_child_node("xdr:blipFill")
				var a_blip = xdr_blip_fill.find_first_child_node("a:blip")
				var rid = a_blip.get_attr("r:embed")
				# 记录
				var image = Image.new()
				image.load_png_from_buffer(PackedByteArray(rid_to_images[rid]))
				cellimages[name] = ImageTexture.create_from_image(image)
		return cellimages

var _sheets : Dictionary = {}
var _sheet_files : Array[String] = []
var _string_value_cache : Array = []
var _sheet_data_list : Array[Dictionary] = []

var _rels : ExcelXMLFile  # rid data
var _rid_to_path_map : Dictionary = {}


#============================================================
#  内置
#============================================================
func _init(zip_reader: ZIPReader):
	self.zip_reader = zip_reader
	self._workbook_xml_file = get_xml_file("xl/workbook.xml")
	self._rels = get_xml_file("xl/_rels/workbook.xml.rels")
	
	# RID file path
	for child in _rels.get_root().get_children():
		var id = child.get_attr("Id")
		var target_path = "xl/" + child.get_attr("Target") # Files in the xl directory
		self._rid_to_path_map[id] = target_path
	
	# Sheets 
	var sheets = _workbook_xml_file.get_root().find_first_node("sheets")
	var rid : String
	var sheet_name : String
	for child in sheets.get_children():
		rid = child.get_attr("r:id")
		sheet_name = child.get_attr("name")
		_sheet_data_list.append({
			"rid": rid,
			"sheet_name": sheet_name,
			"path": _rid_to_path_map[rid],
		})
	
	# 表单文件列表
	for file in zip_reader.get_files():
		if file.begins_with(PATH_XML_SHEET) and file.ends_with(".xml"):
			_sheet_files.append(file)
	self._sheet_files.erase(PATH_XML_SHEET)
	
	# 获取值列表，string 类型单元格数据的缓存
	var sharedStrings = get_xml_file("xl/sharedStrings.xml")
	for si_node in sharedStrings.get_root().get_children():
		shared_strings.append(si_node.get_full_value())


func _to_string():
	return "<%s#%s>" % ["Workbook", get_instance_id()]



#============================================================
#  自定义
#============================================================
func _get_sheet(xml_path: String) -> ExcelSheet:
	return ExcelSheet.new(self, xml_path)

func get_xml_file(path: String) -> ExcelXMLFile:
	return ExcelXMLFile.new(self, path)

func get_sheet_files() -> Array[String]:
	return Array(_sheet_data_list.map(func(item): 
		return item["path"]
	), TYPE_STRING, &"", null)

func get_sheet_name_list() -> Array[String]:
	var sheets = _workbook_xml_file \
		.find_first_node("workbook/_sheets") \
		.get_children() \
		.map( func(item: ExcelXMLNode): return item.get_attr("name") )
	return Array(sheets, TYPE_STRING, "", null)


func get_sheets() -> Array[ExcelSheet]:
	if _sheets.is_empty():
		for xml_path in _sheet_files:
			_sheets[xml_path] = _get_sheet(xml_path)
	return Array(_sheets.values(), TYPE_OBJECT, "RefCounted", ExcelSheet)


func get_path_by_sheet_name(sheet_name: String) -> String:
	for data in _sheet_data_list:
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
	if not _sheets.has(xml_path):
		_sheets[xml_path] = _get_sheet(xml_path)
	
	return _sheets[xml_path]

