#============================================================
#    Xml File
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:52:06
# - version: 4.0
#============================================================
class_name ExcelXMLFile


var workbook: ExcelWorkbook
var xml_path: String

var _root : ExcelXMLNode
var _source_code : String = "":
	set(v):
		_source_code = v
		_source_code_buffer = v.to_utf8_buffer()
var _source_code_buffer : PackedByteArray


#============================================================
#  内置
#============================================================
func _to_string():
	return "<%s#%s>" % ["XMLData", get_instance_id()]


func _init(workbook: ExcelWorkbook, xml_path: String):
	self.workbook = workbook
	self.xml_path = xml_path
	
	var res := workbook.zip_reader.read_file(xml_path)
	var stack = []
	_source_code = PackedByteArray(res).get_string_from_utf8()
	var parser = XMLParser.new()
	if parser.open_buffer(_source_code_buffer) == OK:
		# 第一个节点
		while parser.read() == OK:
			if parser.get_node_type() == XMLParser.NODE_ELEMENT:
				_root = _parse(parser)
				break


#============================================================
#  自定义
#============================================================
func _parse(parser: XMLParser) -> ExcelXMLNode:
	var ret: ExcelXMLNode = ExcelXMLNode.new(parser)
	ret._closure = parser.is_empty()

	if not ret._closure:
		while parser.read() == OK:
			match parser.get_node_type():
				XMLParser.NODE_ELEMENT:
					var child: ExcelXMLNode = _parse(parser)
					ret.add_child(child)

				XMLParser.NODE_ELEMENT_END:
					if parser.get_node_name() != ret._type:
						push_warning("</%s> mismatch with <%s>" % [parser.get_node_name(), ret._type])
					return ret

				XMLParser.NODE_TEXT:
					ret.value = parser.get_node_data()

	return ret

func get_root() -> ExcelXMLNode:
	return _root as ExcelXMLNode

func get_source_code() -> String:
	return _source_code


## 保存数据
func save_as(path: String):
	var writer := ZIPPacker.new()
	var err := writer.open(path)
	if err != OK:
		return err
	
	# 其他数据
	var file_data_map = {}
	for file in workbook.zip_reader.get_files():
		if file != xml_path:
			var file_data = workbook.zip_reader.read_file(file)
			file_data_map[file] = file_data
	
	for file in file_data_map:
		writer.start_file(file)
		writer.write_file(file_data_map[file])
	
	# 新数据
	writer.start_file(xml_path)
	writer.write_file(
		get_root().to_xml().to_utf8_buffer()
	)
	
	writer.close_file()
	writer.close()
	
	return OK

