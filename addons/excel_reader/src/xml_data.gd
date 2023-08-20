#============================================================
#    Xml Data
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:52:06
# - version: 4.0
#============================================================
class_name ExcelXMLData


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


func _init(zip_reader: ZIPReader, xml_path: String):
	var res := zip_reader.read_file(xml_path)
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


