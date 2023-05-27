#============================================================
#    Xml Data
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:52:06
# - version: 4.0
#============================================================
class_name ExcelXMLData


var root : ExcelXMLNode
var source_code : String = "":
	set(v):
		assert(v != "", "Source code cannot be set to empty")
		source_code = v
		code_buffer = v.to_utf8_buffer()
var code_buffer : PackedByteArray


#============================================================
#  内置
#============================================================
func _to_string():
	return "<%s:%s>" % ["XMLData", get_instance_id()]


func _init(zip_reader: ZIPReader, xml_path: String):
	var res := zip_reader.read_file(xml_path)
	var stack = []
	source_code = PackedByteArray(res).get_string_from_utf8()
	source_code = source_code.replace("\n", " ")
	var parser = XMLParser.new()
	if parser.open_buffer(code_buffer) == OK:
		while parser.read() == OK:
			if parser.get_node_type() == XMLParser.NODE_ELEMENT:
				root = ExcelXMLNode.new(parser)
				_parse(parser, root)
				root.closure = false


func _get(property):
	if property.is_valid_int():
		return root.get_child(int(property))


#============================================================
#  自定义
#============================================================
func _parse(parser: XMLParser, parent: ExcelXMLNode) -> void:
	# 只有一个独立节点的闭合节点
	const CLOSURE_TYPE = [
		# workbook
		"fileVersion", "calcPr", "workbookView", "sheet", "workbookPr", 
		
		# sheet1
		"sheetPr", "dimension", "selection", "sheetFormatPr", "col",
		"pageMargins", "headerFooter", 
	]
	
	while parser.read() == OK:
		match parser.get_node_type():
			XMLParser.NODE_ELEMENT:
				var child = ExcelXMLNode.new(parser)
				child.closure = (child.get_type() in CLOSURE_TYPE)
				parent.add_child(child)
				
				# 不是闭合节点
				if not child.closure:
					_parse(parser, child)
					
					# 如果结束的节点和当前父节点相同，则返回
					if parser.get_node_name() == parent.get_type():
						return
			
			XMLParser.NODE_ELEMENT_END:
				parent.closure = false
				return
			
			XMLParser.NODE_TEXT:
				parent.value = parser.get_node_data()
			
#			XMLParser.NODE_COMMENT:
#				print('----NODE_COMMENT')
#			XMLParser.NODE_CDATA:
#				print('----NODE_CDATA')
			XMLParser.NODE_UNKNOWN:
				print('----NODE_UNKNOWN')
				print(parser.get_node_name())
#			XMLParser.NODE_NONE:
#				print('----NODE_NONE')


func get_root() -> ExcelXMLNode:
	return root as ExcelXMLNode

func get_source_code() -> String:
	return source_code


