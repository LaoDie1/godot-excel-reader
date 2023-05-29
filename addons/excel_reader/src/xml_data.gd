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
		var root_node_offset : int = 0
		while parser.read() == OK:
			if parser.get_node_type() == XMLParser.NODE_ELEMENT:
				_root = ExcelXMLNode.new(parser)
				_root._closure = false
				root_node_offset = parser.get_node_offset()
				break
		
		# 节点是否闭合
		var closure_stack : Array[bool] = [false]
		var idx = 0
		while parser.read() == OK:
			match parser.get_node_type():
				XMLParser.NODE_ELEMENT, XMLParser.NODE_ELEMENT_END:
					# 上一个是否是闭合的节点
					closure_stack[idx] = (_source_code_buffer[parser.get_node_offset() - 2] == KEY_SLASH)
					# 当前节点信息
					closure_stack.push_back(false)
					idx += 1
		
		# 遍历
		parser.seek(root_node_offset)
		closure_stack.reverse()
		closure_stack.pop_back()
		_parse(parser, _root, closure_stack)


#============================================================
#  自定义
#============================================================
func _parse(parser: XMLParser, parent: ExcelXMLNode, closure_stack: Array[bool]) -> void:
	while parser.read() == OK:
		match parser.get_node_type():
			XMLParser.NODE_ELEMENT:
				# 当前节点信息
				var child = ExcelXMLNode.new(parser)
				child._closure = closure_stack.back()
				parent.add_child(child)
				
				# 出栈
				closure_stack.pop_back()
				
				# 不是闭合节点则遍历子节点
				if not child.is_closure():
					_parse(parser, child, closure_stack)
					# 如果结束的节点和当前父节点相同，则返回
					if parser.get_node_name() == parent.get_type():
						return
			
			XMLParser.NODE_ELEMENT_END:
				closure_stack.pop_back()
				parent._closure = false
				return
			
			XMLParser.NODE_TEXT:
				parent.value = parser.get_node_data()
			
#			XMLParser.NODE_COMMENT:
#				print('----NODE_COMMENT')
#			XMLParser.NODE_CDATA:
#				print('----NODE_CDATA')
#			XMLParser.NODE_UNKNOWN:
#				print('----NODE_UNKNOWN')
#			XMLParser.NODE_NONE:
#				print('----NODE_NONE')


func get_root() -> ExcelXMLNode:
	return _root as ExcelXMLNode

func get_source_code() -> String:
	return _source_code


