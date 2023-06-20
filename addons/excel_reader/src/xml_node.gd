#============================================================
#    Xml Node
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:52:09
# - version: 4.0
#============================================================
class_name ExcelXMLNode


var value = ""

var _type : String = "" # XML节点类型
var _parent: ExcelXMLNode 
var _children : Array[ExcelXMLNode] = []
var _attributes : Dictionary = {}
var _closure : bool = true
var _indent : int = 0


#============================================================
#  内置
#============================================================
func _init(xml_parser: XMLParser):
	self._type = xml_parser.get_node_name()
	for idx in xml_parser.get_attribute_count():
		var attr_name = xml_parser.get_attribute_name(idx)
		self._attributes[attr_name] = xml_parser.get_attribute_value(idx)


func _to_string():
	return "<%s#%s:type=%s>" % ["ExcelXMLNode", get_instance_id(), _type]


#============================================================
#  自定义
#============================================================
## XML格式化输出
func to_xml():
	# 参数
	var params_list = []
	for k in _attributes:
		params_list.append("%s=\"%s\"" % [k, _attributes[k]])
	var params_str = (" " + " ".join(params_list)) \
		if not params_list.is_empty() \
		else ""
	
	if not _closure:
		# 子节点
		var children_str = ""
		for child in _children:
			child._indent = _indent + 1
			children_str += "\n\t%s%s" % ["\t".repeat(_indent), child.to_xml()]
		
		# 缩进
		var indent_str = ""
		if _indent > 0 and children_str:
			indent_str = "\t".repeat(_indent)
		
		return "<{name}{params}>{_children}{indent}</{name}>".format({
			"name": _type,
			"indent": indent_str,
			"params": params_str,
			"_children": (children_str + "\n") if children_str else value
		})
		
	else:
		return "<{name}{params}/>".format({
			"name": _type,
			"params": params_str,
		})

func is_closure() -> bool:
	return _closure

func get_type() -> String:
	return _type

func get_parent() -> ExcelXMLNode:
	return _parent

func get_attr(property) -> String:
	return _attributes.get(property, "")

func has_attr(property) -> bool:
	return _attributes.has(property)

func get_attr_names() -> Array[String]:
	return Array(_attributes.keys(), TYPE_STRING, "", null)


func add_child(node: ExcelXMLNode) -> void:
	_children.append(node)
	node._parent = self


func get_children() -> Array[ExcelXMLNode]:
	return _children


func get_child(idx: int) -> ExcelXMLNode:
	if idx < _children.size():
		return _children[idx]
	return null


func get_child_count() -> int:
	return _children.size()


func get_first_node(path: String) -> ExcelXMLNode:
	var list = path.split("/")
	var node = find_first_node(list[0])
	if node and list.size() > 1:
		return node.get_first_node("/".join(list.slice(1)))
	return node


func find_first_node(_type: String) -> ExcelXMLNode:
	for child in get_children():
		if child._type == _type:
			return child
	return null

func find_nodes(_type: String) -> Array[ExcelXMLNode]:
	var list : Array[ExcelXMLNode] = []
	for child in get_children():
		if child._type == _type:
			list.append(child)
	return list

func get_value():
	return value


