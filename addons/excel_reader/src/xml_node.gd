#============================================================
#    Xml Node
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:52:09
# - version: 4.0
#============================================================
class_name ExcelXMLNode


var type : String = "":  # XML节点类型
	set(v):
		if type == "":
			type = v
		else:
			assert(false)
		
var value = ""
var _parent: ExcelXMLNode 
var children : Array[ExcelXMLNode] = []
var attributes : Dictionary = {}
var closure : bool = true
var _indent : int = 0


#============================================================
#  内置
#============================================================
func _init(xml_parser: XMLParser):
	self.type = xml_parser.get_node_name()
	for idx in xml_parser.get_attribute_count():
		var attr_name = xml_parser.get_attribute_name(idx)
		self.attributes[attr_name] = xml_parser.get_attribute_value(idx)


func _to_string():
	return "<ExcelXMLNode:name=%s>" % [type]


#============================================================
#  自定义
#============================================================
# XML格式化输出
func to_xml():
	# 参数
	var params_list = []
	for k in attributes:
		params_list.append("%s=\"%s\"" % [k, attributes[k]])
	var params_str = (" " + " ".join(params_list)) \
		if not params_list.is_empty() \
		else ""
	
	if not closure:
		# 子节点
		var children_str = ""
		for child in children:
			child._indent = _indent + 1
			children_str += "\n\t%s%s" % ["\t".repeat(_indent), child.to_xml()]
		
		# 缩进
		var indent_str = ""
		if _indent > 0 and children_str:
			indent_str = "\t".repeat(_indent)
		
		return "<{name}{params}>{children}{indent}</{name}>".format({
			"name": type,
			"indent": indent_str,
			"params": params_str,
			"children": (children_str + "\n") if children_str else value
		})
		
	else:
		return "<{name}{params}/>".format({
			"name": type,
			"params": params_str,
		})


func get_type() -> String:
	return type

func get_parent() -> ExcelXMLNode:
	return _parent

func get_attr(property) -> String:
	return attributes.get(property, "")

func has_attr(property) -> bool:
	return attributes.has(property)

func get_attr_names() -> Array[String]:
	return Array(attributes.keys(), TYPE_STRING, "", null)


func add_child(node: ExcelXMLNode) -> void:
	children.append(node)
	node._parent = self


func get_children() -> Array[ExcelXMLNode]:
	return children


func get_child(idx: int) -> ExcelXMLNode:
	if idx < children.size():
		return children[idx]
	return null


func get_child_count() -> int:
	return children.size()


func get_first_node(path: String) -> ExcelXMLNode:
	var list = path.split("/")
	var node = find_first_node(list[0])
	if node and list.size() > 1:
		return node.get_first_node("/".join(list.slice(1)))
	return node


func find_first_node(type: String) -> ExcelXMLNode:
	for child in get_children():
		if child.type == type:
			return child
	return null

func find_nodes(type: String) -> Array[ExcelXMLNode]:
	var list : Array[ExcelXMLNode] = []
	for child in get_children():
		if child.type == type:
			list.append(child)
	return list

func get_value():
	return value


