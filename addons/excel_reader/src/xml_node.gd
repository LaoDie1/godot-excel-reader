#============================================================
#    Xml Node
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:52:09
# - version: 4.2.1
#============================================================
class_name ExcelXMLNode


var value = ""

var _type : String = "" # XML节点类型
var _parent: ExcelXMLNode 
var _children : Array[ExcelXMLNode] = []
var _attributes : Dictionary = {}
var _closure : bool = true


#============================================================
#  内置
#============================================================
func _init(xml_parser: XMLParser = null):
	if xml_parser:
		self._type = xml_parser.get_node_name()
		for idx in xml_parser.get_attribute_count():
			var attr_name = xml_parser.get_attribute_name(idx)
			self._attributes[attr_name] = xml_parser.get_attribute_value(idx)


func _to_string():
	return "<%s#%s:type=%s>" % ["ExcelXMLNode", get_instance_id(), _type]


#============================================================
#  自定义
#============================================================
## 创建 XML 节点
##[br] - type: 节点名称类型
##[br] - closure: 是否是闭合的节点
##[br] - attributes: 节点属性
static func create(type: String, closure: bool, attributes: Dictionary = {}) -> ExcelXMLNode:
	var node = ExcelXMLNode.new()
	node._type = type
	node._closure = closure
	node._attributes = attributes
	return node


## 转为 xml 格式
func to_xml(indent: int = 0, format: bool = true) -> String:
	# 参数
	var params_list = []
	for k in _attributes:
		params_list.append('%s="%s"' % [k, _attributes[k]])
	var params_str = (" " + " ".join(params_list)) \
		if not params_list.is_empty() \
		else ""
	
	if not _closure or _children.size() > 0:
		# 子节点
		var children_str = ""
		if format:
			for child in _children:
				children_str += "\n\t%s%s" % [
					"\t".repeat(indent),
					child.to_xml(indent + 1, format),
				]
		else:
			for child in _children:
				children_str += child.to_xml(0, format)
		
		# 缩进
		var indent_str = ""
		if format and indent > 0 and children_str:
			indent_str = "\t".repeat(indent)
		
		return "<{name}{params}>{_children}{indent}</{name}>".format({
			"name": _type,
			"indent": indent_str,
			"params": params_str,
			"_children": (children_str + ("\n" if format else "") ) if children_str else value
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

func get_attr(property: String, default = ""):
	return _attributes.get(property, default)

func get_attr_map_by_props(propertys: Array) -> Dictionary:
	var dict : Dictionary = {}
	for property in propertys:
		dict[property] = _attributes.get(property, "")
	return dict

func get_attr_map() -> Dictionary:
	return _attributes

func has_attr(property) -> bool:
	return _attributes.has(property)

func get_attr_names() -> Array[String]:
	return Array(_attributes.keys(), TYPE_STRING, "", null)

func set_attr(property: String, value):
	_attributes[property] = value

func remove_attr(property: String):
	_attributes.erase(property)

func remove_all_child():
	_children.clear()

func remove_child(idx: int):
	_children.remove_at(idx)

func remove_node(node: ExcelXMLNode):
	for i in _children.size():
		if _children[i] == node:
			_children.remove_at(i)
			break

func remove_nodes(nodes: Array[ExcelXMLNode]):
	for i in range(_children.size()-1, -1, -1):
		if nodes.has(_children[i]):
			_children.remove_at(i)

func add_child(node: ExcelXMLNode) -> void:
	_children.append(node)
	node._parent = self
	_closure = false


## 添加到指定索引位置
func add_child_to(node: ExcelXMLNode, idx : int):
	_children.insert(idx, node)
	node._parent = self
	_closure = false


func get_children() -> Array[ExcelXMLNode]:
	return _children


func filter_child(callback: Callable) -> Array[ExcelXMLNode]:
	var list : Array[ExcelXMLNode] = []
	var node : ExcelXMLNode
	for child in _children:
		node = callback.call(child)
		if node:
			list.append(node)
	return list


func get_child(idx: int) -> ExcelXMLNode:
	if idx < _children.size():
		return _children[idx]
	return null

func get_child_count() -> int:
	return _children.size()

func get_value():
	return value

func get_full_value():
	var ret: String = value
	for child in get_children():
		ret += child.get_full_value()
	return ret

func find_first_node(type: String) -> ExcelXMLNode:
	for child in get_children():
		if child._type == type:
			return child
	return null

func find_nodes(type: String) -> Array[ExcelXMLNode]:
	var list : Array[ExcelXMLNode] = []
	for child in get_children():
		if child._type == type:
			list.append(child)
	return list


var _find_node_regex : RegEx = RegEx.new()

func find_first_node_by_path(path: String) -> ExcelXMLNode:
	var list = path.split("/")
	_find_node_regex.compile(list[0])
	for child in get_children():
		if _find_node_regex.search(child.get_type()):
			return child
	return null

##查找所有类型匹配的节点。示例：
##[codeblock]
###查找所有子节点下的所有 a: 开头类型的节点
##var nodes = find_nodes_by_path(./a:.)
##print(nodes)
##[/codeblock]
##以"/"进行切分每个层级的节点
func find_nodes_by_path(path: String) -> Array[ExcelXMLNode]:
	var all : Array[ExcelXMLNode] = []
	var last = [self]
	var items = path.split("/")
	for i in items.size():
		var current = []
		_find_node_regex.compile(items[i])
		for node in last:
			for child in node.get_children():
				if _find_node_regex.search(child.get_type()):
					current.append(child)
		all.append_array(current)
		last = current
	return all
