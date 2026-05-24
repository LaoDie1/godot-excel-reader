#============================================================
#    Xml Node
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:52:09
# - version: 4.2.1
#============================================================
# 已修改为双向链表形式的结构。方便快速插入数据
class_name ExcelXMLNode

var value : String = ""

var _type : String = "" # XML节点类型
var _attributes : Dictionary = {}
var _closure : bool = true
var _parent: ExcelXMLNode 
var _previous: ExcelXMLNode
var _next: ExcelXMLNode
var _child_first: ExcelXMLNode
var _child_last: ExcelXMLNode


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
##[br]
##[br][code]indent[/code]  缩进字符数
##[br][code]format[/code]  xml格式化输出
func to_xml(indent: int = 0, format: bool = false) -> String:
	# 参数
	var params_list = []
	for k in _attributes:
		params_list.append('%s="%s"' % [k, _attributes[k]])
	var params_str = (" " + " ".join(params_list)) \
		if not params_list.is_empty() \
		else ""
	
	var children = get_children()
	if not _closure or children.size() > 0:
		# 子节点
		var children_str = ""
		if format:
			for child in children:
				children_str += "\n\t%s%s" % [
					"\t".repeat(indent),
					child.to_xml(indent + 1, format),
				]
		else:
			for child in children:
				children_str += child.to_xml(0, format)
		
		# 缩进
		var indent_str = ""
		if format and indent > 0 and children_str:
			indent_str = "\t".repeat(indent) #缩进字符
		
		return "<{name}{params}>{children}{indent}</{name}>".format({
			"name": _type,
			"indent": indent_str,
			"params": params_str,
			"children": (children_str + ("\n" if format else "") ) if children_str else value
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

## 获取这个XML节点字典格式的属性数据
func get_attr_map() -> Dictionary:
	return _attributes

func has_attr(property) -> bool:
	return _attributes.has(property)

func get_attr_names() -> Array[String]:
	return Array(_attributes.keys(), TYPE_STRING, "", null)

func set_attr(property: String, _value):
	_attributes[property] = _value

func remove_attr(property: String):
	_attributes.erase(property)

func get_value():
	return value

func get_full_value():
	var ret: String = value
	for child in get_children():
		ret += child.get_full_value()
	return ret

## 插入到当前节点之前
func insert_before(node: ExcelXMLNode) -> void:
	var tmp : ExcelXMLNode = self._previous
	self._previous = node
	node._next = self
	if tmp != null:
		tmp._next = node
		node._previous = tmp
	if _parent and _parent._child_first == self: #如果当前节点就是第一个节点，则修改第一个节点的对象
		_parent._child_first = node

## 插入到这个节点之后
func insert_after(node: ExcelXMLNode) -> void:
	var tmp : ExcelXMLNode = self._next
	self._next = node
	node._previous = self
	if tmp != null:
		node._next = tmp
		tmp._previous = node
	if _parent and _parent._child_last == self: #如果当前节点就是最后一个节点，则修改最后一个节点的对象
		_parent._child_last = node

## 创建子节点
func create_child(type: String, closure: bool, attributes: Dictionary = {}) -> ExcelXMLNode:
	var node : ExcelXMLNode = ExcelXMLNode.new()
	node._type = type
	node._closure = closure
	node._attributes = attributes
	add_node(node)
	return node

## 添加节点
func add_node(node: ExcelXMLNode) -> void:
	node._parent = self
	_closure = false
	if _child_last != null:
		_child_last._next = node
		node._previous = _child_last
	else:
		_child_first = node
	_child_last = node

## 添加到指定索引位置
func add_node_to(node: ExcelXMLNode, idx : int) -> void:
	var to_node = get_child(idx)
	to_node.insert_after(node)
	node._parent = self
	_closure = false

## 过滤筛选子节点
func filter_child(callback: Callable) -> Array[ExcelXMLNode]:
	var list : Array[ExcelXMLNode] = []
	var node : ExcelXMLNode = _child_first
	while node:
		if callback.call(node):
			list.append(node)
		node = node._next
	return list

## 移除所有字节点
func remove_all_child() -> void:
	#var tree = Engine.get_main_loop()
	#if tree is SceneTree:
		#var node : ExcelXMLNode = _child_first
		#while node != null:
			#tree.queue_delete(node)
			#node = node._next
	_child_first = null
	_child_last = null

func get_children() -> Array[ExcelXMLNode]:
	var list : Array[ExcelXMLNode] = []
	var node : ExcelXMLNode = _child_first
	while node != null:
		list.push_back(node)
		node = node._next
	return list

func get_child(idx: int) -> ExcelXMLNode:
	var index : int = 0
	var node : ExcelXMLNode = _child_first
	while index < idx and node != null:
		node = node._next
		index += 1
	return node


func find_first_node(type: String) -> ExcelXMLNode:
	var node : ExcelXMLNode = _child_first
	while node:
		if node._type == type:
			return node
		node = node._next
	return null

func find_nodes(type: String) -> Array[ExcelXMLNode]:
	var list : Array[ExcelXMLNode] = []
	var node : ExcelXMLNode = _child_first
	while node:
		if node._type == type:
			list.append(node)
		node = node._next
	return list


var _find_node_regex : RegEx = RegEx.new()

##查找所有类型匹配的节点。示例：
##[codeblock]
###查找所有子节点下的所有 a: 开头类型的节点
##var nodes = ExcelXMLNode.find_nodes_by_path("./a:.")
##print(nodes)
##[/codeblock]
##以"/"进行切分每个层级的节点
func find_nodes_by_reg(path: String) -> Array[ExcelXMLNode]:
	var all : Array[ExcelXMLNode] = []
	var last = [self]
	var items = path.split("/")
	for i in items.size():
		var currents : Array = []
		_find_node_regex.compile(items[i])
		for node:ExcelXMLNode in last:
			var child: ExcelXMLNode = node
			while child:
				if _find_node_regex.search(child.get_type()):
					currents.append(child)
				child = child._next
		all.append_array(currents)
		last = currents
	return all

func find_first_node_by_reg(path: String) -> ExcelXMLNode:
	var list : PackedStringArray = path.split("/")
	_find_node_regex.compile(list[0])
	var child: ExcelXMLNode = _child_first
	while child:
		if _find_node_regex.search(child.get_type()):
			return child
		child = child._next
	return null
