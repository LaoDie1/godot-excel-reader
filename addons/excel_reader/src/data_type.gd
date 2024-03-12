#============================================================
#    Data Type
#============================================================
# - author: zhangxuetu
# - datetime: 2024-03-12 14:28:17
# - version: 4.0
#============================================================
class_name ExcelDataType


const STRING = "s" ## 字符值的单元格
const EXPRESSION = "str" ## 其他类型的表格


static func get_data_type(cell_node: ExcelXMLNode):
	var data_type = cell_node.get_attr("t")
	match data_type:
		"s":
			var s = cell_node.get_attr("s")
			return ("string" if s == "1" else "general")
		"str": 
			return "expression"
		_:
			match cell_node.get_attr("s"):
				"1": return "string"
				"2": return "number"
				"3": return "date"


static func set_value(cell_node: ExcelXMLNode, value, node_type: String = "v"):
	var v_node = cell_node.find_first_node(node_type)
	if v_node == null:
		v_node = ExcelXMLNode.create(node_type, false)
		cell_node.add_child(v_node)
	v_node.value = value


static func set_text(cell_node: ExcelXMLNode, value: String, string_cache: Array) -> bool:
	cell_node.set_attr("t", STRING)
	cell_node.set_attr("s", "1")
	
	# 记录到字符串到缓存值列表
	var string_idx = string_cache.find(value)
	if string_idx == -1:
		string_cache.append(value)
		value = str(string_cache.size() - 1)
		set_value(cell_node, value)
		return true
	else:
		set_value(cell_node, value)
		return false


static func set_number(cell_node: ExcelXMLNode, value: int):
	cell_node.remove_attr("t")
	cell_node.set_attr("s", "2")
	set_value(cell_node, value)


#static func set_date(cell_node: ExcelXMLNode, value: String):
	#cell_node.remove_attr("t")
	#cell_node.set_attr("s", "3")
	# TODO 字符串转为时间戳


# FIXME 暂时遇到问题，未完成
static func set_image(
	workbook: ExcelWorkbook,
	cell_node: ExcelXMLNode, 
	image: Image, 
	descr: String = "",
):
	cell_node.set_attr("t", EXPRESSION)
	cell_node.remove_attr("s")
	cell_node.remove_all_child()
	
	# 生成一个 IMAGE 的 ID
	var image_id = "ID_" % str(ResourceUID.create_id()).sha256_text().substr(32).to_upper()
	var cell_images_xml_file = workbook.get_xml_file("xl/cellimages.xml")
	var etc_cell_images_node = cell_images_xml_file.get_root()
	var xdr_pic = etc_cell_images_node.find_first_node_by_path("etc:cellImage/xdr:pic")
	var nv_pic_pr = xdr_pic.find_first_node_by_path("xdr:nvPicPr")
	
	# 最大的 ID
	var last_id = 0
	for child in nv_pic_pr.find_nodes_by_path("xdr:cNvPr"):
		var id = int(child.get_attr("id"))
		if id > last_id:
			last_id = id
	
	# 添加图片
	if descr == "":
		descr = image_id
	var xdr_c_nv_pr = ExcelXMLNode.create("xdr:cNvPr", false)
	xdr_c_nv_pr.set_attr("id", last_id + 1)
	xdr_c_nv_pr.set_attr("name", image_id)
	xdr_c_nv_pr.set_attr("descr", descr)
	nv_pic_pr.add_child(xdr_c_nv_pr)
	
	# 创建节点
	var f_node = ExcelXMLNode.create("f", false)
	f_node.value = '_xlfn.DISPIMG("%s",1)' % image_id
	cell_node.add_child(f_node)
	set_value(cell_node, 'DISPIMG("%s",1)' % image_id)
	
	# 保存文件
	var packer = ZIPPacker.new()
	var err = packer.open(workbook.file_path)
	if err != OK:
		return err
	packer.start_file("xl\\media\\" + descr.get_basename())
	packer.write_file( image.get_data() )
	packer.close_file()
	packer.close()
