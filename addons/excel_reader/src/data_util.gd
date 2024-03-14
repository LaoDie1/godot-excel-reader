#============================================================
#    Data Util
#============================================================
# - author: zhangxuetu
# - datetime: 2024-03-12 14:28:17
# - version: 4.2
#============================================================
class_name ExcelDataUtil


## 数据类型
const DataType = {
	NUMBER = "", ## 没有类型的都为数值类型（待定）
	STRING = "s", ## 字符型
	EXPRESSION = "str", ## 表达式
}

## 属性名
const PropertyName = {
	COORD = "r", ## 所在行列坐标
	DATA_TYPE = "t", ## 数据类型
	CELL_FORMAT = "s", ## 单元格格式
}


## 获取数据类型
static func get_data_type(cell_node: ExcelXMLNode) -> String:
	return cell_node.get_attr(PropertyName.DATA_TYPE)


## 日期时间戳转为字符串类型
static func date_stamp_to_string(date_stamp: int) -> String:
	var days = ((date_stamp + 1 - 70 * 365 - 19) * 86400 - 8 * 3600)
	var datetime = Time.get_datetime_string_from_unix_time(days)
	return datetime.split("T")[0]


## 获取这个 xml 的 bytes 数据
static func get_xml_file_data(xml_file: ExcelXMLFile) -> PackedByteArray:
	# 不使用缩进
	return xml_file.get_root() \
		.to_xml(0, false) \
		.to_utf8_buffer()


static func set_value(cell_node: ExcelXMLNode, value):
	var v_node = cell_node.find_first_node("v")
	if v_node == null:
		v_node = ExcelXMLNode.create("v", false)
		cell_node.add_child(v_node)
	v_node.value = value


static func set_text(workbook: ExcelWorkbook, cell_node: ExcelXMLNode, value: String):
	cell_node.remove_attr(PropertyName.CELL_FORMAT) # 字符串没有单元格格式，移除掉这个属性
	cell_node.set_attr(PropertyName.DATA_TYPE, DataType.STRING)
	
	# 记录到字符串到缓存值列表
	var string_idx = workbook.update_shared_string_xml(value)
	set_value(cell_node, string_idx)


static func set_number(
	cell_node: ExcelXMLNode, 
	value: int,
	cell_format: int = -1,
):
	cell_node.remove_attr(PropertyName.DATA_TYPE) # 移除属性类型。默认为数字类型
	if cell_format > -1:
		cell_node.set_attr(PropertyName.CELL_FORMAT, cell_format)
	set_value(cell_node, value)


#static func set_date_by_string(cell_node: ExcelXMLNode, value: String):
	# TODO 字符串转为时间戳
	# var time_stamp
	#set_number(cell_node, time_stamp, CellFormat.DATE)


# FIXME 暂时遇到问题，未完成
static func set_image(
	workbook: ExcelWorkbook,
	cell_node: ExcelXMLNode, 
	image: Image, 
	descr: String = "",
):
	cell_node.remove_all_child()
	
	cell_node.remove_attr(PropertyName.CELL_FORMAT)
	cell_node.set_attr(PropertyName.DATA_TYPE, DataType.EXPRESSION)
	
	# 生成一个 IMAGE 的 ID
	var image_id = "ID_" % str(ResourceUID.create_id()).sha256_text().substr(32).to_upper()
	var cell_images_xml_file = workbook.get_xml_file(ExcelWorkbook.FilePaths.CELL_IMAGES)
	var etc_cell_images_node = cell_images_xml_file.get_root()
	var xdr_pic = etc_cell_images_node.find_first_node_by_path("etc:cellImage/xdr:pic")
	var nv_pic_pr = xdr_pic.find_first_node_by_path("xdr:nvPicPr")
	
	# 最大索引值
	var last_idx = 0
	for child in nv_pic_pr.find_nodes_by_path("xdr:cNvPr"):
		var id = int(child.get_attr("id"))
		if id > last_idx:
			last_idx = id
	
	# 添加图片
	if descr == "":
		descr = image_id
	var xdr_c_nv_pr = ExcelXMLNode.create("xdr:cNvPr", false)
	xdr_c_nv_pr.set_attr("id", last_idx + 1)
	xdr_c_nv_pr.set_attr("name", image_id)
	xdr_c_nv_pr.set_attr("descr", descr)
	nv_pic_pr.add_child(xdr_c_nv_pr)
	
	# 创建节点
	var f_node = ExcelXMLNode.create("f", false)
	f_node.value = '_xlfn.DISPIMG("%s",1)' % image_id
	cell_node.add_child(f_node)
	set_value(cell_node, 'DISPIMG("%s",1)' % image_id)


