#============================================================
#    Xl Cellimages
#============================================================
# - author: zhangxuetu
# - datetime: 2024-03-18 23:06:30
# - version: 4.2.1
#============================================================
## 单元格中的图片数据
class_name ExcelXlCellImages
extends ExcelXlBase


var id_name_to_texture_dict : Dictionary = {}
var data_list : Array = []


func _get_xl_path():
	return "xl/cellimages.xml"


func _init_data():
	# 数据
	for etc_cell_image in xml_file.get_root().get_children():
		var xdr_pic = etc_cell_image.find_first_node("xdr:pic")
		var xdr_nv_pic_pr = xdr_pic.find_first_node("xdr:nvPicPr")
		var xdr_c_nv_pr = xdr_nv_pic_pr.find_first_node("xdr:cNvPr")
		
		var xdr_blip_fill = xdr_pic.find_first_node("xdr:blipFill")
		var a_blip = xdr_blip_fill.find_first_node("a:blip")
		
		data_list.append({
			id = xdr_c_nv_pr.get_attr("id"),
			name = xdr_c_nv_pr.get_attr("name"),
			descr = xdr_c_nv_pr.get_attr("descr"),
			rid = a_blip.get_attr("r:embed")
		}) 
	
	# 加载图片数据
	var _image_loader = func(image_path:String, _buffer:PackedByteArray):
		var image = Image.new()
		var extension = image_path.get_extension().to_lower()
		match extension:
			"png":image.load_png_from_buffer(_buffer)
			"jpg","jpeg":image.load_jpg_from_buffer(_buffer)
			"svg":image.load_svg_from_buffer(_buffer)
			"bmp":image.load_bmp_from_buffer(_buffer)
			"tga":image.load_tga_from_buffer(_buffer)
			"ktx":image.load_ktx_from_buffer(_buffer)
			"webp":image.load_webp_from_buffer(_buffer)
			_: push_error("not supported image type:",image_path)
		if not image.is_empty():
			return ImageTexture.create_from_image(image)
	
		
	for data in data_list:
		var image_path = workbook.xl_rels_cell_images.get_image_path_by_rid(data["rid"])
		id_name_to_texture_dict[ data["name"] ] = _image_loader.call(image_path, workbook.read_file(image_path))
	
	# 读取 rid 对应的图片
	for child in get_xml_file().get_root().get_children():
		# 名称
		var xdr_pic = child.get_child(0)
		var xdr_nv_pic_pr = xdr_pic.find_first_node("xdr:nvPicPr")
		var xdr_cNvPr = xdr_nv_pic_pr.find_first_node("xdr:cNvPr")
		var name = xdr_cNvPr.get_attr("name")
		# rid
		var xdr_blip_fill = xdr_pic.find_first_node("xdr:blipFill")
		var a_blip = xdr_blip_fill.find_first_node("a:blip")
		var rid = a_blip.get_attr("r:embed")
		# 记录
		var image_path = workbook.xl_rels_cell_images.get_image_path_by_rid(rid)
		id_name_to_texture_dict[name] = _image_loader.call(image_path, workbook.read_file(image_path))


func get_image_by_id(id: String) -> ImageTexture:
	return id_name_to_texture_dict.get(id)


