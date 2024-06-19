#============================================================
#    Xl Cellimages
#============================================================
# - author: zhangxuetu
# - datetime: 2024-03-18 23:06:30
# - version: 4.2.1
#============================================================
## 单元格中的图片数据
##
## - 参考：https://blog.csdn.net/renfufei/article/details/77481753
## - 参考：https://learn.microsoft.com/zh-cn/dotnet/api/documentformat.openxml.drawing.spreadsheet.blipfill?view=openxml-3.0.1
## - 参考：https://learn.microsoft.com/zh-cn/dotnet/api/documentformat.openxml.drawing.spreadsheet.picture?view=openxml-3.0.1
class_name ExcelXlCellImages
extends ExcelXlBase


var id_name_to_texture_dict : Dictionary = {}
var data_list : Array = []


func _get_xl_path():
	return "xl/cellimages.xml"


func _init_data():
	# 数据
	if xml_file == null or xml_file.get_root() == null:
		return
	
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
		match image_path.get_extension().to_lower():
			"png":image.load_png_from_buffer(_buffer)
			"jpg","jpeg":image.load_jpg_from_buffer(_buffer)
			"svg":image.load_svg_from_buffer(_buffer)
			"bmp":image.load_bmp_from_buffer(_buffer)
			"tga":image.load_tga_from_buffer(_buffer)
			"ktx":image.load_ktx_from_buffer(_buffer)
			"webp":image.load_webp_from_buffer(_buffer)
			_: push_error("not supported image type:",image_path)
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
		
		var src_rect = xdr_blip_fill.find_first_node_by_path("a:srcRect")
		var xdr_spPr = xdr_pic.find_first_node("xdr:spPr")
		if src_rect and xdr_spPr:
			var a_xfrm = xdr_spPr.find_first_node("a:xfrm") # Transform2D 属性
			var a_off = a_xfrm.find_first_node("a:off") # 偏移
			var a_ext = a_xfrm.find_first_node("a:ext") # 大小
			var offset = Vector2(int(a_off.get_attr("x", 0)), int(a_off.get_attr("y", 0)))
			var ext = Vector2(int(a_ext.get_attr("cx", 0)), int(a_ext.get_attr("cy", 0)))
			var rect = Rect2(offset, ext)
			var image_rect = Rect2(emu_to_px(rect.position), emu_to_px(rect.size))
			printt(rect, image_rect)
			
			#var left = int(src_rect.get_attr("l", 0))
			#var right = int(src_rect.get_attr("r", 0))
			#var top = int(src_rect.get_attr("t", 0))
			#var bottom = int(src_rect.get_attr("b", 0))
			#var rect = Rect2i(left, top, right, bottom)
			#var image_rect = Rect2(emu_to_px(rect.position), emu_to_px(rect.size))
			
			var texture_image = _image_loader.call(image_path, workbook.read_file(image_path)) as ImageTexture
			id_name_to_texture_dict[name] = ImageTexture.create_from_image( texture_image.get_image().get_region(image_rect) )
			continue
			
		id_name_to_texture_dict[name] = _image_loader.call(image_path, workbook.read_file(image_path))


func get_image_by_id(id: String) -> ImageTexture:
	return id_name_to_texture_dict.get(id)


#============================================================
#  单位换算
#============================================================
#1 in=914400 EMUs，1 cm=360000 EMUs
#
# - 参考：https://blog.csdn.net/oy538730875/article/details/84687585
# - 参考：https://blog.csdn.net/MooreLxr/article/details/120859899
# - 参考：https://blog.csdn.net/qq_24127015/article/details/119608686

## 毫米转厘米
static func mm_to_cm(mm):
	return mm / 10.0;

## 厘米转英寸
static func cm_to_inch(cm):
	return cm / 2.54;

## 英寸转磅
static func inch_to_pt(inch):
	return inch * 72.0;

## 磅转缇
static func pt_to_dxa(pt):
	return pt * 20.0;

## 缇转英寸
static func dxa_to_inch(dxa):
	return dxa_to_points(dxa) / 72.0;

## 缇转像素点
static func dxa_to_points(dxa):
	return dxa / 20.0;

## 缇转EMUs
static func dxa_to_emu(dxa):
	return 914400.0 * dxa_to_inch(dxa);

## EMUs转缇
static func emu_to_dxa(emu):
	return pt_to_dxa(inch_to_pt(emu)) / 914400.0;

## 缇转像素
static func dxa_to_px(dxa):
	return dxa / 15.0

static func dxa_to_cm(dxa):
	return dxa / 567.0


static func emu_to_px(emu):
	return emu / 9525.0
