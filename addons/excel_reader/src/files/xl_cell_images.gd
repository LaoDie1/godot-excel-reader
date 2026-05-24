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


var _id_name_to_texture : Dictionary = {}
var _path_to_image: Dictionary = {} #需要这样进行记录，否则方法结束之后 Image 会释放掉

func _get_xl_path():
	return "xl/cellimages.xml"


func _init_data():
	# 数据
	if xml_file == null or xml_file.get_root() == null:
		return
	
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
		var image_path : String = workbook.xl_rels_cell_images.get_image_path_by_rid(rid)
		var image : Image = _image_loader(image_path)
		
		var src_rect = xdr_blip_fill.find_first_node_by_reg("a:srcRect")
		var xdr_spPr = xdr_pic.find_first_node("xdr:spPr")
		if xdr_spPr and src_rect:
			# 外框
			var a_xfrm = xdr_spPr.find_first_node("a:xfrm") # Transform2D 属性
			var a_off = a_xfrm.find_first_node("a:off") # 偏移坐标
			var a_ext = a_xfrm.find_first_node("a:ext") # 大小
			var offset = Vector2(int(a_off.get_attr("x", 0)), int(a_off.get_attr("y", 0)))
			var ext = Vector2(int(a_ext.get_attr("cx", 0)), int(a_ext.get_attr("cy", 0)))
			# 裁剪
			# OOXML 中 srcRect 的 l/t/r/b 属性使用的是 ST_Percentage 类型，它的定义是：
			#1 = 0.001% = 0.00001
			#100000 = 100% = 1.0
			var l = float(src_rect.get_attr("l")) / 100000 #左裁剪比例
			var t = float(src_rect.get_attr("t")) / 100000 #上裁剪
			var r = float(src_rect.get_attr("r")) / 100000 #右裁剪
			var b = float(src_rect.get_attr("b")) / 100000 #下裁剪
			var origin_size = Vector2(image.get_size())
			var clip_rect = Rect2( origin_size * Vector2(l, t), origin_size * Vector2(1.0 - l - r, 1 - t - b) )
			_id_name_to_texture[name] = ImageTexture.create_from_image( image.get_region(clip_rect) )
			#_id_name_to_texture[name] = ImageTexture.create_from_image( image )
		else:
			_id_name_to_texture[name] = ImageTexture.create_from_image( image )


func get_image_by_id(id: String) -> ImageTexture:
	return _id_name_to_texture.get(id)

func _image_loader(image_path:String) -> Image:
	if _path_to_image.has(image_path):
		return _path_to_image[image_path]
	var _buffer : PackedByteArray = workbook.read_file(image_path)
	var image = Image.new()
	var err : Error
	match image_path.get_extension().to_lower():
		"png": err = image.load_png_from_buffer(_buffer)
		"jpg","jpeg": err = image.load_jpg_from_buffer(_buffer)
		"svg": err = image.load_svg_from_buffer(_buffer)
		"bmp": err = image.load_bmp_from_buffer(_buffer)
		"tga": err = image.load_tga_from_buffer(_buffer)
		"ktx": err = image.load_ktx_from_buffer(_buffer)
		"webp": err = image.load_webp_from_buffer(_buffer)
		_: push_error("not supported image type:",image_path)
	if err != OK:
		push_error(err, " ", error_string(err))
	_path_to_image[image_path] = image
	return image

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
	#return emu * 96 / 914400
	return round(emu / 9525)
