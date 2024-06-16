#============================================================
#    Rels Cell Images
#============================================================
# - author: zhangxuetu
# - datetime: 2024-03-18 23:12:13
# - version: 4.2.1
#============================================================
## 单元格图片关系。rid 对应的图片路径
class_name ExcelXlRelsCellImages
extends ExcelXlBase


var rid_to_image_path : Dictionary = {}
var max_id : int = 0


func _get_xl_path():
	return "xl/_rels/cellimages.xml.rels"


func _init_data():
	if xml_file == null or xml_file.get_root() == null:
		return
	
	for child in xml_file.get_root().get_children():
		var rid = child.get_attr("Id")
		var image_path = child.get_attr("Target")
		rid_to_image_path[rid] = "xl/".path_join(image_path) 
		
		var id = int(rid)
		if max_id < id:
			max_id = id


## 添加了图片
func add_image(image_path: String) -> String:
	if image_path.begins_with("xl/"):
		image_path = image_path.substr(3)
	elif image_path.begins_with("/xl/"):
		image_path = image_path.substr(4)
	else:
		assert(false)
	
	max_id += 1
	var rid = "rId%d" % max_id
	var relationship = ExcelXMLNode.create("Relationship", true, {
		"Id": max_id,
		"Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
		"Target": image_path,
	})
	xml_file.get_root().add_child(relationship)
	
	notify_change()
	
	return rid


## 获取这个 rid 的图片路径
func get_image_path_by_rid(rid:String) -> String:
	return rid_to_image_path[rid]
