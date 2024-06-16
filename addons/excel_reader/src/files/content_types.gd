#============================================================
#    Content Types
#============================================================
# - author: zhangxuetu
# - datetime: 2024-03-19 13:45:01
# - version: 4.2.1
#============================================================
class_name ExcelContentTypes
extends ExcelXlBase


func _get_xl_path():
	return "[Content_Types].xml"


func _init_data():
	pass


func add_file(content_type: String, xml_path: String) -> void:
	var new_override = ExcelXMLNode.create("Override", true, {
		"PartName": xml_path,
		"ContentType": content_type,
	})
	xml_file.get_root().add_child(new_override)
	notify_change()
