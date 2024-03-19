#============================================================
#    Xl File Base
#============================================================
# - author: zhangxuetu
# - datetime: 2024-03-18 17:22:17
# - version: 4.2.1
#============================================================
class_name ExcelXlBase


var workbook : ExcelWorkbook
var xml_file : ExcelXMLFile


#============================================================
#  内置
#============================================================
func _init(workbook: ExcelWorkbook, xml_path: String = ""):
	self.workbook = workbook
	if xml_path == "":
		self.xml_file = workbook.get_xml_file( _get_xl_path() )
	else:
		self.xml_file = workbook.get_xml_file(xml_path)
	_init_data()



#============================================================
#  自定义
#============================================================
func _init_data():
	pass

func _get_xl_path() -> String:
	assert(false, "没有重写 _get_xl_path 方法")
	return ""

func get_xml_file() -> ExcelXMLFile:
	return xml_file

func to_xml(format: bool = false):
	return xml_file.to_xml(format)

## 通知发生了改变
func notify_change() -> void:
	workbook.add_changed_file( _get_xl_path() )


