#============================================================
#    Xl File Base
#============================================================
# - author: zhangxuetu
# - datetime: 2024-03-18 17:22:17
# - version: 4.0
#============================================================
class_name ExcelXlFileBase


var workbook : ExcelWorkbook
var xml_file : ExcelXMLFile


func _get_xl_file_path() -> String:
	assert(false, "没有重写 _get_xl_file_path 方法")
	return ""


#============================================================
#  内置
#============================================================
func _init(workbook: ExcelWorkbook):
	self.workbook = workbook
	self.xml_file = workbook.get_xml_file(_get_xl_file_path())
