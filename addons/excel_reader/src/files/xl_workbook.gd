#============================================================
#    Xl Workbook
#============================================================
# - author: zhangxuetu
# - datetime: 2024-03-18 17:21:11
# - version: 4.2.1
#============================================================
## 获取 Sheet 数据
class_name ExcelXlWorkbook
extends ExcelXlBase


# 每个 sheet 的数据
var _sheet_data_list : Array[Dictionary] = []


#@override
func _init_data():
	var root = xml_file.get_root()
	var sheets = root.find_first_node("sheets")
	for child in sheets.get_children():
		var xml_path : String = workbook.xl_rels_workbook.get_path_by_id(child.get_attr("r:id"))
		_sheet_data_list.append({
			"name": child.get_attr("name"),
			"sheetId": child.get_attr("sheetId"),
			"r:id": child.get_attr("r:id"),
		})


#@override
func _get_xl_path():
	return "xl/workbook.xml"

func get_sheet_data_list() -> Array[Dictionary]:
	return _sheet_data_list

func get_sheet_data_by_name(name: String) -> Dictionary:
	for data in _sheet_data_list:
		if data["name"] == name:
			return data
	return {}


## 添加 Sheet。返回这个 Sheet 的 rId
func add_sheet(sheet_name: String) -> String:
	var root = xml_file.get_root()
	var sheets = root.find_first_node("sheets")
	
	var max_id : int = 0
	for sheet_data in _sheet_data_list:
		if max_id < int(sheet_data["sheetId"]):
			max_id = int(sheet_data["sheetId"])
	
	var sheet_id : int = max_id + 1
	var r_id : String = "rId%d" % sheet_id
	var sheet = ExcelXMLNode.create("sheet", true, {
		"name": sheet_name,
		"sheetId": sheet_id,
		"r:id": r_id,
	})
	sheets.add_child(sheet)
	notify_change()
	
	return r_id


## 创建新的 Sheet
func create_sheet(id: String, xml_path: String, sheet_name: String, data: Dictionary = {}) -> ExcelXMLNode:
	# 根节点
	var worksheet : ExcelXMLNode = ExcelXMLNode.create("worksheet", false, {
		"xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
		"xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		"xmlns:xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
		"xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
		"xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
		"xmlns:etc": "http://www.wps.cn/officeDocument/2017/etCustomData",
	})
	
	# 表单分界线
	var sheet_pr = ExcelXMLNode.create("sheetPr", true)
	worksheet.add_child(sheet_pr)
	
	# 有效范围
	var dimension = ExcelXMLNode.create("dimension", true)
	dimension.set_attr("ref", "A1:A1")
	worksheet.add_child(dimension)
	
	# 选中范围
	var sheetViews = ExcelXMLNode.create("sheetViews", false)
	var sheetView = ExcelXMLNode.create("sheetView", false)
	sheetView.set_attr("tabSelected", 1)
	sheetView.set_attr("workbookViewId", 0)
	var selection = ExcelXMLNode.create("selection", true)
	selection.set_attr("activeCell", "A1")
	selection.set_attr("sqref", "A1")
	sheetView.add_child(selection)
	sheetViews.add_child(sheetView)
	worksheet.add_child(sheetViews)
	
	# 表单格式分界线
	var sheetFormatPr = ExcelXMLNode.create("sheetFormatPr", true)
	sheetFormatPr.set_attr("defaultColWidth", 8.88888888888889)
	sheetFormatPr.set_attr("defaultRowHeight", 14.4)
	worksheet.add_child(sheetFormatPr)
	
	# 列宽
	var cols = ExcelXMLNode.create("cols", false)
	worksheet.add_child(cols)
	
	# 表单数据
	var sheetData = ExcelXMLNode.create("sheetData", false)
	var row = ExcelXMLNode.create("row", false, {
		"r": 1,
		"spans": "1:1",
	})
	sheetData.add_child(row)
	worksheet.add_child(sheetData)
	
	# 添加这个 Sheet 的数据
	if not data.is_empty():
		ExcelDataUtil.add_node_by_data(workbook, sheetData, data)
	
	# 其他
	worksheet.add_child(ExcelXMLNode.create("pageMargins", true, {
		"left": "0.75",
		"right": "0.75",
		"top": "1",
		"bottom": "1",
		"header": "0.5",
		"footer": "0.5",
	}))
	worksheet.add_child(ExcelXMLNode.create("headerFooter", true))
	
	# workbook
	var sheets = get_xml_file().get_root().find_first_node("sheets")
	var sheet = ExcelXMLNode.create("sheet", true, {
		"name": sheet_name,
		"sheetId": 2,
		"r:id": id,
	})
	sheets.add_child(sheet)
	
	notify_change()
	
	return worksheet

