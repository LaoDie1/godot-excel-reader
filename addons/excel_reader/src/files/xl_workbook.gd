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
		_record_sheet_data(child)


#@override
func _get_xl_path():
	return "xl/workbook.xml"


func _record_sheet_data(sheet_node: ExcelXMLNode) -> void:
	#var xml_path : String = workbook.xl_rels_workbook.get_path_by_id(child.get_attr("r:id"))
	_sheet_data_list.append({
		"name": sheet_node.get_attr("name"),
		"sheetId": sheet_node.get_attr("sheetId"),
		"r:id": sheet_node.get_attr("r:id"),
	})


func get_sheet_data_list() -> Array[Dictionary]:
	return _sheet_data_list

func get_sheet_data_by_name(name: String) -> Dictionary:
	for data in _sheet_data_list:
		if data["name"] == name:
			return data
	return {}

func get_new_id() -> int:
	var max_id : int = 0
	for sheet_data in _sheet_data_list:
		if max_id < int(sheet_data["sheetId"]):
			max_id = int(sheet_data["sheetId"])
	return max_id + 1


## 添加 Sheet。返回这个 Sheet 的 rId
func add_sheet(sheet_id: int, sheet_name: String) -> String:
	var r_id : String = "rId%d" % sheet_id
	var sheet = ExcelXMLNode.create("sheet", true, {
		"name": sheet_name,
		"sheetId": sheet_id,
		"r:id": r_id,
	})
	var root = xml_file.get_root()
	var sheets = root.find_first_node("sheets")
	sheets.add_child(sheet)
	
	_record_sheet_data(sheet)
	
	notify_change()
	return r_id


## 创建新的 Sheet
func create_sheet(sheet_id: int, sheet_rid: String, xml_path: String, sheet_name: String, data: Dictionary = {}) -> ExcelXMLNode:
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
	
	# 获取非空数据
	if not data.is_empty():
		var new_data : Dictionary = {}
		for row in data:
			if not data[row].is_empty():
				new_data[row] = data[row]
		data = new_data
	
	# 有效范围
	var min_coords : Vector2i = Vector2i(1,1)
	if not data.is_empty():
		var first_row = data.keys()[0]
		min_coords.y = first_row
		min_coords.x = data[first_row].values()[0]
	var max_coords : Vector2i = Vector2i(1,1)
	for row in data:
		max_coords.y = max(max_coords.y, row)
		min_coords.y = min(min_coords.y, row)
		for column in data[row]:
			max_coords.x = max(max_coords.x, column)
			min_coords.x = min(min_coords.x, column)
	var dimension : ExcelXMLNode = ExcelXMLNode.create("dimension", true)
	var ref : String = ExcelDataUtil.to_dimension(Rect2i( min_coords, max_coords - min_coords))
	dimension.set_attr("ref", ref)
	worksheet.add_child(dimension)
	
	# 可见区域选中范围
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
	var sheetFormatPr = ExcelXMLNode.create("sheetFormatPr", true, {
		"defaultColWidth": 9,
		"defaultRowHeight": 13.5,
		"outlineLevelCol": 1,
	})
	worksheet.add_child(sheetFormatPr)
	
	# 表单数据
	var sheetData = ExcelXMLNode.create("sheetData", false)
	worksheet.add_child(sheetData)
	if not data.is_empty():
		ExcelDataUtil.add_node_by_data(workbook, sheetData, data)
	else:
		var row = ExcelXMLNode.create("row", false, {
			"r": 1,
			"spans": "1:1",
		})
		sheetData.add_child(row)
	
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
	
	notify_change()
	
	return worksheet
