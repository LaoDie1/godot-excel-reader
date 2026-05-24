#============================================================
#    Format Data Import
#============================================================
# - author: zhangxuetu
# - datetime: 2026-05-24 17:21:52
# - version: 4.7.0.beta3
#============================================================
@tool
extends EditorImportPlugin


#============================================================
#  导入
#============================================================
func _get_importer_name():
	return "zhangxuetu.ExcelReader.format"

func _get_visible_name():
	return "Format Data"

func _get_recognized_extensions():
	return ["xlsx"]

func _get_save_extension():
	return "res"

func _get_resource_type():
	return "Resource"

func _import(
	source_file: String, 
	save_path: String, 
	options: Dictionary, 
	platform_variants: Array[String], 
	gen_files: Array[String]
) -> Error:
	var excel = ExcelFile.open_file(source_file, true)
	if excel == null:
		push_error("打开 xlsx 失败，这个文件可能正在被编辑")
		return FAILED
	var workbook = excel.get_workbook()
	var file_to_data_dict : Dictionary = {}
	var sheet_name_list : Array[String] = []
	for name in workbook.get_sheet_name_list():
		var sheet = workbook.get_sheet(name)
		file_to_data_dict[name] = get_format_array(sheet.get_table_data(), name, options["head_line"], options["ignore_empty_data"])
		sheet_name_list.append(name)
	
	var excel_data = ExcelFileData.new()
	excel_data.sheet_name_list = sheet_name_list
	excel_data.data = file_to_data_dict
	var path = "%s.%s" % [ save_path, _get_save_extension() ]
	return ResourceSaver.save(excel_data, path)


## 以 [head_line] 行中的数据作为 [code]key[/code] 记录到数据列表中方便使用 
func get_format_array(sheet_data: Dictionary, sheet_name: String, head_line: int, ignore_empty_data: bool) -> Array[Dictionary]:
	var head_data : Dictionary = sheet_data[head_line]
	var new_data: Array[Dictionary] = []
	var item : Dictionary
	for line in sheet_data:
		if line != head_line:
			item = {}
			if ignore_empty_data:
				for column in sheet_data[line]:
					item[ head_data.get(column) ] = sheet_data[line][column]
			else:
				for column in head_data:
					item[ head_data[column] ] = sheet_data[line].get(column)
			new_data.append(item)
	return new_data


#============================================================
#  预设
#============================================================
func _get_import_options(path, preset_index):
	return [
		{
			"name": "head_line",
			"default_value": 1,
			"property_hint": PROPERTY_HINT_RANGE,
			"hint_string": "1,1,1,or_greater,hide_control"
		},
		{
			"name": "ignore_empty_data",
			"default_value": true,
		}
	]

func _get_option_visibility(path: String, option_name: StringName, options: Dictionary) -> bool:
	return true

func _get_import_order() -> int:
	return ResourceImporter.IMPORT_ORDER_DEFAULT
