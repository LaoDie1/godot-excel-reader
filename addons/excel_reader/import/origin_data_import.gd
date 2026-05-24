#============================================================
#    Origin Data Import
#============================================================
# - author: zhangxuetu
# - datetime: 2023-07-19 20:10:42
# - version: 4.2.1
#============================================================
@tool
extends EditorImportPlugin


#============================================================
#  导入
#============================================================
func _get_importer_name():
	return "zhangxuetu.ExcelReader.origin"

func _get_visible_name():
	return "Origin Data"

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
		file_to_data_dict[name] = sheet.get_table_data()
		sheet_name_list.append(name)
	
	var excel_data = ExcelFileData.new()
	excel_data.sheet_name_list = sheet_name_list
	excel_data.data = file_to_data_dict
	var path = "%s.%s" % [ save_path, _get_save_extension() ]
	return ResourceSaver.save(excel_data, path)


#============================================================
#  预设
#============================================================
func _get_import_options(path, preset_index):
	return Array([], TYPE_DICTIONARY, "", null)

func _get_option_visibility(path: String, option_name: StringName, options: Dictionary) -> bool:
	return true

func _get_priority() -> float:
	return 5.0

func _get_import_order() -> int:
	return ResourceImporter.IMPORT_ORDER_DEFAULT
