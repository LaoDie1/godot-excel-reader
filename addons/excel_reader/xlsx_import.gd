#============================================================
#    Xlsx Import
#============================================================
# - author: zhangxuetu
# - datetime: 2023-07-19 20:10:42
# - version: 4.0
#============================================================
@tool
extends EditorImportPlugin


func _get_importer_name():
	return "zhangxuetu.excelreader"

func _get_visible_name():
	return "Excel Reader"

func _get_recognized_extensions():
	return ["xlsx"]

func _get_save_extension():
	return "res"

func _get_resource_type():
	return "JSON"


#============================================================
#  预设
#============================================================
enum Presets {
	DEFAULT
}

func _get_preset_count():
	return 1

func _get_preset_name(preset_index):
	if preset_index == Presets.DEFAULT:
		return "Default"
	else:
		return "Unknown"

func _get_import_options(path, preset_index):
	if preset_index == Presets.DEFAULT:
		return [
			{name = "导入为数据格式", default_value = "json"},
		]
	else:
		return []


#============================================================
#  导入
#============================================================
func _import(source_file, save_path, options, platform_variants, gen_files):
	pass
	
	ExcelFile.open_file(source_file, true)
	


