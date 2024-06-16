@tool
extends EditorPlugin


const XlsxImport = preload("res://addons/excel_reader/import/xlsx_import.gd")

var import_plugin: EditorImportPlugin

func _enter_tree():
	import_plugin = XlsxImport.new()
	add_import_plugin(import_plugin)

func _exit_tree():
	remove_import_plugin(import_plugin)
	import_plugin = null
