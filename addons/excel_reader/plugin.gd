@tool
extends EditorPlugin


const OriginDataImport = preload("uid://3pplq4gjy8et")
const FormatDataImport = preload("uid://cplsmrwurvm65")

var import_plugins: Array[EditorImportPlugin]

func _enter_tree():
	for import_plugin in [
		OriginDataImport.new(),
		FormatDataImport.new(),
	]:
		add_import_plugin(import_plugin)
		import_plugins.append(import_plugin)


func _exit_tree():
	for import_plugin in import_plugins:
		remove_import_plugin(import_plugin)
	import_plugins = []
