@tool
extends EditorPlugin


func _enter_tree():
	# Initialization of the plugin goes here.
	pass
	
	print("example:")
	print('var excel = ExcelFile.open_file("C:\\Users\\z\\Desktop\\role_data.xlsx")')
	


func _exit_tree():
	# Clean-up of the plugin goes here.
	pass
