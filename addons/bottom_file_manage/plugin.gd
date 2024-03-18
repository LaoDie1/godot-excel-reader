@tool
extends EditorPlugin


var file_dock_parent : Control


func _enter_tree():
	await Engine.get_main_loop().process_frame
	var file_dock = get_editor_interface().get_file_system_dock()
	file_dock_parent = file_dock.get_parent_control()
	file_dock_parent.remove_child(file_dock)
	file_dock_parent.visible = (file_dock_parent.get_child_count() > 0)
	
	add_control_to_bottom_panel(file_dock, tr(file_dock.name))


func _exit_tree():
	var file_dock = get_editor_interface().get_file_system_dock()
	remove_control_from_bottom_panel(file_dock)
	file_dock_parent.add_child(file_dock)
	file_dock_parent.visible = true

