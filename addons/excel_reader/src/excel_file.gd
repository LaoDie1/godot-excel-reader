#============================================================
#    Excel File
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:51:54
# - version: 4.2.1
#============================================================
## Excel 文件
##
##读取、保存文件数据
##[codeblock]
##var excel = ExcelFile.open_file( file_path )
##[/codeblock]
class_name ExcelFile


var file_path: String
var zip_reader : ZIPReader
var workbook: ExcelWorkbook


#============================================================
#  内置
#============================================================
func _notification(what):
	if what == NOTIFICATION_PREDELETE:
		if zip_reader:
			zip_reader.close()
			zip_reader = null


func _to_string():
	return "<%s#%s>" % ["ExcelFile", get_instance_id()]


#============================================================
#  自定义
#============================================================
static func open_file(path: String, auto_close: bool = false) -> ExcelFile:
	if FileAccess.file_exists(path):
		var excel_file = ExcelFile.new()
		excel_file.open(path)
		if auto_close:
			Engine.get_main_loop() \
				.process_frame \
				.connect(excel_file.close, Object.CONNECT_ONE_SHOT)
		return excel_file
	return null


func get_workbook() -> ExcelWorkbook:
	return workbook


func open(path: String) -> Error:
	self.file_path = path
	if zip_reader != null:
		close()
	zip_reader = ZIPReader.new()
	
	var err = zip_reader.open(path)
	if err != OK:
		printerr("Open failed: ", error_string(err))
		return err
	
	workbook = ExcelWorkbook.new(zip_reader)
	workbook.file_path = path
	
	return OK


func close() -> void:
	if zip_reader:
		zip_reader.close()
		zip_reader = null


## 保存文件。
func save(path: String = "") -> Error:
	if path == "":
		path = self.file_path
	
	assert(path != "")
	
	var writer := ZIPPacker.new()
	var err := writer.open(path)
	if err != OK:
		printerr("Save failed: ", error_string(err))
		return err
	
	# 写入数据
	var file_data_dict = workbook.get_files_bytes()
	for file in file_data_dict:
		writer.start_file(file)
		writer.write_file(file_data_dict[file])
	
	writer.close_file()
	writer.close()
	return OK


