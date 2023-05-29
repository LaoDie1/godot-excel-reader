#============================================================
#    Excel File
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:51:54
# - version: 4.0
#============================================================
class_name ExcelFile


var _zip_reader : ZIPReader
var _workbook: ExcelWorkbook


#============================================================
#  内置
#============================================================
func _notification(what):
	if what == NOTIFICATION_PREDELETE:
		if _zip_reader:
			_zip_reader.close()
			_zip_reader = null


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


func close() -> void:
	if _zip_reader:
		_zip_reader.close()
		_zip_reader = null


func open(path: String) -> Error:
	if _zip_reader != null:
		_zip_reader.close()
	_zip_reader = ZIPReader.new()
	
	var err = _zip_reader.open(path)
	if err != OK:
		print("Open failed: ", error_string(err))
		return err
	
	_workbook = ExcelWorkbook.new(_zip_reader)
	
	return OK


func get_workbook() -> ExcelWorkbook:
	return _workbook


