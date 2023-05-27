#============================================================
#    Excel File
#============================================================
# - author: zhangxuetu
# - datetime: 2023-05-27 21:51:54
# - version: 4.0
#============================================================
class_name ExcelFile


var zip_reader : ZIPReader
var workbook: ExcelWorkbook


#============================================================
#  内置
#============================================================
func _notification(what):
	if what == NOTIFICATION_PREDELETE:
		if zip_reader:
			zip_reader.close()


#============================================================
#  自定义
#============================================================
static func open_file(path: String, auto_close: bool = false) -> ExcelFile:
	if FileAccess.file_exists(path):
		var excel_reader = ExcelFile.new()
		excel_reader.open(path)
		if auto_close:
			Engine.get_main_loop().process_frame.connect(func():
				if excel_reader.zip_reader:
					excel_reader.zip_reader.close()
					excel_reader.zip_reader = null
			, Object.CONNECT_ONE_SHOT)
		
		return excel_reader
	return null


func close() -> void:
	if zip_reader:
		zip_reader.close()
		zip_reader = null


func open(path: String) -> Error:
	if zip_reader != null:
		zip_reader.close()
	zip_reader = ZIPReader.new()
	
	var err = zip_reader.open(path)
	if err != OK:
		print("Open failed: ", error_string(err))
		return err
	
	workbook = ExcelWorkbook.new(zip_reader, "xl/workbook.xml")
	
	return OK


func get_workbook() -> ExcelWorkbook:
	return workbook


