![Plugin Logo](icon.svg)

# Godot Excel Reader

[![Godot Engine 4.2.1](https://img.shields.io/badge/Godot%20Engine-4.0.2-blue)](https://godotengine.org/)
[![MIT license](https://img.shields.io/badge/license-MIT-blue.svg)](https://lbesson.mit-license.org/)

Reading excel files. 

**During the writing function test, there may be risks associated with its use**.

**写入功能试验中，使用可能会有风险！**


---



## Example

Load xlsx file data

```gdscript
#var excel_data = load(xlsx_path) as ExcelFileData
var excel_data = ExcelFileData.load_file(xlsx_path)
var table_data = excel_data.get_sheet_data("Sheet1")
print(JSON.stringify(table_data, "\t"))
```


> Read source file data:
>
> ```gdscript
> var excel = ExcelFile.open_file("xlsx file path")
> var workbook = excel.get_workbook()
> 
> var sheet = workbook.get_sheet(0)
> # Or use the following line, where two lines of code are equivalent
> #var sheet = workbook.get_sheet("sheet1") as ExcelSheet
> var table_data = sheet.get_table_data()
> print(JSON.stringify(table_data, "\t"))
> 
> # Output by row and column
> var table_data = sheet.get_table_data()
> for row in table_data:
> 	var column_data = table_data[row]
> 	for column in column_data:
> 		print(column_data[column])
> ```
>



## Contribute

Any contributions is welcome! If you find any bugs, please report in `issues`.
