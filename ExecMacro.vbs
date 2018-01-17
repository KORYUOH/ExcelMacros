'================================================================
' vim:fenc=cp932:ft=vb
' Brief : マクロ外部実行
' Author : KORYUOH
' github : KORYUOH/ExcelMacros
' Create : 2018/01/18
' Update : 2018/01/18
' Version : 0.01
'================================================================

'================================================================
' Settings
'================================================================
Dim MacroName = ""
Dim ExcelFileName = ""
Dim Importer = ""
'================================================================
Dim Obj
Set Obj = WScript.CreateObject("Excel.Application")
Dim Path
Set Path = CreateObject("Scripting.FileSystemObject").GetFolder(".")
Obj.Visible = True
CreateObject("WScript.Shell").AppActivate Obj.Caption
Dim Workbook
Set Workbook = Obj.Workbooks.Open Path & "\" & ExcelFileName
Workbook.VBProject.VBComponents.Import Importer
Obj.Application.Run MacroName
Obj.Application.Run ClearAllModules


