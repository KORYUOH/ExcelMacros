'================================================================
' vim:fenc=cp932:ft=vb
' Brief : マクロ外部実行
' Author : KORYUOH
' github : KORYUOH/ExcelMacros
' Create : 2018/01/18
' Update : 2018/01/23
' Version : 0.80
'================================================================

'================================================================
' Settings
'================================================================
Dim MacroName 
Dim ExcelFileName
Dim Importer 
MacroName = "TempMacro"
ExcelFileName = "test.xlsx"
Importer = "ImportChecker.bas"
'================================================================
Dim Obj
Set Obj = WScript.CreateObject("Excel.Application")
Dim Path
Set Path = CreateObject("Scripting.FileSystemObject").GetFolder(".")
Obj.Visible = True
Obj.DisplayAlerts = False
CreateObject("WScript.Shell").AppActivate Obj.Caption
With Obj
	Dim book
	Set book = .Application.Workbooks.Open (Path & "\" & ExcelFileName)
	If book is Nothing Then
	Else
		With book
			.VBProject.VBComponents.Import Path &"\"& Importer
			Obj.Application.Run MacroName
			Obj.Application.Run ClearTmpModules
			.Save
			.Close
		End With
		.Application.Quit
	End If
End With

