'================================================================
' vim:fenc=cp932:ft=vb
' Brief : ��{�}�N���t�@�C�������C���|�[�g
' Author : KORYUOH
' github : KORYUOH/ExcelMacros
' Create : 2018/07/31
' Update : 2018/07/31
' Version : 0.10
'================================================================

'================================================================
' Settings
'================================================================
Option Explicit

Dim ExcelApp
Set ExcelApp = WScript.CreateObject("Excel.Application")
Dim Modules
Modules = Array( "CommonMacroLib", "ConfigParser", "ReloadMacros" )
Dim ModuleFolder

Dim PathArray
Set PathArray = WScript.Arguments
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

If PathArray.Count = 0 Then
	WScript.Echo "Excel�t�@�C����DD���邩�����Ɏw�肵�Ď��s���Ă�������"
	WScript.Quit
End If

ModuleFolder = FSO.GetFolder(".")
' �G�N�Z����\�����邩
ExcelApp.Visible = True
' �A���[�g(�x����\�����邩)
ExcelApp.DisplayAlerts = False

Dim ItrPath
For Each ItrPath In PathArray
	Dim FName
	FName = FSO.GetFileName(ItrPath)
	Dim FExt
	FExt = FSO.GetExtensionName(ItrPath)
	If InStr(UCASE(FExt) , "XLS" ) Then
		With ExcelApp
			Dim book
			Set book = .Application.Workbooks.Open (ItrPath)
			If Not book is Nothing Then
				With book
					Dim Module 
					For Each Module In Modules
						CheckImport Module , book
						Dim ModName
						ModName = ModuleFolder & "\" & Module & ".bas"
						.VBProject.VBComponents.Import (ModName)
					Next 
					CheckConfigSheet( book )
					.Save
					.Close
				End With
			End If
			.Application.Quit
		End With
	End If
Next

Sub CheckImport( MacroName , book )
	Dim Components
	Set Components = book.VBProject.VBComponents
	Dim Component
	For Each Component In Components
		If Component.Type = 1 Then
			If Component.Name = MacroName Then
				Component.Name = Component.Name & "OLD"
				book.VBProject.VBComponents.Remove Component
				Exit Sub
			End If
		End if
	Next 
End Sub

Sub CheckConfigSheet( Book )
	If Not ExcelApp.Application.Run ("SearchConfigSheet") Then
		With Book
			Dim Sheet
			Set Sheet = .Worksheets.Add ( , .Worksheets( .Worksheets.Count ))
			If Sheet Is Nothing Then
				Exit Sub
			End If
			With Sheet
				.Name = "Config"
				.Range("A1") = "Config"
				.Range("A3") = "MacroRoot"
				.Range("B3") = ModuleFolder
				.Range("A4") = "_Import"
				Dim ModuleItr
				For ModuleItr = 0 To UBOUND( Modules )
					.Range(.Cells(4, 2 + ModuleItr) , .Cells(4,2 + ModuleItr) ) = Modules(ModuleItr) & ".bas"
				Next
			End With
		End With
	End If
End Sub
