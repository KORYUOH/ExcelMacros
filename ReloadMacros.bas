'================================================================
' vim:fenc=cp932:ft=vb
' Brief  : マクロ再読み込み
' Author : KORYUOH
' github : KORYUOH/ExcelMacros
' Create : 2017/12/14
' Update : 2018/01/14
' Version : 0.11
'================================================================
Attribute VB_Name = "ReloadMacros"
Option Explicit

Private Const ROOT As String = "MacroRoot"
Private Const IMPORT_MACRO_KEY As String = "_Import"

'-------------------------------------------
' 本体
'-------------------------------------------
Public Sub ReloadMacro()

	If Not SearchConfigSheet Then
		MsgBox "Not Found Config Sheet"
		Exit Sub
	End If

	Dim MacroRoot As String
	MacroRoot = GetKeyData( ROOT )

	Dim MacroFolder As String
	If Left(MacroRoot,1) = "." Then
		MacroFolder = GetAbsFilePath(MacroRoot)
	Else
		MacroFolder = GetAbsFilePath("", MacroRoot )
	End If

	if not Right( MacroFolder , 1 ) = "\" then
		MacroFolder = MacroFolder & "\"
	End If

	Dim DataNum As Integer

	DataNum = GetKeyDataNum( IMPORT_MACRO_KEY )
	if DataNum < 0 Then
		MsgBox "Not Found Key : " & IMPORT_MACRO_KEY
		Exit Sub
	End If

	Dim Itr As Integer
	Dim FileName AS String
	For Itr = 1 To DataNum
		FileName = GetKeyData( IMPORT_MACRO_KEY , DataOfs:=Itr )
		IncludeMacro MacroFolder & FileName
	Next Itr

End Sub


'-------------------------------------------
' モジュールをすべて開放
'-------------------------------------------
Private Sub ClearAllModules()
	Dim Component As Variant
	With ThisWorkbook.VBProject
		For Each Component In .VBComponents
			If Component.Type = 1 Then
				Component.Name = Component.Name & "OLD"
				.VBComponents.Remove Component
			End If
		Next Component
	End With
End Sub

'-------------------------------------------
' マクロをインポート
' マクロへのパス : FilePath
'-------------------------------------------
Private Sub IncludeMacro( FilePath As String )
	Dim Root As String
	Dim FileName As String
	If Left( FilePath , 1 ) = "." Then
		Root = ThisWorkbook.Path
		FileName = FilePath
	Else
		Dim pos As Long
		pos = InStrRev( FilePath , "\" )
		Root = Left( FilePath , pos - 1 )
		FileName = Mid( FilePath , pos + 1 )
	End If

	Dim Path As String
	Path = GetAbsFilePath( FileName , Root )
	
	ThisWorkbook.VBProject.VBComponents.Import Path

End Sub

