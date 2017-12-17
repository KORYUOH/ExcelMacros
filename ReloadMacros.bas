'================================================================
' vim:fenc=cp932:ft=vb
' Brief : 
' Author : KORYUOH
' Create : 2017/12/14
' Update : 2017/12/18
' Version : 0.01
'================================================================
Attribute VB_Name = "ReloadMacros"
Option Explicit


'-------------------------------------------
' 本体
'-------------------------------------------
Public Sub ReloadMacro()
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
	
	ThisWorkbook.VBProject.Import Path

End Sub

