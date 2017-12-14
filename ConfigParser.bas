'================================================================
' vim:fenc=cp932:ft=vb
' Brief : コンフィグシートパーサー
' Author : KORYUOH
' Create : 2017/12/14
' Update : 2017/12/14
' Version : 0.01
'================================================================
Attribute VB_Name = "ConfigParser"
Option Explicit

Private Config As WorkSheet

'-------------------------------------------
' コンフィグシートを探す
' あればやらない
' 無い場合Falseを返す
'-------------------------------------------
Function SearchConfigSheet() As Boolean

	SearchConfigSheet = False

	If Not Config Is Nothing Then
		SearchConfigSheet = True
		Exit Function
	End If

	Dim sheet As WorkSheet

	For Each sheet In ThisWorkbook.WorkSheets
		If sheet.Cells(1,1).Value = "Config" Then
			Set Config = sheet
			SearchConfigSheet = True
			Exit Function
		End If
	Next sheet

End If

'-------------------------------------------
' キーデータを探す
'-------------------------------------------
Function GetKeyData( Key As String ,Optional KeyCollum As Long = 1 , Optional DataOfs As Long = 1) As String
	If Not SearchConfigSheet Then
		GetKeyData = ""
		Exit Function
	End If

	Dim MaxRow As Integer
	MaxRow = GetMaxRow( Config , KeyCollum )
	Dim Itr As Integer
	For Itr = 1 To MaxRow
		With Config
			If .Cells( Itr , KeyCollum) = Key Then
				GetKeyData = .Cells( Itr , KeyCollum + DataOfs )
				Exit For
			End If
		End With
	Next Itr
End Function


