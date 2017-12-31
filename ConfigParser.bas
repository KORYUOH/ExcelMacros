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

End Function

'-------------------------------------------
' キーデータを探す
'-------------------------------------------
Function GetKeyData( Key As String ,Optional KeyCollum As Integer = 1 , Optional DataOfs As Integer = 1) As String
	If Not SearchConfigSheet Then
		GetKeyData = ""
		Exit Function
	End If

	Dim KeyRow AS Integer
	KeyRow = GetKeyRow( Key , KeyCollum )
	If KeyRow < 0 Then
		GetKeyData = ""
		Exit Function
	End If

	With Config
		GetKeyData = .Cells( KeyRow , KeyCollum + DataOfs )
	End With

End Function

Function GetKeyRow(Key As String , Optional KeyCollum AS Integer = 1) As Integer

	If Not SearchConfigSheet Then
		GetKeyRow = -1
		Exit Function
	End If

	Dim MaxRow AS Integer
	MaxRow = GetMaxRow( Config , KeyCollum )
	Dim Itr AS Integer
	For Itr = 1 To MaxRow
		With Config
			If .Cells( Itr , KeyCollum ) = Key Then
				GetKeyRow = Itr
				Exit Function
			End If
		End With
	Next Itr
End Function

Function GetKeyDataNum( Key As String , Optional KeyCollum As Integer = 1) As Integer

	If Not SearchConfigSheet Then
		GetKeyDataNum = -1
		Exit Function
	End If

	Dim KeyRow AS Integer
	KeyRow = GetKeyRow( Key , KeyCollum )

	If KeyRow < 0 Then
		GetKeyDataNum = -1
		Exit Function
	End If

	With Config
		GetKeyDataNum = .Cells( KeyRow , .Columns.Count ).End(xlLeft).Row
	End With

End Function




