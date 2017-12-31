'================================================================
' vim:fenc=cp932:ft=vb
' Brief : 汎用マクロモジュール
' Author : KORYUOH
' Create : 2017/12/14
' Update : 2017/12/22
' Version : 0.01
'================================================================
Attribute VB_Name = "CommonMacroLib"
Option Explicit

'-------------------------------------------
' バッチ起動のコマンドプロンプトの表示状態
'-------------------------------------------
Public Const CMD_HIDE As Integer = 0    ' 非表示
Public Const CMD_NORMAL As Integer = 1  ' 通常
Public Const CMD_MINIMUM As Integer = 2 ' 最小化
Public Const CMD_MAXIMUM As Integer = 3 ' 最大化


'-------------------------------------------
' シートが有るか
' 検索ブック : book
' 検索シート名 : sheetName
'-------------------------------------------
Public Function IsExistSheet( book As Workbook , sheetName As String ) As Boolean
	Dim i As Integer
	IsExistSheet = False
	
	With book
		For i = 1 To .Sheets.Count
			If .Sheets(i).Name = sheetName Then
				IsExistSheet = True
				Exit For
			End If
		Next i
	End With
End Function

'-------------------------------------------
' シートの最終データ行を取得する
' 対象シート : Sheet
' 対象列 : Columns [省略可能]
'-------------------------------------------
Public Function GetMaxRow( Sheet As WorkSheet , Optional Columns As Integer = 1 ) As Integer

	If Sheet Is Nothing Then
		GetMaxRow = -1
		Exit Function
	End If

	If Columns <= 0 Then
		GetMaxRow = -1
		Exit Function
	End If

	With Sheet
		GetMaxRow = .Cells( .Rows.Count , Columns ).End(xlUp).Row
	End With
End Function

'-------------------------------------------
' バッチの実行
' ファイルへのパス : FilePath
' 終了待ちをするか : bSync [省略可能] True
' コマンドウィンドウの表示状態 : nCmdShow [省略可能] 通常
'-------------------------------------------
Public Sub ExecBatch( FilePath As String , Optional bSync As Boolean = True , Optional nCmdShow As Integer = CMD_NORMAL )
	Dim Cmd As Object
	Set Cmd = CreateObject("WScript.Shell")

	Cmd.Run FilePath , nCmdShow , bSync

	Set Cmd = Nothing
End Sub

'-------------------------------------------
' 相対パスから絶対パスを作成する
' 相対パス : RelativePath
' 基本パス : RootPath[ 省略時インポートブック ]
'-------------------------------------------
Public Function GetAbsFilePath( RelativePath As String , Optional RootPath As String = "" ) As String
	Dim Path As String
	If Len(RootPath) > 0 Then
		Path = RootPath
	Else
		Path = ThisWorkbook.Path
	End If

	Path = Path & "/" & RelativePath
	
	Dim FSO As Object
	Set FSO = CreateObject("Scripting.FileSystemObject")

	GetAbsFilePath =  FSO.GetAbsolutePathName(Path)

	Set FSO = Nothing
End Function


'-------------------------------------------
' フォーカス状態にするプロパティ
'-------------------------------------------
Property Let Focus( ByVal Flag As Boolean )
	With Application
		.ScreenUpdate = Not Flag
		.EnableEvents = Not Flag
		.Calculation = IIf( Flag , xlCalculationManual , xlCalculationAutomatic )
	End With
End Property

'-------------------------------------------
' ファイルの有無を確認する
' ファイルパス : FilePath
'-------------------------------------------
Public Function IsExistFile( FilePath As String ) As Boolean
	Dim Path As String
	If Left( FilePath , 1 ) = "." Then
		Path = GetAbsFilePath( FilePath )
	Else
		Path = FilePath
	End If

	With CreateObject("Scripting.FileSystemObject")
		IsExistFile = .IsFileExists( Path )
	End With

End Function














