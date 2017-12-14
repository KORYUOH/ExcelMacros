'================================================================
' vim:fenc=cp932:ft=vb
' Brief : �ėp�}�N�����W���[��
' Author : KORYUOH
' Create : 2017/12/14
' Update : 2017/12/14
' Version : 0.01
'================================================================
Attribute VB_Name = "CommonMacroLib"
Option Explicit

'-------------------------------------------
' �V�[�g���L�邩
' �����u�b�N : book
' �����V�[�g�� : sheetName
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
End Function

'-------------------------------------------
' �V�[�g�̍ŏI�f�[�^�s���擾����
' �ΏۃV�[�g : Sheet
' �Ώۗ� : Columns [�ȗ��\]
'-------------------------------------------
Public Function GetMaxRow( Sheet As WorkSheet , Optional Columns As Long = 1 ) As Integer

	If Sheet Is Nothing Then
		GetMaxRow = -1
		Exit Function
	End If

	If Columns <= 0 Then
		GetMaxRow = -1
		Exit Function
	End If

	With Sheet
		.Cells( .Row.Count , Columns ).End(xlUp).Row
	End With
End Function

'-------------------------------------------
' �o�b�`�N���̃R�}���h�v�����v�g�̕\�����
'-------------------------------------------
Public Const CMD_HIDE As Integer = 0    ' ��\��
Public Const CMD_NORMAL As Integer = 1  ' �ʏ�
Public Const CMD_MINIMUM As Integer = 2 ' �ŏ���
Public Const CMD_MAXIMUM As Integer = 3 ' �ő剻

'-------------------------------------------
' �o�b�`�̎��s
' �t�@�C���ւ̃p�X : FilePath
' �I���҂������邩 : bSync [�ȗ��\] True
' �R�}���h�E�B���h�E�̕\����� : nCmdShow [�ȗ��\] �ʏ�
'-------------------------------------------
Public Sub ExecBatch( FilePath As String , Optional bSync As Boolean = True , Optional nCmdShow As Integer = CMD_NORMAL )
	Dim Cmd As Object
	Set Cmd = CreateObject("WScript.Shell")

	Cmd.Run FilePath , nCmdShow , bSync

	Set Cmd = Nothing
End Sub

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

Private Type FocusSettings
	Updateing As Boolean
	Events As Boolean
	Calc As Integer
End Type

Property Let Trans( ByVal Flag As Boolean )
	With Application
		.ScreenUpdate = Not Flag
		.EnableEvents = Not Flag
		.Calculation = IIf( Flag , xlCalclationManual , xlCalclationAutomatic )
	End With
End Property

Sub Focus( InSet As FocusSettings )

	With Application
		.ScreenUpdate = InSet.Updateing
		.EnableEvents = InSet.Events
		.Calculation = InSet.Calc
	End With

End Sub

Function EnableFocus() As FocusSettings
	With Application
		EnableFocus.Updateing = .ScreenUpdateing
		EnableFocus.Events = .EnableEvents
		EnableFocus.Calc = .Calculation
		.ScreenUpdate = False
		.EnableEvents = False
		.Calculation = xlCalclationManual
	End With
End Function

Sub DisableFocus( Optional Before As FocusSettings = Nothing )
	With Application
		If Before Is Nothing Then
			.ScreenUpdateing = True
			.EnableEvents = True
			.Calculation = xlCalclationAutomatic
		Else
			.ScreenUpdateing = Before.Updateing
			.EnableEvents = Before.Events
			.Calculation = Before.Calc
		End If
	End With
End Sub












