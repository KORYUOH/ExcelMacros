'================================================================
' vim:fenc=cp932:ft=vb
' Brief : �ėp�}�N�����W���[��
' Author : KORYUOH
' github : KORYUOH/ExcelMacros
' Create : 2017/12/14
' Update : 2018/02/19
' Version : 0.15
'================================================================
Attribute VB_Name = "CommonMacroLib"
Option Explicit

'-------------------------------------------
' �o�b�`�N���̃R�}���h�v�����v�g�̕\�����
'-------------------------------------------
Public Const CMD_HIDE As Integer = 0    ' ��\��
Public Const CMD_NORMAL As Integer = 1  ' �ʏ�
Public Const CMD_MINIMUM As Integer = 2 ' �ŏ���
Public Const CMD_MAXIMUM As Integer = 3 ' �ő剻


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
	End With
End Function

'-------------------------------------------
' �u�b�N���J����Ă��邩
'-------------------------------------------
Function IsOpenedBook( BookName As String ) as Boolean
	IsOpenedBook = False
	Dim wb As Workbook
	For Each wb In Workbooks
		if UCase( wb.Name ) = UCase( BookName ) Then
			IsOpenedBook = True
			Exit Function
		End If
	Next wb
End Function

'-------------------------------------------
' �u�b�N���J��
' �J����Ă����炻����擾����
'-------------------------------------------
Function OpenBook( BookName As String , Optional Path As String = "." )As Workbook
	Set OpenBook = Nothing
	If IsOpenedBook( BookName ) Then
		Dim wb As Workbook
		For Each wb In Workbooks
			If UCase( wb.Name ) = UCase( BookName ) Then
				Set OpenBook = wb
				Exit Function
			End If
		Next wb
	End If

	Dim FilePath As String

	if Path = "." Then
		FilePath = ThisWorkbook.Path
	Else
		FilePath = Path
	End If


	If Not Right( FilePath , 1 ) = "\" Then
		FilePath = FilePath &"\"&BookName
	Else
		FilePath = FilePath & BookName
	End If

	Set OpenBook = Workbooks.Open ( FilePath )
End Function

'-------------------------------------------
' �V�[�g�̍ŏI�f�[�^�s���擾����
' �ΏۃV�[�g : Sheet
' �Ώۗ� : Columns [�ȗ��\]
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

'-------------------------------------------
' ���΃p�X�����΃p�X���쐬����
' ���΃p�X : RelativePath
' ��{�p�X : RootPath[ �ȗ����C���|�[�g�u�b�N ]
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
' �t�H�[�J�X��Ԃɂ���v���p�e�B
'-------------------------------------------
Property Let Focus( ByVal Flag As Boolean )
	With Application
		.ScreenUpdating = Not Flag
		.EnableEvents = Not Flag
		.Calculation = IIf( Flag , xlCalculationManual , xlCalculationAutomatic )
	End With
End Property

'-------------------------------------------
' �t�@�C���̗L�����m�F����
' �t�@�C���p�X : FilePath
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














