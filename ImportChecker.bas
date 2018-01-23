' vim:fenc=cp932:ft=vb
Option Explicit
Attribute VB_Name = "TempMacros"
const FileName As String = "TempMacros"

Sub ClearTmpModules
	Dim Component As Variant
	With ThisWorkbook.VBProject
		For Each Component In .VBComponents
			If Component.Type = 1 Then
				If Component.Name = FileName Then
					Component.Name = Component.Name & "OLD"
					.VBComponents.Remove Component
				End If
			End If
		Next Component
	End With
End Sub

Sub TempMacro
	' Ç±Ç±Ç…èëÇ≠
	' ThisWorkbook.Sheets(1).Range("A1").Value = "Hello,world"
End Sub
