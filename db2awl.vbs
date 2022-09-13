'******************************************************************************'
<<<<<<< HEAD
' File        : db2awl.vbs
' Author      : yuriestevan@gmail.com
' Date        : 2022/09/13
' Description : Convert .db to .awl
' Revision    : 0.0
' Date      Author  Description
' ------------------------------------------------------------------------------
' 22/09/13  yes     First commit
'
'******************************************************************************'

If WScript.Arguments.Count > 0 Then
    Const xlCellTypeLastCell = 11
    Dim file
    Dim fso
    Dim dict
    Dim objWb
    Dim objWs
    Dim objXl
    Dim iColTagName
    Dim iColDescription
    Dim iColAddress









msgbox (wscript.arguments.count)

=======
' File	: db2awl.vsb
' Date	: 2022.09.13
'******************************************************************************'
Option Explicit

Dim sArg

For Each sArg In Wscript.Arguments
	Wscript.Echo sArg, IsValidFile(sArg)
Next 'sArg


Private Function IsValidFile( sFilePath )
	Dim bIsValidFile : bIsValidFile = False

	With CreateObject("Scripting.FileSystemObject")
		If ( .FileExists(sFilePath) ) Then
			If (InStr( 1, .GetExtensionName(sFilePath), "db", vbTextCompare ) > 0) Then
				bIsValidFile = True
			End If
		End If
	End With

	IsValidFile = bIsValidFile
End Function
>>>>>>> feature/check-if-files-is-valid
