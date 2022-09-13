'******************************************************************************'
' File	: db2awl.vsb
' Date	: 2022.09.13
'******************************************************************************'
Option Explicit

Dim oArg

For Each oArg In Wscript.Arguments
	Wscript.Echo oArg, IsValidFile(oArg)
Next 'oArgs


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