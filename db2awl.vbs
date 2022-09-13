'******************************************************************************'
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

