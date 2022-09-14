'******************************************************************************'
' File  : db2awl.vbs
' Date  : 2022.09.13
'******************************************************************************'
Option Explicit

Dim sArg

For Each sArg In Wscript.Arguments
    If (IsValidFile(sArg)) Then
        ConvertDbToAwl sArg
    End If
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

Private Function ConvertDbToAwl( ByVal sFilePath )
    Const adTypeText = 2
    Const adModeReadWrite = 3
    Const adSaveCreateOverWrite = 2
    Dim sText

    ' Read .db
    With CreateObject("ADODB.Stream")
        .Charset = "utf-8"
        .Type = adTypeText
        .Open
        .LoadFromFile sFilePath
        ' .LineSeparator = adLF
        
        sText = .ReadText

        .Close
    End With

    With New RegExp
        .IgnoreCase = True
        .Global = True
        
        ' Match to, e.g.: { S7_SetPoint := 'True'}
        .Pattern = "\{(\s+)?S7\_\w+(\s+)?:=(\s+)?'(True|False)'(\s+)?\}(\s+)?"
        sText = .Replace(sText, "")
    End With

    ' Write/save .awl
    With CreateObject("ADODB.Stream")
        .Type = adTypeText
        .Mode = adModeReadWrite
        .Open
        .Position = 0
        .WriteText sText

        .SaveToFile Replace(sFilePath, ".db", ".AWL", 1, -1, vbTextCompare), _
                adSaveCreateOverWrite
        .Close
    End With
End Function