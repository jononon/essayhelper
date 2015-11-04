Sub ChangeFonts()
Dim doc As Document
Set doc = ActiveDocument

'DEPRECIATED: Too slow ¯\_(O_O)_/¯
'For i = 1 To doc.Range.Characters.Count
'    If doc.Range.Characters(i) = " " Then
'        doc.Range.Characters(i).Font.Size = "14"
'    If doc.Range.Characters(i) = "," Then
'        doc.Range.Characters(i).Font.Size = "14"
'    End If
'    If doc.Range.Characters(i) = "," Then
'        doc.Range.Characters(i).Font.Size = "14"
'    End If
'Next

Dim FontSize As Integer
FontSize = doc.Range.Characters(1).Font.Size + 2

With Selection.Find
    .ClearFormatting
    .Text = " "
    .Replacement.ClearFormatting
    .Replacement.Font.Size = FontSize
    .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
End With
With Selection.Find
    .ClearFormatting
    .Text = ","
    .Replacement.ClearFormatting
    .Replacement.Font.Size = FontSize
    .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
End With
With Selection.Find
    .ClearFormatting
    .Text = "."
    .Replacement.ClearFormatting
    .Replacement.Font.Size = FontSize
    .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
End With
With Selection.Find
    .ClearFormatting
    .Text = "-"
    .Replacement.ClearFormatting
    .Replacement.Font.Size = FontSize
    .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
End With



End Sub
