Sub SplitNotes()
Dim oDoc As Document
Dim lngIndex As Long, lngCount As Long
Dim oRng As Range
Dim oCol As New Collection
Dim bFound As Boolean
Dim outFolder As String
Dim strDelim As String
Dim strFileName As String

strDelim = "<BreakHere>"
outFolder = GetFolder()

  bFound = False
  Set oDoc = ActiveDocument
  Set oRng = oDoc.Range
    strFileName = Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1)
  With oRng.Find
    .Text = strDelim
    While .Execute
      If lngCount = 0 Then
        oRng.Start = ActiveDocument.Range.Start
        oCol.Add oRng.Duplicate
        oRng.Collapse wdCollapseEnd
        lngCount = lngCount + 1
        bFound = True
      Else
        oRng.Start = oCol.Item(oCol.Count).End
        oCol.Add oRng.Duplicate
        oRng.Collapse wdCollapseEnd
      End If
    Wend
    If bFound Then
      oRng.End = ActiveDocument.Range.End - 1
      oRng.InsertAfter strDelim
      oCol.Add oRng.Duplicate
    End If
  End With
  If oCol.Count > 0 Then
    If MsgBox("This will split the document into " & oCol.Count - 1 & " sections. Do you wish to proceed?", _
               vbQuestion + vbYesNo, "SPlIT") = vbNo Then Exit Sub
  End If
  For lngIndex = 2 To oCol.Count
    Set oDoc = Documents.Add
    
    With oDoc.PageSetup
        .TopMargin = CentimetersToPoints(1)
        .BottomMargin = CentimetersToPoints(1.25)
        .LeftMargin = CentimetersToPoints(1.25)
        .RightMargin = CentimetersToPoints(1.25)
    End With
    
    With oDoc.Styles(wdStyleNormal).Font
        .Name = "Arial"
        .Size = 12
    End With
      
    oDoc.Range.FormattedText = oCol.Item(lngIndex).FormattedText
    For lngCount = 1 To Len(strDelim)
      oDoc.Range.Characters.Last.Previous.Delete
    Next
    
    With oDoc
        
        With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        End With
        
    Selection.EndKey Unit:=wdStory
    Selection.TypeBackspace
    Selection.TypeBackspace
    End With
    
    oDoc.SaveAs outFolder & "\" & strFileName & " " & Format(lngIndex - 1, "000")
    oDoc.Close True
    
  Next lngIndex
lbl_Exit:
  Exit Sub
End Sub