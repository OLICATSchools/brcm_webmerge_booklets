Function GetFolder() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

Sub BatchPrintWordDocuments()
      
    Dim strFolder As String
    Dim iCount As Integer
    strFolder = GetFolder()
  
    iCount = 0
    Dim myDoc As Word.Document
    Dim strFile As String
    strFile = Dir(strFolder & "\" & "*.doc*", vbNormal)
    
    
    While strFile <> ""
            Application.Documents.Open FileName:=strFolder & "\" & strFile, Visible:=False ' Document to open in hidden mode
            Set myDoc = Application.Documents(strFolder & "\" & strFile)
            myDoc.PrintOut
            myDoc.Close savechanges:=False ' closes my Doc
            Set myDoc = Nothing
            strFile = Dir()
            iCount = iCount + 1
    Wend
 
    MsgBox iCount & " documents have been queued for printing."
End Sub
