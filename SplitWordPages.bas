Function GetFolder() As String
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    .InitialFileName = ""
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetFolder = sItem
Set fldr = Nothing
End Function

Sub SavePagesAsDoc()
    Dim orig As Document
    Dim page As Document
    Dim numPages As Integer
    Dim idx As Integer
    Dim fn As String
    Dim oldFilename As String
    fn = GetFolder()
    If fn = "" Then
        MsgBox ("Nie wybrano œcie¿ki, koñczê program!")
        End
    End If
        oldFilename = ActiveDocument.Name
    If Right(oldFilename, 5) = ".docx" Then
'        MsgBox ("subtract .docx")
        oldFilename = Left(oldFilename, Len(oldFilename) - 5)
    ElseIf Right(oldFilename, 4) = ".doc" Then
'        MsgBox ("subtract .doc")
        oldFilename = Left(oldFilename, Len(oldFilename) - 4)
    Else
        MsgBox ("no extension yet")
    End If

    If fn <> "" Then
        If Right$(fn, 1) <> "\" Then fn = fn + "\"
            
            ' Keep a reference to the current document.
                Set orig = ActiveDocument
            ' Calculate the number of pages
                numPages = ActiveDocument.Range.Information(wdActiveEndPageNumber)
            ' Create a new document
                Set page = Documents.Add(, , , False)
                For idx = 1 To numPages
            ' Make sure the document is active
                orig.Activate
            ' Go to the page with index idx
                Selection.GoTo What:=wdGoToPage, Name:=idx
            ' Select the current page
                Selection.GoTo What:=wdGoToBookmark, Name:="\page"
            On Error GoTo Error_MayCauseAnError
            ' Copy the selection
                Selection.Copy
            ' Activate it
                page.Activate
                Selection.TypeText (vbCrLf + "/####################|" + CStr(idx) + "|####################\" + vbCrLf)
            ' Paste the selection
                Selection.Paste
Error_MayCauseAnError:
            page.Activate
            Selection.TypeText ("This page does not contain content")
             Next
    End If
        ' Generate the file name
             fntxt = fn + oldFilename
        ' Save the document as Word 97-2003
             page.SaveAs FileName:=fntxt, FileFormat:=wdFormatText, AddToRecentFiles:=False
        ' Close the document
            page.Close

End Sub
