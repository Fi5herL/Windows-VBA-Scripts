Attribute VB_Name = "NewMacros"
Sub ExportUnresolvedCommentsToEndOfDocument()
    Dim doc As Document
    Dim comment As comment
    Dim unresolvedComments As String
    Dim docPath As String
    Dim docName As String
    
    Set doc = ActiveDocument
    docPath = ThisDocument.Path & "\"
    docName = Left(doc.Name, InStrRev(doc.Name, ".") - 1) ' 取得不含副檔名的檔名
    
    ' 檢查並收集所有未解決的評論
    For Each comment In doc.Comments
        If Not comment.Done Then
            unresolvedComments = unresolvedComments & "Author: " & comment.Author & vbCrLf & _
                                 "Time: " & comment.Date & vbCrLf & _
                                 "Comment: " & comment.Range.Text & vbCrLf & vbCrLf
        End If
    Next comment
    
    ' 在文件末端插入未解決的評論
    Selection.EndKey Unit:=wdStory
    Selection.TypeText Text:=unresolvedComments
End Sub
