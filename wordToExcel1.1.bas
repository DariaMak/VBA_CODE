Attribute VB_Name = "Module1"
Sub UpdateWordDoc()
    Dim WordApp As Word.Application
    Dim WordDoc As Word.Document
    Dim strWFile As String
 
    ThisFileName = ThisWorkbook.Name
    strWFile = ActiveWorkbook.Path & "\excel2word1.1.docx"
    Set WordDoc = GetObject(strWFile)
      
    WordDoc.Fields.Update
End Sub

    
Sub UpdateSelectedLine()
    Worksheets("KPI").Activate
 
    Worksheets("KPI").Range("r1").Activate
    ActiveDocument.Sentences(1) = ActiveCell.Value & Chr(10)
    
    Worksheets("KPI").Range("r3").Activate
    ActiveDocument.Sentences(2) = ActiveCell.Value & Chr(10)
    
    Worksheets("KPI").Range("r5").Activate
    ActiveDocument.Sentences(3) = ActiveCell.Value & Chr(10)
 
End Sub


Sub updateBookM1()
    Worksheets("KPI").Range("r13").Activate
    ActiveDocument.Bookmarks("line1").Select
    Dim txt As Variant
    
    Set txt = Worksheets("KPI").Range("r13")
   'Insert text after named bookmark
    ActiveDocument.Bookmarks("line1").Range.InsertBefore txt & Chr(32)
 End Sub
 
 
Sub updateBookM2()
    
    Worksheets("KPI").Range("r15").Activate
    ActiveDocument.Bookmarks("line2").Select
    Dim txt As Variant
    
    Set txt = Worksheets("KPI").Range("r15")
    'Replace the text inside the bookmark but also delete the bookmark
    ActiveDocument.Bookmarks("line2").Range.Text = CStr(txt) & Chr(32)
End Sub

Sub updateBookM3()
    
    Worksheets("KPI").Range("r17").Activate
    ActiveDocument.Bookmarks("line3").Select
    Dim txt As Variant
    
    Set txt = Worksheets("KPI").Range("r17")
    'insert text in place
    ActiveDocument.Bookmarks("line3").Range.Text = CStr(txt) & Chr(32)
End Sub


Sub setAutoUpdateToManual()
    Dim WordApp As Word.Application
    Dim WordDoc As Document
    Dim strWFile As String
 
    ThisFileName = ThisWorkbook.Name
    strWFile = ActiveWorkbook.Path & "\excel2word1.1.docx"
    Set WordDoc = GetObject(strWFile)
    
    Documents(WordDoc).Activate
    
    For Each fieldLoop In ActiveDocument.Fields
        If fieldLoop.LinkFormat.AutoUpdate = True Then _
        fieldLoop.LinkFormat.AutoUpdate = False
    Next fieldLoop
End Sub

Sub HighlightBookmarkedItemsInADoc()
  Dim objBookmark As Bookmark
  Dim objDoc As Document
 
  Application.ScreenUpdating = False
 
  Set objDoc = ActiveDocument
 
  With objDoc
    For Each objBookmark In .Bookmarks
      objBookmark.Range.HighlightColorIndex = wdBrightGreen
    Next objBookmark
  End With
  Application.ScreenUpdating = True
End Sub


Sub UpdateDeveloper()
    Dim WordApp As Word.Application
    Dim WordDoc As Word.Document
    Dim strWFile As String
    Dim c As Integer
 
    ThisFileName = ThisWorkbook.Name
    strWFile = ActiveWorkbook.Path & "\excel2word1.1.docx"
    Set WordDoc = GetObject(strWFile)
    c = 20
    
    For i = 1 To 3
        WordDoc.ContentControls(i).Range.Text = Sheets("KPI").Cells(25, c)
        c = c + 1
    Next i
    
    
End Sub

