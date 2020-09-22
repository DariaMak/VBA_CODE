Attribute VB_Name = "Module1"
Sub UpdateByTitle()
    Dim WordApp As Word.Application
    Dim WordDoc As Word.Document
    Dim ws As Worksheet
    Dim ExCell As Range
    Dim ExRange As Excel.Range
    Dim strWFile As String
    
    Dim a, b, c As Object
    Dim n As Integer
    
    ThisFileName = ThisWorkbook.Name
    strWFile = ActiveWorkbook.Path & "\test.docx"
    Set WordDoc = GetObject(strWFile)
    
    n = 2
    Set ws = ActiveSheet
    ws.Name = "Sheet1"
    Set ExRange = ws.Range("c2:c9")
    Set ExCell = ExRange.Range("c" & n)
    
        For Each ExCell In ExRange
            ws.Range("b" & n).Activate
            b = ActiveCell.Value
            byTitle CStr(ExCell), CStr(b)
            n = n + 1
        Next ExCell
End Sub


Sub getTitleContentTag()
    Dim ctrl As ContentControl
        For Each ctrl In ActiveDocument.SelectContentControlsByTag("title")
            Debug.Print ctrl.Tag, ctrl.PlaceholderText, ctrl.Title
        Next
End Sub


Sub setPlaceHolderText()
    Dim cc As ContentControl
        For Each cc In ActiveDocument.ContentControls
            If cc.Tag = "title" Then
                If cc.ShowingPlaceholderText Then
                    cc.setPlaceHolderText Text:="da"
                End If
            End If
        Next cc
End Sub


Sub UpdateTest()
    Dim WordApp As Word.Application
    Dim WordDoc As Word.Document
    Dim ws As Worksheet
    Dim strWFile As String
    
    Dim cc As Word.ContentControl
    Dim rngCC As Word.Range
    Dim a, b, c As String
    Dim n As Integer
    
    ThisFileName = ThisWorkbook.Name
    strWFile = ActiveWorkbook.Path & "\test.docx"
    Set WordDoc = GetObject(strWFile)
    'WordDoc.Activate
    
    Set ws = ActiveSheet
    ws.Name = "Sheet1"
    n = 2
    
        For Each cc In WordDoc.ContentControls
            If cc.Tag = "title" Then
            With ws.Range("c2:c9")
                Set c = .Find(cc.Title, LookIn:=xlValues)
            End With
                
                c = ActiveCell.Value
                If cc.Title = c Then
                    ws.Range("b" & n).Activate
                    b = ActiveCell.Value
                        Debug.Print b
                        Debug.Print n
                    cc.Range.Text = ""
                    cc.setPlaceHolderText Text:=CStr(b)
                End If
             End If
         n = n + 1
         Next cc
End Sub



Function byTitle(x, y)
    Dim wDoc As Word.Document
    Dim cc As ContentControl
    Dim ccs As ContentControls
    
    Set wDoc = ActiveDocument
    Set ccs = wDoc.SelectContentControlsByTitle(x)
        For Each cc In ccs
            If ccs.Count > 0 Then
            cc.Range.Text = ""
            cc.setPlaceHolderText Text:=y
            End If
        Next
End Function


