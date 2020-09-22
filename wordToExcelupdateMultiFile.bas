Attribute VB_Name = "Module3"
Sub TimeExcelToWord()
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    StartTime = Timer
    ExcelToWord
    SecondsElapsed = Round(Timer - StartTime, 2)
    MsgBox "This code ran successfully in " & SecondsElapsed & " seconds"
End Sub

Sub ExcelToWord()
    Dim FirstAddress As String
    Dim MySearch As String
    Dim Rng As range
    Dim sh As Worksheet
    Dim dict As Scripting.dictionary
    
    MySearch = "#Field#"
    'WordName = "Example template1.docx"
    Set dict = New Scripting.dictionary
    
    For Each sh In ActiveWorkbook.Worksheets
        Set Rng = sh.Cells(1, 1).EntireRow.find(What:=MySearch, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext)

        If Not Rng Is Nothing Then
            FirstAddress = Rng.Address
            Do
                If Rng.Offset(0, 1).Value = "#Value#" Then
                    Set dict = updateDictionary(Rng.Column, Rng.Worksheet.Name, dict)
                    Debug.Print dict.count
                End If
                Set Rng = sh.Cells(1, 1).EntireRow.FindNext(Rng)
            Loop While Not Rng Is Nothing And Rng.Address <> FirstAddress
        End If
    Next sh
    
     UpdateValues dict
    
End Sub


Function updateDictionary(cellN, sheetName, dict) As Scripting.dictionary
    With Worksheets(sheetName)
        lastRow = .Cells(Rows.count, cellN).End(xlUp).Row
    
        For counter = 2 To lastRow
            Set curCell = .Cells(counter, cellN)
            If Len(curCell.Value) > 0 Then
                dict.Add Key:=curCell.Value, Item:=curCell.Offset(0, 1).Value
            End If
        Next counter
    End With
    Set updateDictionary = dict
End Function

Function UpdateValues(dict)
    On Error GoTo ErrorHandling
    
    Dim WordApp As Object, WordDoc As Object
    Dim openDoc As Word.Document
    Dim cc As ContentControl, ccs As ContentControls
    Dim Key As Variant
    Dim WordName As String
    
    WordName = dict("Word_File")
 '--------------------------------------array
    Dim arr As Variant
    Dim i, w As Long
    i = 0
    ReDim arr(i)
    
    For Each openDoc In Documents
        If CStr(openDoc.Name) = WordName Then
            ReDim Preserve arr(i + 1)
            arr(i) = CStr(openDoc.Path & "\" & openDoc.Name)
            Debug.Print openDoc.Name
            Debug.Print openDoc.Path
            Debug.Print arr(i)
            i = i + 1
        End If
    Next openDoc
 '--------------------------------------array
    
    w = UBound(arr)
    For i = 0 To w
        If IsEmpty(arr(i)) Then
            'nothing
        Else
            Set WordApp = GetObject(Class:="Word.Application")
            Set WordDoc = WordApp.Documents.Open(Filename:=CStr(arr(i)), ReadOnly:=False)
                With WordDoc
                    For Each Key In dict
                        Debug.Print (Key & " " & dict(Key))
                        Set ccs = .SelectContentControlsByTitle(Key)
                        If ccs.count > 0 Then
                            For Each cc In ccs
                                cc.range.Text = dict(Key)
                            Next
                        End If
                    Next Key
                End With
         End If
     Next i
    
Done:
    Exit Function
ErrorHandling:
    If Err.Number = 429 Or Err.Number = 4160 Then
        MsgBox "Please open the following Word file: " & WordName
    Else:
        MsgBox "The following error occurred: " & Err.Number & " " & Err.Description
    End If
    
         
End Function

