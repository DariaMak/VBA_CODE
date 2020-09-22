Attribute VB_Name = "Module7"
'Update from array all(tables and charts) as JPG 2 - last version


'step 1 - go through every chart's name in ChartNames EXCEL sheet/make array
Function DICTofCharts() As Scripting.Dictionary
    Dim ExcelChDICT As Scripting.Dictionary
    Dim xChart As ChartObject
    Dim xSheet As Worksheet
    Dim nm As String
    Dim rm As Chart
    
    Set ExcelChDICT = New Scripting.Dictionary
    For Each xSheet In Worksheets
        For Each xChart In xSheet.ChartObjects
            nm = CStr(xChart.name)
            Set rm = xSheet.ChartObjects(nm).Chart
            If Not ExcelChDICT.Exists(nm) Then
                ExcelChDICT.Add key:=nm, Item:=rm
            End If
        Next xChart
    Next xSheet
Set DICTofCharts = ExcelChDICT
End Function


'step 2 - go through every table's name in ChartNames EXCEL sheet/make array
Function DICTofTbls() As Scripting.Dictionary
    Dim ExcelTDICT As Scripting.Dictionary
    Dim xTable As ListObject
    Dim xSheet As Worksheet
    Dim nm As String
    Dim rm As Range
    
    Set ExcelTDICT = New Scripting.Dictionary
    For Each xSheet In Worksheets
        For Each xTable In xSheet.ListObjects
            nm = CStr(xTable.name)
            Set rm = xSheet.ListObjects(nm).Range
            If Not ExcelTDICT.Exists(nm) Then
                ExcelTDICT.Add key:=nm, Item:=rm
            End If
        Next xTable
    Next xSheet
Set DICTofTbls = ExcelTDICT
End Function


'step 3 - update textboxes
Function DICTofTBs() As Scripting.Dictionary
    Dim TextBoxDICT As Scripting.Dictionary
    Dim xSheet As Worksheet
    Dim cell As Range
    Dim n As Long
    Dim LastRow As Long
    Dim nm As String
    Dim rm As String
    
    Set xSheet = ThisWorkbook.Sheets("Slides")
    LastRow = xSheet.Cells(xSheet.Rows.Count, "m").End(xlUp).Row
    Set TextBoxDICT = New Scripting.Dictionary
    
    For n = 3 To LastRow
        For Each cell In xSheet.Cells(n, 13)
                nm = cell.Offset(0, 1).Value
                rm = cell.Value
                TextBoxDICT.Add key:=nm, Item:=rm
                'Debug.Print nm, rm
         Next cell
     Next n
Set DICTofTBs = TextBoxDICT
End Function


'step 4 - MAIN - find PPT shape's name in PPT objects
Sub UpdateMAIN()
    Dim PPApp As PowerPoint.Application
    Dim PPPres As PowerPoint.Presentation
    Dim PPSlide As PowerPoint.slide
    Dim PPShape As PowerPoint.Shape
    Dim PPChart As PowerPoint.Chart
    Dim Check As Boolean
    
    ipath = ThisWorkbook.Path & "\PPT2.pptx"
    Set PPApp = New PowerPoint.Application
    Set PPPres = PPApp.Presentations(ipath)
    
    For Each PPSlide In PPPres.Slides
        Debug.Print PPSlide.name & vbCrLf;
            For Each PPShape In PPSlide.Shapes
                'step 5 -------------------------------- step 1
                If PPShape.name = findKey(DICTofCharts, PPShape.name) Then
                '------------------ step 6 -------------
                    Call UpdateChartByName(PPShape.name, PPSlide.SlideIndex + 0, 1)
                'step 5 ------------------------------- step 2
                ElseIf PPShape.name = findKey(DICTofTbls, PPShape.name) Then
                '------------------ step 6 -------------
                    Call UpdateChartByName(PPShape.name, PPSlide.SlideIndex + 0, 2)
                'step 5 ------------------------------- step 3
                ElseIf PPShape.name = findKey(DICTofTBs, PPShape.name) Then
                '------------------ step 6 -------------
                    Call UpdateChartByName(PPShape.name, PPSlide.SlideIndex + 0, 3)
                End If
            Next PPShape
    Next PPSlide
End Sub


'step 5 - find PPT shape's name in excel dictionary
Function findKey(dict As Scripting.Dictionary, shapeName As String) As String
Dim key As Variant
    For Each key In dict.Keys
       If key = shapeName Then
        findKey = CStr(key)
        Exit Function
       End If
    Next
End Function


'step 6 - update chart by name, copy and paste
Function UpdateChartByName(objectName As String, slideNumber As Long, caseIndex As Long)
    Dim PPApp As PowerPoint.Application
    Dim PPPres As PowerPoint.Presentation
    Dim PPSlide As PowerPoint.slide
    Dim PPShape As PowerPoint.Shape
    Dim ipath As String
    Dim ws As Excel.Worksheet

    ipath = ThisWorkbook.Path & "\PPT2.pptx"
    Set PPApp = New PowerPoint.Application
    Set PPPres = PPApp.Presentations(ipath)
    Set ws = ThisWorkbook.Worksheets("Slides")
           
    Set PPSlide = PPPres.Slides(slideNumber)
    With PPSlide.Shapes(objectName)
        Dim left As Double
        Dim top As Double
        Dim height As Double
        Dim width As Double
        
        left = .left
        top = .top
        height = .height
        width = .width
            
        Select Case caseIndex
            Case 1
            PPSlide.Shapes(objectName).Delete
            Dim data1 As Chart
            Set data1 = DICTofCharts(objectName)
            With data1
                .ChartArea.Copy
                With PPSlide.Shapes.PasteSpecial(ppPasteJPG)
                    .left = left
                    .top = top
                    .width = width
                    .height = height
                    .name = CStr(objectName)
                End With
            End With
            Case 2
            PPSlide.Shapes(objectName).Delete
            Dim data2 As Range
            Set data2 = DICTofTbls(objectName)
            With data2
                .Copy
                With PPSlide.Shapes.PasteSpecial(ppPasteBitmap)
                    .left = left
                    .top = top
                    .width = width
                    .height = height
                    .name = CStr(objectName)
                End With
            End With
            Case 3
            Dim data3 As String
            data3 = DICTofTBs(objectName)
            .TextFrame.TextRange = data3
        End Select
    End With
End Function

