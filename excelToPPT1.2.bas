Attribute VB_Name = "Module1"
 Dim ppt_app As New PowerPoint.Application
 Dim pre As PowerPoint.Presentation
 Dim slde As PowerPoint.Slide
 Dim shp As PowerPoint.Shape
 Dim chrt As PowerPoint.Chart
 Dim chrtdata As ChartData
 Dim ipath As String
 

Sub GetObjectNameSize()
  ipath = ThisWorkbook.Path & "\test.pptx"
  
  Set ppt_app = CreateObject("PowerPoint.Application")
  Set pre = ppt_app.Presentations(ipath)

  Dim left As Double
  Dim top As Double
  Dim height As Double
  Dim width As Double
  
  'Objects name and size
   
    For Each shp In pre.Slides(1).Shapes
        Debug.Print shp.name
        Debug.Print shp.left
        Debug.Print shp.width
        Debug.Print shp.top
        Debug.Print shp.height
        Next
    
End Sub


Sub UpdateChartRange()
    ipath = ThisWorkbook.Path & "\test.pptx"
    
    Set ppt_app = CreateObject("PowerPoint.Application")
    Set pre = ppt_app.Presentations(ipath)
    
    
    Worksheets("Admin").ChartObjects(1).Activate
    ActiveChart.SeriesCollection.Add _
    Source:=Worksheets("Admin").Range("c19:w22")
    
    
    
End Sub


Sub UpdateData()
    ipath = ThisWorkbook.Path & "\test.pptx"
    
    Set ppt_app = CreateObject("PowerPoint.Application")
    Set pre = ppt_app.Presentations(ipath)
    
    With pre.Slides(1)
       
        .Shapes("textbox1").TextFrame.TextRange = ThisWorkbook.Sheets("Admin").Range("o7").Value
        .Shapes("textbox2").TextFrame.TextRange = ThisWorkbook.Sheets("Admin").Range("o8").Value
        .Shapes("textbox3").TextFrame.TextRange = ThisWorkbook.Sheets("Admin").Range("o9").Value
        .Shapes("textbox4").TextFrame.TextRange = ThisWorkbook.Sheets("Admin").Range("o10").Value
       
    End With
    
End Sub

Sub DeleteSlide()
    ipath = ThisWorkbook.Path & "\test.pptx"
    
    Set ppt_app = CreateObject("PowerPoint.Application")
    Set pre = ppt_app.Presentations(ipath)
    
    ' start from the last slide
    For i = pre.Slides.Count To 1 Step -1
        If ThisWorkbook.Sheets("Admin").Cells(19 + i, 15).Value = True Then pre.Slides(i).Delete
        Next
        
End Sub



Sub UpdateLinks()
    ipath = ThisWorkbook.Path & "\test.pptx"
    
    Set ppt_app = CreateObject("PowerPoint.Application")
    Set pre = ppt_app.Presentations(ipath)

    For Each slde In pre.Slides
    
       For Each shp In slde.Shapes
         On Error Resume Next
         shp.LinkFormat.update
        Next
    Next

End Sub


