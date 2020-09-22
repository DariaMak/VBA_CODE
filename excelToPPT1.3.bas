Attribute VB_Name = "Module6"
        Dim rng As Range
        Dim wb As Workbook
        Dim PPApp As PowerPoint.Application
        Dim PPPres As PowerPoint.Presentation
        Dim PPSlide As PowerPoint.Slide
        Dim shp As PowerPoint.Shape
        ' Dim chrtName As String
        
        
Sub RangeToPresentation()

        ' Make sure a table is selected
        
    If Not TypeName(Selection) = "Range" Then
        MsgBox "Please select a worksheet range and try again.", vbExclamation, _
            "No Range Selected"
    Else
        Set PPApp = GetObject(, "Powerpoint.Application")
        ' Reference active presentation
        Set PPPres = PPApp.ActivePresentation
        PPApp.ActiveWindow.ViewType = ppViewSlide
        ' Reference active slide
        Set PPSlide = PPPres.Slides(PPApp.ActiveWindow.Selection.SlideRange.SlideIndex)
     
        ' Copy the range as a picture
        Selection.Copy
        ' Paste the range
        PPSlide.Shapes.PasteSpecial(Link:=True).Select
     
        
        ' Clean up
        Set PPSlide = Nothing
        Set PPPres = Nothing
        Set PPApp = Nothing
    End If
End Sub

Sub ChartToPPTLink()
        Dim left As Double
        Dim top As Double
        Dim height As Double
        Dim width As Double
       
        ' Make sure a chart is selected
        
    If ActiveChart Is Nothing Then
        MsgBox "Please select a chart and try again.", vbExclamation, _
            "No Chart Selected"
    Else
        Set PPApp = GetObject(, "Powerpoint.Application")
        ' Reference active presentation
        Set PPPres = PPApp.ActivePresentation
        PPApp.ActiveWindow.ViewType = ppViewSlide
        ' Reference active slide
        Set PPSlide = PPPres.Slides _
            (PPApp.ActiveWindow.Selection.SlideRange.SlideIndex)
        
        ' Copy chart as a picture
        Worksheets(1).ChartObjects(1).Activate
        ActiveChart.ChartType = xlColumnStacked
        ActiveChart.HasTitle = True
        ActiveChart.ChartArea.Copy
        ' Paste chart link
        PPSlide.Shapes.PasteSpecial(Link:=True).Select
        
   
  
  'Objects name and size
        'Set shp = Worksheets(1).ChartObjects(1).ActiveChart.Shapes
        'shp.name
        
        'ActiveChart.ChartArea.Shapes("chrtName").left = 46.71
        'ActiveChart.ChartArea.Shapes("chrtName").width = 851.45
        'ActiveChart.ChartArea.Shapes("chrtName").top = 113.99
        'ActiveChart.ChartArea.Shapes("chrtName").height = 347.71
   
        
        ' Clean up
        Set PPSlide = Nothing
        Set PPPres = Nothing
        Set PPApp = Nothing
    End If
End Sub

Sub ChartsToPowerPoint()

    Dim pptApp As PowerPoint.Application
    Dim pptPres As PowerPoint.Presentation
    Dim pptSlide As PowerPoint.Slide
    Dim objChart As Chart

    'Open PowerPoint and create an invisible new presentation.
    Set pptApp = New PowerPoint.Application
    Set pptPres = pptApp.Presentations.Add(msoFalse)

        Worksheets(1).ChartObjects(1).Activate
        ActiveChart.ChartType = xlColumnStacked
        ActiveChart.HasTitle = True
        ActiveChart.CopyPicture
    

'Pause the application for ONE SECOND
Application.Wait Now + #12:00:01 AM#


    Set pptSlide = pptPres.Slides.Add(1, ppLayoutBlank)
    pptSlide.Shapes.PasteSpecial DataType:=ppPasteDefault, Link:=msoFalse

    'Save Images as png
    Path = CreateObject("Wscript.Shell").SpecialFolders("Desktop") & "\"

    For j = 1 To pptSlide.Shapes.Count
        With pptSlide.Shapes(j)
        .Export Path & j & ".png", ppShapeFormatPNG
        End With
    Next j

    pptApp.Quit

    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing

End Sub

