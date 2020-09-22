Attribute VB_Name = "Module2"
Option Explicit

'app
'  pre
'   slide
'    shapes
'     text frame
'      text

Sub ExporttoPPT()


Dim ppt_app As New PowerPoint.Application
Dim pre As PowerPoint.Presentation
Dim slde As PowerPoint.Slide
Dim shp As PowerPoint.Shape
Dim wb As Workbook
Dim rng As Range

Dim vSheet$
Dim vRange$
Dim vWidth As Double
Dim vHeight As Double
Dim vTop As Double
Dim vLeft As Double
Dim vSlide_No As Long
Dim expRng As Range

Dim adminSh As Worksheet
Dim cofigRng As Range
Dim xlfile$
Dim pptfile$


Application.DisplayAlerts = False

Set adminSh = ThisWorkbook.Sheets("Admin")
Set cofigRng = adminSh.Range("Rng_sheets")

xlfile = adminSh.[excelPth]
pptfile = adminSh.[pptPth]
      
Set wb = Workbooks.Open(xlfile)
Set pre = ppt_app.Presentations.Open(pptfile)


For Each rng In cofigRng
   
   '----------------- set VARIABLES
   With adminSh
      vSheet$ = .Cells(rng.Row, 4).Value
      vRange$ = .Cells(rng.Row, 5).Value
      vWidth = .Cells(rng.Row, 6).Value
      vHeight = .Cells(rng.Row, 7).Value
      vTop = .Cells(rng.Row, 8).Value
      vLeft = .Cells(rng.Row, 9).Value
      vSlide_No = .Cells(rng.Row, 10).Value
   End With
   
   
   '----------------- EXPORT TO PPT
   
            wb.Activate
            Sheets(vSheet$).Activate
            Set expRng = Sheets(vSheet$).Range(vRange$)
            expRng.Copy
            
            Set slde = pre.Slides(vSlide_No)
            slde.Shapes.PasteSpecial ppPasteBitmap
            Set shp = slde.Shapes(1)
            
            With shp
               
               .top = vTop
               .left = vLeft
               .width = vWidth
               .height = vHeight
               
            End With
            
            
            Set shp = Nothing
            Set slde = Nothing
            Set expRng = Nothing
   
   Application.CutCopyMode = False
   Set expRng = Nothing
Next rng

pre.Save
pre.Close

Set pre = Nothing
Set ppt_app = Nothing

wb.Close False
Set wb = Nothing

Application.DisplayAlerts = True
 End Sub
