Attribute VB_Name = "Module2"
Option Explicit
'Description: Schedule the extraction of Cooper Price Data (once per day) and represent it by its trend over the last 24-36 months
'Date of first Release 0.4/03.03.2016
'Author: George Calin
'Date of Last Release: 03.03.2016

Sub Scheduler()
    'schedules the execution of the file to once every 24 hrs
    Dim dtScheduler As Date
    dtScheduler = Now + TimeValue("12:00:00")
    Application.OnTime dtScheduler, "Scheduler"
    Worksheets("Counter").Activate
    Range("A6").Value = Range("A6").Value + 1
    Call ExtractCooperPriceHistory
    'or viceversa  Application.OnTime dtScheduler, "ExtractCooperPriceHistory" and inside Sub ExtractCooperPriceHistory() place a
    'Call Scheduler at the end. this way it'll be never ending loop
End Sub

Sub ExtractCooperPriceHistory()
'connects to a live webpage http://www.westmetall.com/de/markdaten.php?action=show_table_average&field=DEL_high and extracts data from "Monatsdurchschnitte"
'which is processed in a chart and saved as an image file. Dashboard reads image file and prints it out on the screen

        Dim FinalRowCurrentYear As Variant
        Dim FinalRowLastYear As Variant
        Dim FinalRowPenultimateYear As Variant
        Dim FinalRowChartSheet As Variant
        Dim i, j, k As Integer
        Dim rgExp As Range
        
        'refresh data extracted from webpage
        ActiveWorkbook.RefreshAll
        
        Sheets("Source").Select
        Columns("B:B").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Columns("B:B").ColumnWidth = 25
        
        FinalRowCurrentYear = Range("A21").End(xlDown).Row
        'MsgBox FinalRowCurrentYear only for verification
        
        FinalRowLastYear = Range("A" & FinalRowCurrentYear + 6).End(xlDown).Row
        'MsgBox FinalRowLastYear only for verification
        
        FinalRowPenultimateYear = Range("A" & FinalRowLastYear + 6).End(xlDown).Row
        'MsgBox FinalRowPenultimateYear only for verification

        Range("B21").Select
        ActiveCell.FormulaR1C1 = "Monat Jahr"
        For i = 0 To FinalRowCurrentYear - 22
            Range("B" & 22 + i).Select
            ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-1],"" "",LEFT(R[" & -3 - i & "]C[-1],4))"   'ActiveCell.FormulaR1C1 = "=MyFunction(R[" & var1 & "]C[" & var2 & "])" mind how the variable are addressed in a R1C1 formula
            
        Next i
     
        Sheets("Source").Select
        Range("B" & FinalRowCurrentYear + 6).Select
        ActiveCell.FormulaR1C1 = "Monat Jahr"
        For j = 0 To 11
            Range("B" & FinalRowCurrentYear + 7 + j).Select
            ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-1],"" "",MID(R[" & -3 - j & "]C[-1],6,4))"
        Next j
        
        Sheets("Source").Select
        Range("B" & FinalRowLastYear + 6).Select
        ActiveCell.FormulaR1C1 = "Monat Jahr"
        For j = 0 To 11
            Range("B" & FinalRowLastYear + 7 + j).Select
            ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-1],"" "",MID(R[" & -3 - j & "]C[-1],11,4))"
        Next j
        
        'making sure that Worksheet('Chart") is empty
        Sheets("Chart").Cells.Delete
        
        
'        'filter the dataset with the goal to extract only the numbers
        
        
        With Sheets("Source")
                .AutoFilterMode = False
            With .Range("$A$21" & ":" & "$C$" & 300)
                 .AutoFilter Field:=1, Criteria1:=Array("April", "August", "Dezember", "Februar", "Januar", "Juli", "Juni", "Mai", "März", "November", "Oktober", "September"), Operator:=xlFilterValues
                 .AutoFilter Field:=2, Criteria1:="<>"
                 ActiveSheet.AutoFilter.Range.Copy
                 Sheets("Chart").Select
                 Range("A7").Select
                Sheets("Chart").Paste
             End With
            End With

    Columns("A:A").ColumnWidth = 20
    Columns("B:B").ColumnWidth = 20
    Columns("C:C").ColumnWidth = 10

    FinalRowChartSheet = Range("B7").End(xlDown).Row
    'Msgbox FinalRowChartSheet

    Range("B8" & ":" & "C" & FinalRowChartSheet).Select
    ActiveSheet.Shapes.AddChart.Select
    
    'set the name of the Chart to "Cooper" to reference it later
    ActiveChart.Parent.Name = "Cooper"
    
    ActiveChart.ChartType = xlLine
    ActiveChart.ChartStyle = 4
    ActiveChart.SetSourceData Source:=Range("Chart!" & "B8" & ":" & "$C$" & FinalRowChartSheet)
    ActiveChart.Legend.Select
    ActiveChart.SeriesCollection(1).Name = "=""obere Kupfer DEL-Notiz (in Euro per 100 kg)"""
     ActiveChart.Legend.Select
    Selection.Delete
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).ReversePlotOrder = True
    
    'change the font of the Chart's title to Verdana 10 points
    ActiveChart.ChartTitle.Select
    With Selection.Format.TextFrame2.TextRange.Font
        .NameComplexScript = "Verdana"
        .NameFarEast = "Verdana"
        .Name = "Verdana"
    End With
    Selection.Format.TextFrame2.TextRange.Font.Size = 10
    


    'set plotting area to dark gray color
    ActiveChart.PlotArea.Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.5
        .Solid
    End With
    
    'set the color of the Chart Area to gray
    With ActiveSheet.Shapes("Cooper").Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.5
        .Solid
    End With
    
    'modify the vertical axis in between 400eur and 700eur
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 0
    ActiveChart.Axes(xlValue).MinimumScale = 400
    ActiveChart.Axes(xlValue).MaximumScale = 700
    
    'set the Font Color of the Vertical Axis to white
     ActiveChart.Axes(xlValue).TickLabels.Font.Color = RGB(255, 255, 255)
     'set the Font Color of the Horizontal Axis to white
     ActiveChart.Axes(xlCategory).TickLabels.Font.Color = RGB(255, 255, 255)
     
     'set the Major Horizontal Gridlines to color white
     ActiveChart.Axes(xlValue).MajorGridlines.Select
      With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0.7   ' 0 minimum 1 maximum
    End With
     
    'set the series line to color white
    ActiveChart.SeriesCollection(1).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
    End With
    
    'set the width of the series line to 4pt
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 2
    End With
    
    'set a Shadow effect to the series line
     With Selection.Format.Shadow
        .Type = msoShadow25
        .Visible = msoTrue
        .Style = msoShadowStyleOuterShadow
        .Blur = 4
        .OffsetX = 6.1232339957E-17
        .OffsetY = 1
        .RotateWithShape = msoFalse
        .ForeColor.RGB = RGB(255, 192, 0)
        .Transparency = 0
        .Size = 100
    End With
    
    'set the color of the Horizontal Axis to white and its Width to 1.5pt
    ActiveChart.Axes(xlCategory).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Weight = 1.5
    End With
    
    With ActiveChart.ChartArea
        .Width = 705   'add for the precise clipping @1920x1080px
        .Height = 210  'add for precise clipping @1920x1080px
        .Left = 470
        .Top = 17
    End With
    
    
     ' take the GridLines off the Excel Window
      ActiveWindow.DisplayGridlines = False
      
    ' Set Range you want to export to file
    
    
    Set rgExp = Range("H2:V11")
    
    ''' Copy range as picture onto Clipboard
    rgExp.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    ''' Create an empty chart with exact size of range copied
    With ActiveSheet.ChartObjects.Add(Left:=rgExp.Left, Top:=rgExp.Top, _
    Width:=rgExp.Width, Height:=rgExp.Height)
    .Name = "CooperPrice"
    ' CATCH ***** this removes the border line of the temporary Chart Object *****
    .Border.LineStyle = xlNone
    ' ****************************************************************************************
    .Activate
    End With
    
    ''' Paste into chart area, export to file -comment or uncomment accordingly
    ActiveChart.Paste
    
    'ActiveSheet.ChartObjects("CooperPrice").Chart.Export "C:\Users\SomeUser\Desktop\CooperPrice.jpg" 'to use on own's machine
    ActiveSheet.ChartObjects("CooperPrice").Chart.Export "C:\inetpub\vhosts\somePath\CooperPrice.jpg"   'save image to directory of vhosts on SRVERP96
    

'        Sheets("Chart").Cells.Delete
'         Sheets("Chart").Select
'
         


End Sub


