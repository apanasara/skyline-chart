Attribute VB_Name = "SkyLineChart"
Private Sub Replot_Click()
    
    Dim pv As PivotTable
    
    Lrw = Sheets("Data").UsedRange.Rows.Count
    Set pv = Sheets("Pivot").PivotTables("PivotTable1")
    
    pv.ChangePivotCache ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Data!R1C1:R" & Lrw & "C5", Version:=6)
        
    pv.DataLabelRange.Cells(2).Group Start:=True, End:=True, Periods:=Array(False, False, False, False, True, False, True)
        
    Call chart_formater
End Sub

Sub chart_formater()
'
' chart_formater Macro
'
    Application.ScreenUpdating = False
    
    Dim label() As String
    'Dim prcnt As Long
    
    total_tags = Sheets("Skyline").ChartObjects("Skyline").Chart.FullSeriesCollection.Count
    For i = 1 To total_tags
    Set Tag = Sheets("Skyline").ChartObjects("Skyline").Chart.FullSeriesCollection(i)
    
    '---------black outline----------

    With Tag.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Weight = 1
    End With
    
   '---------filling colour-----------

   label = Split(Tag.Name, Chr(10))
   
    If label(2) = "0" Then '---yet not intiated
        Tag.Format.Fill.ForeColor.RGB = RGB(255, 50, 0)
    
    
    ElseIf label(2) = label(1) Then '---completed
        Tag.Format.Fill.ForeColor.RGB = RGB(130, 190, 90)
    
    
    Else '---under progress
    
    prcnt = 1 - (Int(label(2)) / Int(label(1)))
    Tag.Select
    With Tag.Format.Fill
        .TwoColorGradient msoGradientHorizontal, 1
        .GradientAngle = 90
        
        .ForeColor.RGB = RGB(255, 190, 0)
        .BackColor.RGB = RGB(255, 255, 255)
            
        .GradientStops(1).Color = RGB(255, 255, 255)
        .GradientStops(1).Position = 0
        
        .GradientStops(2).Color = RGB(255, 255, 255)
        .GradientStops(2).Position = prcnt
        
        If .GradientStops.Count < 3 Then: .GradientStops.Insert RGB(255, 190, 0), prcnt
        .GradientStops(3).Color = RGB(255, 190, 0)
        .GradientStops(3).Position = prcnt
        
        If .GradientStops.Count < 4 Then: .GradientStops.Insert RGB(255, 190, 0), 1
        .GradientStops(4).Color = RGB(255, 190, 0)
        .GradientStops(4).Position = 1
        
    End With
    End If
    
    
    '----------Labeling----------------
    Tag.HasDataLabels = True
    Tag.DataLabels.ShowValue = False
    Tag.DataLabels.ShowSeriesName = True
    Next
    
    Application.ScreenUpdating = True
End Sub

