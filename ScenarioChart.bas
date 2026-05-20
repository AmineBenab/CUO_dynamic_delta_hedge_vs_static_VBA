Sub CreateChartSheet()

Dim wSheet As Worksheet
Application.DisplayAlerts = False
For Each wSheet In ThisWorkbook.Worksheets
    If wSheet.Name = "UOC_Chart" Then
        wSheet.Delete
        Exit For
    End If
Next wSheet
Application.DisplayAlerts = True

Dim wsChart As Worksheet
Set wsChart = ThisWorkbook.Sheets.Add
wsChart.Name = "UOC_Chart"

wsChart.Cells(1, 1).Value = "Choose S and T"
wsChart.Cells(2, 1).Value = "S ="
wsChart.Cells(2, 2).Value = 100
wsChart.Cells(3, 1).Value = "T ="
wsChart.Cells(3, 2).Value = 0.5

Dim btn As Button
Set btn = wsChart.Buttons.Add(10, 100, 120, 30)
btn.Caption = "Refresh Chart"
btn.OnAction = "'" & ThisWorkbook.Name & "'!RefreshChart"

End Sub

Sub RefreshChart()

Set wb = ThisWorkbook
Dim wsChart As Worksheet
Dim wsResults As Worksheet

Set wsChart = wb.Worksheets("UOC_Chart")
Set wsResults = wb.Worksheets("UOC_Results")

Dim S_sel As Double
Dim T_sel As Double
S_sel = wsChart.Cells(2, 2).Value
T_sel = wsChart.Cells(3, 2).Value
    
Dim targetLabel As String
targetLabel = "S=" & S_sel & ", T=" & T_sel

Dim colFound As Integer
colFound = 0
Dim c As Integer
For c = 1 To 121
    If wsResults.Cells(1, c).Value = targetLabel Then
        colFound = c
        Exit For
    End If
Next c
    
If colFound = 0 Then
    MsgBox "Scenario not found, change S or T "
    Exit Sub
End If

nbTraj = ThisWorkbook.Worksheets("dynDeltaHedge").Range("G4").Value

Dim i As Long
Dim normCosts() As Double
ReDim normCosts(1 To nbTraj)

For i = 1 To nbTraj
    normCosts(i) = wsResults.Cells(i + 2, colFound + 1).Value
Next i

Dim staticNorm As Double
staticNorm = wsResults.Cells(3, colFound + 3).Value

Dim cht As ChartObject
For Each cht In wsChart.ChartObjects
    cht.Delete
Next cht

Dim nbBins As Integer
nbBins = 30
Dim minVal As Double, maxVal As Double
minVal = Application.WorksheetFunction.Min(normCosts)
maxVal = Application.WorksheetFunction.Max(normCosts)

Dim binWidth As Double
binWidth = (maxVal - minVal) / nbBins

Dim binLabels(1 To 30) As String
Dim binCounts(1 To 30) As Long
Dim j As Integer
For j = 1 To nbBins
binLabels(j) = Format(minVal + (j - 1) * binWidth, "0.0")
    binCounts(j) = 0
Next j

For i = 1 To nbTraj
    j = Int((normCosts(i) - minVal) / binWidth) + 1
    If j > nbBins Then j = nbBins  ' dernier point tombe sur maxVal
    binCounts(j) = binCounts(j) + 1
Next i

Dim myChartObject As ChartObject
Dim myChart As Chart
Set myChartObject = wsChart.ChartObjects.Add(150, 10, 500, 300)
Set myChart = myChartObject.Chart

With myChart
    .Axes(xlCategory).TickLabels.NumberFormat = "0.00"
    .ChartType = xlColumnClustered
    .SeriesCollection.NewSeries
    With .SeriesCollection(1)
        .Name = "Delta Hedge , Red Band = Static)"
        .Values = binCounts
        .XValues = binLabels
    End With
    .ChartGroups(1).GapWidth = 0
    .HasTitle = True
    .ChartTitle.Text = "Distribution NormCost for " & targetLabel
End With

Dim staticBin As Integer
staticBin = Int((staticNorm - minVal) / binWidth) + 1
If staticBin < 1 Then staticBin = 1
If staticBin > nbBins Then staticBin = nbBins
myChart.SeriesCollection(1).Points(staticBin).Interior.Color = RGB(220, 50, 50)

myChart.HasLegend = False
wsChart.Cells(4, 1).Value = "Red Bar = Static hedge"

End Sub
