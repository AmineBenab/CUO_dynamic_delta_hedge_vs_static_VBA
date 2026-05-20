
Option Base 1

Function myHedge_CUO(s As Double, _
                    K As Double, _
                    H As Double, _
                    r As Double, _
                    q As Double, _
                    sigma As Double, _
                    mu As Double, _
                    T As Double, _
                    nbTraj As Long, _
                    nbSteps As Long) As Variant

Dim eps As Double
Dim dt As Double
Dim UOC_Price As Double



ReDim Results(nbTraj, 3) As Variant

eps = 0.01
dt = T / nbSteps
UOC_Price = cuo_fran("call", s, K, r, q, sigma, T, H)
If UOC_Price < 0.01 Then
    myHedge_CUO = Results
    Exit Function
End If

Dim i As Long, j As Long
Dim Z As Double
Dim S_t As Double
Dim HedgeAccount As Double
Dim Delta_old As Double, Delta_new As Double
Dim tleft As Double
Dim Payoff As Double
Dim isKnocked As Boolean
Dim SharesHeld As Double

Randomize
For i = 1 To nbTraj
    S_t = s
    HedgeAccount = UOC_Price
    isKnocked = False
    tleft = T
    Delta_old = numdelta_cuo("call", s, K, r, q, sigma, T, H, eps)
    HedgeAccount = HedgeAccount - Delta_old * S_t
    SharesHeld = Delta_old
    For j = 1 To nbSteps
        HedgeAccount = HedgeAccount * Exp(dt * r)
        Z = Application.NormSInv(Rnd())
        S_t = S_t * Exp((mu - 0.5 * sigma ^ 2) * dt + sigma * Sqr(dt) * Z)
        If S_t > H Then
            HedgeAccount = HedgeAccount + SharesHeld * S_t
            SharesHeld = 0
            Results(i, 1) = HedgeAccount
            isKnocked = True
            Exit For
        Else
            tleft = tleft - dt
            Delta_new = numdelta_cuo("call", S_t, K, r, q, sigma, tleft, H, eps)
            SharesHeld = SharesHeld + (Delta_new - Delta_old)
            HedgeAccount = HedgeAccount - (Delta_new - Delta_old) * S_t
            Delta_old = Delta_new
        End If
    Next j
    If isKnocked = False Then
        HedgeAccount = HedgeAccount + SharesHeld * S_t
        Payoff = Application.Max(S_t - K, 0)
        Results(i, 1) = HedgeAccount - Payoff
    End If
    Results(i, 2) = Results(i, 1) / UOC_Price
    Results(i, 3) = isKnocked


Next i

myHedge_CUO = Results

    

End Function

Sub Scenarios_UOC()
Run Simulation
Call CreateChartSheet

Worksheets("dynDeltaHedge").Activate
Dim K As Double, H As Double
Dim r As Double, q As Double, sigma As Double, mu As Double
Dim nbTraj As Long, nbSteps As Long
    
K = Range("C7").Value
r = Range("C9").Value
mu = Range("C10").Value
sigma = Range("C11").Value
q = Range("C12").Value
H = Range("G6").Value
nbTraj = Range("G4").Value
nbSteps = Range("G5").Value

Dim S_vals(5) As Double
Dim T_vals(5) As Double

S_vals(1) = 85
S_vals(2) = 95
S_vals(3) = 100
S_vals(4) = 105
S_vals(5) = 115

T_vals(1) = 0.1
T_vals(2) = 0.25
T_vals(3) = 0.5
T_vals(4) = 0.75
T_vals(5) = 1

Dim wSheet As Worksheet
Application.DisplayAlerts = False
For Each wSheet In ThisWorkbook.Worksheets
    If wSheet.Name = "UOC_Results" Then
        wSheet.Delete
        Exit For
    End If
Next wSheet

Dim wsResults As Worksheet
Set wsResults = ThisWorkbook.Sheets.Add
wsResults.Name = "UOC_Results"

Dim wsSummary As Worksheet
For Each wSheet In ThisWorkbook.Worksheets
    If wSheet.Name = "UOC_Summary" Then
        wSheet.Delete
        Exit For
    End If
Next wSheet
Application.DisplayAlerts = True

Set wsSummary = ThisWorkbook.Sheets.Add
wsSummary.Name = "UOC_Summary"
wsSummary.Cells(1, 1).Value = "Scenario"
wsSummary.Cells(1, 2).Value = "UOC_Price"
wsSummary.Cells(1, 3).Value = "Mean(NormCost)"
wsSummary.Cells(1, 4).Value = "Hull Ratio"
wsSummary.Cells(1, 5).Value = "StaticCost/UOC_Price"
wsSummary.Cells(1, 6).Value = "Pct_Knocked"
wsSummary.Cells(1, 7).Value = "P95(NormCost)"
wsSummary.Columns("F").NumberFormat = "0.00%"

Dim Si As Integer, Ti As Integer
Dim ColStart As Integer
Dim result As Variant
Dim label As String
Dim i As Long
Dim StaticPrice As Double
Dim staticCost As Double
Dim UOC_Price As Double
Dim summaryRow As Integer
Dim sumNorm As Double, sumNorm2 As Double
Dim meanN As Double, stdN As Double
Dim nbKnocked As Long
Dim tempArr() As Double
Dim p95 As Double

summaryRow = 2
ColStart = 1
nbKnocked = 0

For Si = 1 To 5
    For Ti = 1 To 5
        label = "S=" & S_vals(Si) & ", T=" & T_vals(Ti)
        wsResults.Cells(1, ColStart).Value = label
        wsResults.Cells(2, ColStart).Value = "CostBrut"
        wsResults.Cells(2, ColStart + 1).Value = "NormCost"
        wsResults.Cells(2, ColStart + 3).Value = "StaticCost"
        wsResults.Cells(2, ColStart + 2).Value = "IsKnocked"
        
        result = myHedge_CUO(S_vals(Si), K, H, r, q, sigma, mu, T_vals(Ti), nbTraj, nbSteps)
        
        UOC_Price = cuo_fran("call", S_vals(Si), K, r, q, sigma, T_vals(Ti), H)
        StaticPrice = Stat_hedge_cuo("Call", S_vals(Si), K, r, q, sigma, T_vals(Ti), H, nbSteps)
        staticCost = StaticPrice - UOC_Price
        wsResults.Cells(3, ColStart + 3).Value = staticCost
        
        sumNorm = 0
        sumNorm2 = 0
        nbKnocked = 0
        
        ReDim tempArr(1 To nbTraj)
        
        For i = 1 To nbTraj
            sumNorm = sumNorm + result(i, 2)
            sumNorm2 = sumNorm2 + result(i, 2) ^ 2
            tempArr(i) = result(i, 2)
            wsResults.Cells(i + 2, ColStart).Value = result(i, 1)
            wsResults.Cells(i + 2, ColStart + 1).Value = result(i, 2)
            wsResults.Cells(i + 2, ColStart + 2).Value = result(i, 3)
            If result(i, 3) = True Then nbKnocked = nbKnocked + 1
        Next i
        
    meanN = sumNorm / nbTraj
    stdN = Sqr(sumNorm2 / nbTraj - meanN ^ 2)
    p95 = WorksheetFunction.Percentile(tempArr, 0.95)

    wsSummary.Cells(summaryRow, 1).Value = label
    wsSummary.Cells(summaryRow, 2).Value = UOC_Price
    wsSummary.Cells(summaryRow, 3).Value = meanN
    wsSummary.Cells(summaryRow, 4).Value = stdN
    wsSummary.Cells(summaryRow, 5).Value = staticCost / UOC_Price
    wsSummary.Cells(summaryRow, 6).Value = nbKnocked / nbTraj
    wsSummary.Cells(summaryRow, 7).Value = p95
    summaryRow = summaryRow + 1
    
    ColStart = ColStart + 5
    Next Ti
Next Si

wsSummary.Columns("A:G").AutoFit

Call RefreshChart

End Sub
