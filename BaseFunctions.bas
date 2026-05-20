Option Explicit
Option Base 0
Dim d1 As Double


Function myN(x)

myN = WorksheetFunction.Norm_S_Dist(x, True)

End Function

Function mySwitch(TypeOption)
    TypeOption = UCase(TypeOption) 'we rewrite on the same variable (but transformed into upper case)
    If TypeOption = "C" Or TypeOption = "CALL" Then
        mySwitch = 1
    ElseIf TypeOption = "P" Or TypeOption = "PUT" Then
        mySwitch = -1
    Else
        mySwitch = "argument has to be call or put"
    End If
End Function

Function myBS(typeop, s, K, r, div, sigma, T)
    Dim d1 As Double, d2 As Double
    Dim Z As Variant
        
    d1 = (Log(s / K) + (r - div + 0.5 * sigma ^ 2) * T) / _
        (sigma * Sqr(T))
    d2 = d1 - sigma * Sqr(T)
    Z = mySwitch(typeop)
    
    If Z = 1 Or Z = -1 Then
    
        myBS = Z * (s * Exp(-div * T) * myN(Z * d1) - _
            K * Exp(-r * T) * myN(Z * d2))
    Else
        myBS = "error in typeOption"
    End If
End Function


Function myDelta(TypeOption, s, K, r, div, sigma, T)

Dim d1 As Double

d1 = (Log(s / K) + (r - div + 0.5 * sigma ^ 2) * T) / _
    (sigma * Sqr(T))
myDelta = mySwitch(TypeOption) * Exp(-div * T) * myN(mySwitch(TypeOption) * d1)

End Function

Function myGamma(s, K, r, div, sigma, T)
Dim d1 As Double

d1 = (Log(s / K) + (r - div + 0.5 * sigma ^ 2) * T) / _
    (sigma * Sqr(T))
myGamma = WorksheetFunction.Norm_S_Dist(d1, False) * Exp(-div * T) / _
    (s * sigma * Sqr(T))

End Function

Function myTheta(TypeOption, s, K, r, div, sigma, T)
    Dim d1 As Double, d2 As Double
    Dim Z As Integer
    
    d1 = (Log(s / K) + (r - div + 0.5 * sigma ^ 2) * T) / _
    (sigma * Sqr(T))
    d2 = d1 - sigma * Sqr(T)
    Z = mySwitch(TypeOption)
    
myTheta = -s * WorksheetFunction.Norm_S_Dist(d1, False) * sigma * Exp(-div * T) / _
    (2 * Sqr(T)) + Z * div * s * myN(Z * d1) * Exp(-div * T) - _
    Z * r * K * Exp(-r * T) * myN(Z * d2)
End Function

Function myVega(s, K, r, div, sigma, T)

Dim d1 As Double
d1 = (Log(s / K) + (r - div + 0.5 * sigma ^ 2) * T) / _
    (sigma * Sqr(T))
myVega = s * Sqr(T) * WorksheetFunction.Norm_S_Dist(d1, False) * Exp(-div * T)

End Function

Function myRho(TypeOption, s, K, r, div, sigma, T)
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(s / K) + (r - div + 0.5 * sigma ^ 2) * T) / _
    (sigma * Sqr(T))
    d2 = d1 - sigma * Sqr(T)
myRho = mySwitch(TypeOption) * K * T * Exp(-r * T) * myN(mySwitch(TypeOption) * d2)
End Function

Function myCRR(typeop, s, K, r, div, sigma, T, nbSteps)

Dim delta_t As Double
Dim u As Double, d As Double
Dim j As Integer '# of up moves
Dim i As Integer 'stands for the date
Dim p As Double

'step 0
'calculate u and d

delta_t = T / nbSteps
u = Exp(sigma * Sqr(delta_t))
d = 1 / u
'optional at this stage p =
p = (Exp(r * delta_t) - d) / (u - d)

'build Final prices
ReDim finalS(nbSteps + 1) As Double

For j = 0 To nbSteps

finalS(j) = s * u ^ j * d ^ (nbSteps - j)

Next

ReDim optVal(0 To nbSteps, 0 To nbSteps)

'step 2: option final payoffs (i = nbSteps)

For j = 0 To nbSteps 'j stands for # up moves

    optVal(nbSteps, j) = myMax(mySwitch(typeop) * (finalS(j) - K), 0)

Next

'step 3: working backwards in the tree

For i = nbSteps - 1 To 0 Step -1

    For j = 0 To i
        
        
    optVal(i, j) = Exp(-r * delta_t) * (p * optVal(i + 1, j + 1) + (1 - p) * optVal(i + 1, j))
    
    Next
    
Next

myCRR = optVal(0, 0)


End Function
