Function numdelta_cuo(TypeOption, s, K, r, q, sigma, T, H, eps)
numdelta_cuo = (cuo_fran(TypeOption, s + eps, K, r, q, sigma, T, H) - _
    cuo_fran(TypeOption, s - eps, K, r, q, sigma, T, H)) / (2 * eps)
End Function

Function cuo_fran(TypeOption, s, K, r, q, sigma, T, H)
    If s >= H Then
        cuo_fran = 0
        Exit Function
    End If
    If T <= 0 Then
        cuo_fran = 0
        Exit Function
    End If
    A = (H / s) ^ ((2 * (r - q)) / (sigma ^ 2) - 1)
    b = (H / s) ^ ((2 * (r - q)) / (sigma ^ 2) + 1)
    
    z1 = (Log(s / K) + (r - q + (sigma ^ 2) / 2) * T) / (sigma * Sqr(T))
    z2 = z1 - sigma * Sqr(T)
    z3 = (Log(s / H) + (r - q + (sigma ^ 2) / 2) * T) / (sigma * Sqr(T))
    z4 = z3 - sigma * Sqr(T)
    z5 = (Log(s / H) - (r - q - (sigma ^ 2) / 2) * T) / (sigma * Sqr(T))
    z6 = z5 - sigma * Sqr(T)
    z7 = (Log(s * K / H ^ 2) - (r - q - (sigma ^ 2) / 2) * T) / (sigma * Sqr(T))
    z8 = z7 - sigma * Sqr(T)
   
    cuo_fran = s * Exp(-q * T) * (myN(z1) - myN(z3) - b * (myN(z6) - myN(z8))) - _
        K * Exp(-r * T) * (myN(z2) - myN(z4) - A * (myN(z5) - myN(z7)))
End Function

Function Stat_hedge_cuo(typeop, s, K, r, q, sigma, T, H, N)
dt = T / (N - 1)
ReDim NB_Van(N), BS_date(N), Str_Van(N), Ech_Van(N) As Double
For i = 1 To N
NB_Van(N) = 0
BS_date(N) = 0
Next i
NB_Van(1) = 1
Str_Van(1) = K
Ech_Van(1) = T
For i = 2 To N
    datev = T - (i - 1) * dt
    Str_Van(i) = H
    Ech_Van(i) = T + (2 - i) * dt
    For j = 1 To i - 1
        BS_date(j) = myBS(typeop, H, Str_Van(j), r, q, sigma, Ech_Van(j) - datev)
    Next j
    Ptf_val = WorksheetFunction.SumProduct(NB_Van, BS_date)
    BS_new = myBS(typeop, H, Str_Van(i), r, q, sigma, Ech_Van(i) - datev)
    NB_Van(i) = -Ptf_val / BS_new
Next i
For j = 1 To N
    BS_date(j) = myBS(typeop, s, Str_Van(j), r, q, sigma, Ech_Van(j))
Next j
Stat_hedge_cuo = WorksheetFunction.SumProduct(NB_Van, BS_date)
End Function
