Attribute VB_Name = "LotCalculation"
Option Explicit

Public Function CalculateIntercept(Y() As Double, x() As Double) As Double
    Dim n As Integer
    Dim avgX As Double
    Dim avgY As Double
    Dim diffProdSum As Double
    Dim squareDiffSum As Double
    

    Dim i As Integer

    n = UBound(x)
    For i = 1 To UBound(x)
        avgX = avgX + x(i)
        avgY = avgY + Y(i)
    Next i
    avgX = avgX / n
    avgY = avgY / n

    For i = 1 To UBound(x)
        diffProdSum = diffProdSum + (x(i) - avgX) * (Y(i) - avgY)
        squareDiffSum = squareDiffSum + (x(i) - avgX) ^ 2
    Next i

    Dim Slope As Double
    Slope = diffProdSum / squareDiffSum

    CalculateIntercept = avgY - Slope * avgX
End Function

Public Function CalculateSlope(Y() As Double, x() As Double) As Double
    Dim n As Integer
    Dim avgX As Double
    Dim avgY As Double
    Dim diffProdSum As Double
    Dim squareDiffSum As Double
    Dim i As Integer

    n = UBound(x)
    For i = 1 To UBound(x)
        avgX = avgX + x(i)
        avgY = avgY + Y(i)
    Next i
    avgX = avgX / n
    avgY = avgY / n

    For i = 1 To UBound(x)
        diffProdSum = diffProdSum + (x(i) - avgX) * (Y(i) - avgY)
        squareDiffSum = squareDiffSum + (x(i) - avgX) ^ 2
    Next i

    CalculateSlope = diffProdSum / squareDiffSum
End Function

Public Function CalculateSumXMY2(Y() As Double, Ycalc() As Double) As Double
    Dim Sum As Double
    Dim i As Integer

    If UBound(Y) <> UBound(Ycalc) Then
        Err.Raise NUMBER:=5, Description:="Array dimensions must match"
        Exit Function
    End If

    For i = 1 To UBound(Y)
        Sum = Sum + (Y(i) - Ycalc(i)) ^ 2
    Next i

    CalculateSumXMY2 = Sum
End Function



Public Function CalculateDevSq(Ycalc() As Double) As Double
    Dim Sum As Double
    Dim mean As Double
    Dim i As Integer
    Dim n As Integer

    n = UBound(Ycalc)

    ' Calculate the mean
    For i = 1 To UBound(Ycalc)
        mean = mean + Ycalc(i)
    Next i
    mean = mean / n

    ' Calculate the sum of squares of deviations
    For i = 1 To UBound(Ycalc)
        Sum = Sum + (Ycalc(i) - mean) ^ 2
    Next i

    CalculateDevSq = Sum
End Function

Public Function CalculateSTEYX(Y() As Double, x() As Double) As Double
    Dim n As Integer
    Dim avgX As Double
    Dim avgY As Double
    Dim diffProdSum As Double
    Dim squareDiffSumX As Double
    Dim squareDiffSumY As Double
    Dim i As Integer
    Dim SteYX As Double
    Dim Denom As Double
    
    On Error GoTo ERR_ST

    n = UBound(x)
    For i = 1 To UBound(x)
        avgX = avgX + x(i)
        avgY = avgY + Y(i)
    Next i
    avgX = avgX / n
    avgY = avgY / n

    For i = 1 To UBound(x)
        diffProdSum = diffProdSum + (x(i) - avgX) * (Y(i) - avgY)
        squareDiffSumX = squareDiffSumX + (x(i) - avgX) ^ 2
        squareDiffSumY = squareDiffSumY + (Y(i) - avgY) ^ 2
    Next i

    Dim Slope As Double
    Slope = diffProdSum / squareDiffSumX

    Dim devSq As Double
    devSq = squareDiffSumY - Slope * diffProdSum
ERR_END:
    On Error GoTo 0
    SteYX = Sqr(devSq / (n - 2))
    Denom = Sqr(Abs(squareDiffSumX - n * avgX ^ 2))
    SteYX = SteYX / Denom
    
    
    CalculateSTEYX = SteYX
    Exit Function
ERR_ST:
    MsgBox Err.Description
    Resume Next
    
    
End Function



'In this code, Gamma is a helper function that calculates the Gamma function, which is used in the calculation of the t-distribution.
'CalculateTDist is the function that calculates the t-distribution.
'It takes three arguments: x is the numeric value at which to evaluate the distribution,
'df is the number of degrees of freedom, and tails specifies the number of distribution tails to return.


Private Function Gamma(z As Double) As Double
    Dim tmp1 As Double, tmp2 As Double
    Dim k As Integer
    Dim a As Double
    Dim b(0 To 5) As Double
    Dim c(0 To 6) As Double

    b(0) = 76.18009173
    b(1) = -86.50532033
    b(2) = 24.01409822
    b(3) = -1.231739516
    b(4) = 0.00120858003
    b(5) = -0.00000536382

    tmp1 = z + 5.5
    tmp1 = (z + 0.5) * Log(tmp1) - tmp1
    tmp2 = 1.00000000019

    For k = 0 To 5
        tmp2 = tmp2 + b(k) / (z + k + 1)
    Next k

    a = tmp1 + Log(2.50662827465 * tmp2 / (z + 1))
    Gamma = Exp(a)
End Function

Public Function CalculateTDist(x As Double, df As Integer, tails As Integer) As Double
    Dim a As Double
    Dim b As Double
    Dim c As Double

    a = Gamma((df + 1) / 2) / (Sqr(df * 3.14159265358979) * Gamma(df / 2))
    b = 1 / (1 + (x ^ 2) / df)
    c = df / 2

    CalculateTDist = a * (b ^ c)

    If tails = 2 Then
        CalculateTDist = 2 * (1 - CalculateTDist)
    End If
End Function



Private Function Beta(a As Double, b As Double) As Double
    Beta = Gamma(a) * Gamma(b) / Gamma(a + b)
End Function

Private Function IncompleteBeta(x As Double, a As Double, b As Double) As Double
    Dim Sum As Double
    Dim term As Double
    Dim m As Integer

    term = x ^ a * (1 - x) ^ b / a / Beta(a, b)
    Sum = term

    m = 0
    Do While m <= 100
        term = term * (a + m) * (a + b + m) / (a + 2 * m + 1) / (m + 1) * x
        Sum = Sum + term
        m = m + 1
        If term = 0.0000000001 Then GoTo cont:
    Loop
cont:
    IncompleteBeta = Sum
End Function


'In this code, Beta is a helper function that calculates the Beta function,
'IncompleteBeta is a helper function that calculates the incomplete Beta function, and
'CalculateTInv is the function that calculates the inverse of the Student’s t-distribution.
'It takes two arguments: alpha is the probability associated with the two-tailed Student’s t-distribution,
'and df is the number of degrees of freedom.

Public Function CalculateTInv(alpha As Double, df As Integer) As Double
    Dim x As Double
    Dim Dx As Double
    Dim term As Double
    Dim i As Integer

    x = 0.5
    Dx = 0.5
    i = 0

    Do While i <= 100
        If IncompleteBeta(x, df / 2, 0.5) > alpha Then
            x = x - Dx
        Else
            x = x + Dx
        End If
        Dx = Dx / 2
        i = i + 1
        If term = 0.0000000001 Then GoTo cont:
    Loop
cont:
    CalculateTInv = Sqr(df * (1 - x) / x)
End Function








Function TINV(Probabilita As Double, GradiDiLiberta As Integer) As Double
'   Funzione TINV personalizzata in VB6
'   Calcola il valore t della distribuzione t di Student
'   Parametri:
'       Probabilita: la probabilità per cui si calcola il valore t
'       GradiDiLiberta: il numero di gradi di libertà
'   Restituisce: il valore t

    ' Valori di riferimento per l'approssimazione di Edgeworth
    Const a1 = 0.333333333
    Const a2 = -0.01986813
    Const a3 = 0.0001531948
    Const a4 = -0.0000007152
    Const b1 = 0.67449275
    Const b2 = 0.09625325
    Const b3 = 0.00793789
    Const b4 = 0.00025931

    ' Valore iniziale per l'iterazione
    Dim t0 As Double

    ' Se la probabilità è minore o uguale a 0, restituisce -infinito
    If Probabilita <= 0 Then
        TINV = 0
        Exit Function
    End If

    ' Se la probabilità è maggiore o uguale a 1, restituisce infinito
    If Probabilita >= 1 Then
        TINV = 0
        Exit Function
    End If

    ' Calcolo del valore iniziale per l'iterazione
    t0 = (b1 + Log(Probabilita / (1 - Probabilita))) / b2

    ' Ciclo di iterazione per il calcolo del valore t
    Do While True
        ' Approssimazione di Edgeworth
        Dim z As Double
        z = t0 + Atan2(Sqr(GradiDiLiberta * (t0 * t0 - 1)), GradiDiLiberta * t0)
        t0 = (z * (a1 + z * (a2 + z * (a3 + z * a4)))) / (b1 + z * (b2 + z * (b3 + z * b4)))

        ' Condizione di uscita dal ciclo
        If Abs(Probabilita - (0.5 + 0.5 * Erf(t0 / Sqr(GradiDiLiberta)))) < 0.0000001 Then Exit Do
    Loop

    ' Restituisce il valore t calcolato
    TINV = t0

End Function

Function Atan2(Y As Double, x As Double) As Double
'   Funzione Atan2 personalizzata in VB6
'   Calcola l'angolo arcotangente di due numeri
'   Parametri:
'       Y: il valore sull'asse y
'       X: il valore sull'asse x
'   Restituisce: l'angolo arcotangente in radianti

    Dim PI As Double
    
    PI = 2 * Atn(1)

    Dim angle As Double

    If x = 0 Then
        If Y > 0 Then
            angle = PI / 2
        Else
            angle = -PI / 2
        End If
    Else
        angle = Atan(Y / x)
        If x < 0 Then
            If angle >= 0 Then
                angle = angle - PI
            Else
                angle = angle + 2 * PI
            End If
        End If
    End If

    Atan2 = angle

End Function
Function Erf(x As Double) As Double
'   Funzione Erf personalizzata in VB6
'   Calcola la funzione errore di Gauss
'   Parametro:
'       X: il valore per cui si calcola la funzione Erf
'   Restituisce: il valore di Erf(X)

    ' Valori di coefficienti per l'approssimazione di Chebyshev
    Const a1 = 0.27609759562909
    Const a2 = 0.90563051201728
    Const a3 = -0.28444277929376
    Const a4 = -0.08106207180899
    Const b1 = 0.99942506785928
    Const b2 = -0.00338649034802
    Const b3 = 1.50933484762E-05
    Const b4 = -1.1631059605E-07

    Dim t As Double
    Dim P As Double

    t = 1# / (1# + a1 * x)
    P = (a2 + t * (a3 + t * a4)) / (b1 + t * (b2 + t * (b3 + t * b4)))

    Erf = x * P + (1# - P) / Exp(-x * x / 2#)

End Function
Function Atan(Numero As Double) As Double
'   Funzione Atan personalizzata in VB6
'   Calcola l'arco tangente di un numero
'   Parametro:
'       Numero: il numero per cui si calcola l'arco tangente
'   Restituisce: l'arco tangente in radianti

    Atan = Atan2(Numero, 1)
End Function


Public Function GetMax(Uncertainty() As Double) As Double
    ' Dichiarazione delle variabili
    Dim i As Integer
    Dim MaxValue As Double

    ' Imposta il valore massimo sul primo elemento dell'array'
    MaxValue = Uncertainty(0)

    ' Ciclo per scorrere l'array
    For i = 1 To UBound(Uncertainty)
        ' Confronta il valore corrente con il valore massimo finora trovato'
        If Uncertainty(i) > MaxValue Then
            MaxValue = Uncertainty(i)
        End If
    Next i

    ' Restituisce il valore massimo'
    GetMax = MaxValue
End Function
