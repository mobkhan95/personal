Attribute VB_Name = "Module2"
Global Const Pi = 3.14159265358979
Public m As Integer
Public mu_max As Integer
Public del_tau As Double
Public del_x As Double
Public lambda As Double
Public kappa As Double
Public alpha As Double
Public beta As Double
Public bsprice As Double
Public A As Double
Public B As Double
Public pubflag As Integer
Public iflag As Integer


'// The normal distribution function
Public Function ND(X As Double) As Double
    ND = 1 / Sqr(2 * Pi) * Exp(-X ^ 2 / 2)
End Function

'// The cumulative normal distribution function
Public Function CND(X As Double) As Double
    
    Dim L As Double, K As Double
    Const a1 = 0.31938153:  Const a2 = -0.356563782: Const a3 = 1.781477937:
    Const a4 = -1.821255978:  Const a5 = 1.330274429
    
    L = Abs(X)
    K = 1 / (1 + 0.2316419 * L)
    CND = 1 - 1 / Sqr(2 * Pi) * Exp(-L ^ 2 / 2) * (a1 * K + a2 * K ^ 2 + a3 * K ^ 3 + a4 * K ^ 4 + a5 * K ^ 5)
    
    If X < 0 Then
        CND = 1 - CND
    End If
End Function
'// Black Scholes Options Prices for Call Options on Currency
'// Function arguments are (spot, strike, start date, expiration date, domestic interest rate, volatility, foreign interest rate)
Public Function BlackScholes(s As Double, K _
                As Double, T As Date, E As Date, d As Double, V As Double, f As Double) As Double
                
    Dim d1 As Double, d2 As Double, tau As Double
    
    tau = (E - T) / 365
    d1 = (Log(s / K) + ((d - f + ((V ^ 2) / 2)) * tau)) / (V * Sqr(tau))
    d2 = d1 - (V * Sqr(tau))
    BlackScholes = (Exp(-f * tau) * s * CND(d1)) - (K * Exp(-d * tau) * CND(d2))
End Function
Public Function Calculate(theta As Double, Optional flag As Integer = 0) As Double

Dim spot As Double, dom As Double, foreign As Double, vol As Double, strike As Double, today As Date, expiry As Date
Dim A As Double, B As Double
pubflag = Range("b28").Value
iflag = Range("b38").Value
spot = Range("b1").Value
dom = Range("b2").Value
foreign = Range("b3").Value
vol = Range("b4").Value
strike = Range("b5").Value
today = Range("b6").Value
expiry = Range("b7").Value
bsprice = BlackScholes(spot, strike, today, expiry, dom, vol, foreign)
m = Range("B9").Value
mu_max = Range("B10").Value
A = Range("b12").Value
Range("b20").Value = lambda
kappa = ((2 * (dom - foreign)) / (vol * vol))
alpha = (-1 * 0.5 * (kappa - 1))
beta = ((-0.25 * (kappa - 1) * (kappa - 1)) - ((2 * dom) / (vol * vol)))


del_tau = (0.5 * vol * vol * ((expiry - today) / 365)) / mu_max
B = A - (A * m / (m / 2)) + (m / (m / 2) * Log(spot / strike))
del_x = (B - A) / m
lambda = (del_tau / (del_x * del_x))




Dim Ar As Variant
ReDim Ar(1 To (m - 1), 1 To (m - 1)) As Double
Dim Br As Variant
ReDim Br(1 To (m - 1), 1 To (m - 1)) As Double
Dim w As Variant
ReDim w(1 To (m - 1), 0 To mu_max) As Double
Dim d As Variant
ReDim d(1 To (m - 1), 0 To (mu_max - 1)) As Double
Dim underlying As Variant
ReDim underlying(0 To m, 0 To mu_max, 1 To 2) As Double
'Dim punderlying As Variant
'ReDim punderlying(0 To m, 0 To mu_max) As Double

For i = 0 To m
    For j = 0 To mu_max
        underlying(i, j, 1) = A + (i * del_x)
        underlying(i, j, 2) = j * del_tau
        'punderlying(i, j) = underlying(i, j, 2)
        Next j
    Next i

If flag = 0 Then
Range("b20").Value = lambda
Range("b14").Value = Exp(underlying(0, 0, 1)) * strike
Range("b15").Value = Exp(underlying(m, 0, 1)) * strike
End If

For i = 1 To (m - 1)
    Ar(i, i) = (2 * theta * lambda) + 1
    If i < (m - 1) Then
        Ar(i, i + 1) = -1 * theta * lambda
    End If
    If i > 1 Then
        Ar(i, i - 1) = -1 * theta * lambda
    End If
    Br(i, i) = (1 - (2 * lambda * (1 - theta)))
    If i < (m - 1) Then
        Br(i, i + 1) = (1 - theta) * lambda
    End If
    If i > 1 Then
        Br(i, i - 1) = (1 - theta) * lambda
    End If
Next i



d = BC(d, underlying)
w = IC(w, underlying)
w = Ops(w, underlying, Ar, Br, d)

Dim wprime As Variant
ReDim wprime(0 To m, 0 To mu_max) As Double
Dim V As Variant
ReDim V(0 To m, 0 To mu_max) As Double

For j = 0 To mu_max
    For i = 1 To (m - 1)
        wprime(i, j) = w(i, j)
    Next i
Next j

wprime(0, 0) = Exp(-1 * alpha * underlying(0, 0, 1)) * Application.Max(Exp(underlying(0, 0, 1)) - 1, 0)
wprime(m, 0) = Exp(-1 * alpha * underlying(m, 0, 1)) * Application.Max(Exp(underlying(m, 0, 1)) - 1, 0)

For j = 1 To mu_max
    wprime(0, j) = lowerBC(underlying, j)
    wprime(m, j) = upperBC(underlying, j)
Next j


If flag = 0 Then

    Dim stri As String
    stri = Range("G2").Offset(m, mu_max).Address
    Range("G2:" & stri) = wprime
    Dim N As Variant, N2 As Variant
    For j = LBound(underlying, 1) To UBound(underlying, 1)
        N2 = Cells(Rows.Count, "F").End(xlUp).Row + 1
        If iflag = 1 Then
            Cells(N2, "F") = underlying(j, 0, 1)
        ElseIf iflag = 2 Then
            Cells(N2, "F") = j
        End If
        Next j
    Range("G1") = underlying(0, 0, 2)
    For i = (LBound(underlying, 2) + 1) To UBound(underlying, 2)
        N = Cells(1, Columns.Count).End(xlToLeft).Column + 1
        If iflag = 1 Then
            Cells(1, N) = underlying(0, i, 2)
        ElseIf iflag = 2 Then
            Cells(1, N) = i
        End If
        Next i
    Dim taus As Variant
    Dim xs As Variant
    taus = Range("G1:" & Cells(1, Cells(1, Columns.Count).End(xlToLeft).Column).Address)
    xs = Range("F2:F" & CStr(Cells(Rows.Count, "F").End(xlUp).Row))
    
    If iflag = 1 Then
        For i = LBound(taus, 2) To UBound(taus, 2)
            taus(1, i) = ((expiry - today) / 365) - (2 * taus(1, i) / (vol * vol))
            Next i
        For i = LBound(xs) To UBound(xs)
            xs(i, 1) = strike * Exp(xs(i, 1))
            Next i
    End If
End If

For j = 0 To mu_max
    For i = 0 To m
        V(i, j) = strike * Exp((alpha * underlying(i, j, 1)) + (beta * underlying(i, j, 2))) * wprime(i, j)
    Next i
Next j

If flag = 0 Then
    Range(Range("F" & CStr(Cells(Rows.Count, "F").End(xlUp).Row + 5)).Resize(m + 1, 1).Address) = xs
    Range(Range("F" & CStr(Cells(Rows.Count, "F").End(xlUp).Row - m)).Offset(-1, 1).Resize(1, mu_max + 1).Address) = taus
    Range(Range("G" & Cells(Rows.Count, "G").End(xlUp).Row + 1).Resize(m + 1, mu_max + 1).Address) = V
End If
Calculate = V(m / 2, mu_max)
'Range("o24") = "PDE Price"
'Range("P24") = V(m / 2, mu_max)
'Range("p25") = A
'Range("O25") = "X0"
'Range("p23") = lambda
'Range("o23") = "Lambda"
End Function
Sub Clear()
Range(Range("C1"), Cells(Rows.Count, Columns.Count)).ClearContents
Dim chtObj As ChartObject
    For Each chtObj In ActiveSheet.ChartObjects
        chtObj.Delete
    Next
    
End Sub
Sub Runner()

Range(Range("F1"), Cells(Rows.Count, Columns.Count)).ClearContents
Dim chtObj As ChartObject
    For Each chtObj In ActiveSheet.ChartObjects
        chtObj.Delete
    Next
    
Range("B19") = Calculate(Range("B11").Value)
End Sub
Public Function lowerBC(underlying As Variant, ByVal mu As Integer) As Double
lowerBC = 0
End Function
Public Function upperBC(underlying As Variant, ByVal mu As Integer) As Double
If pubflag = 1 Then
    upperBC = Exp((0.5 * (kappa + 1) * underlying(m, mu, 1)) + (0.25 * (kappa + 1) * (kappa + 1) * underlying(m, mu, 2))) - Exp((0.5 * (kappa - 1) * underlying(m, mu, 1)) + (0.25 * (kappa - 1) * (kappa - 1) * underlying(m, mu, 2)))
End If
If pubflag = 2 Then
    upperBC = Exp((0.5 * (kappa + 1) * underlying(m, mu, 1)) + (0.25 * (kappa + 1) * (kappa + 1) * underlying(m, mu, 2)))
End If
End Function

Public Function BC(d As Variant, underlying As Variant) As Variant
For q = 1 To (mu_max - 1)
    d(1, q) = (theta * lambda * lowerBC(underlying, q + 1)) + ((1 - theta) * lambda * lowerBC(underlying, q))
    d(m - 1, q) = (theta * lambda * upperBC(underlying, q + 1)) + ((1 - theta) * lambda * upperBC(underlying, q))
Next q

d(1, 0) = (theta * lambda * lowerBC(underlying, 1)) + (((1 - theta) * lambda * (Exp(-1 * alpha * underlying(0, 0, 1)) * Application.Max(Exp(underlying(0, 0, 1)) - 1, 0))))
d(m - 1, 0) = (theta * lambda * upperBC(underlying, 1)) + (((1 - theta) * lambda * (Exp(-1 * alpha * underlying(m, 0, 1)) * Application.Max(Exp(underlying(m, 0, 1)) - 1, 0))))

BC = d
End Function
Public Function IC(w As Variant, underlying As Variant) As Variant
For i = 1 To (m - 1)
    w(i, 0) = Exp(-1 * alpha * underlying(i, 0, 1)) * Application.Max(Exp(underlying(i, 0, 1)) - 1, 0)
Next i
IC = w
End Function

Public Function Ops(w As Variant, underlying As Variant, Ar As Variant, Br As Variant, d As Variant) As Variant
Dim C As Variant, E As Variant, Ainv As Variant
Ainv = WorksheetFunction.MInverse(Ar)
C = WorksheetFunction.MMult(WorksheetFunction.MInverse(Ar), Br)
For j = 1 To mu_max
    E = Mult(Ainv, d, j - 1)
    For i = 1 To (m - 1)
        w(i, j) = Mult(C, w, j - 1)(i) + E(i)
    Next i
Next j
Ops = w
End Function

Public Function Mult(arr1 As Variant, arr2 As Variant, ByVal colnum As Variant) As Variant
Dim arr3 As Variant
ReDim arr3(1 To (m - 1)) As Double
For i = LBound(arr1, 1) To UBound(arr1, 2)
    For j = LBound(arr1, 2) To UBound(arr1, 2)
        arr3(i) = arr3(i) + arr1(i, j) * arr2(j, colnum)
    Next j
Next i
Mult = arr3
End Function
Sub Graph1()
    Dim chtObj As ChartObject
    For Each chtObj In ActiveSheet.ChartObjects
        chtObj.Delete
    Next
    Range("C1:D200").ClearContents
    Dim start As Double, endspot As Double, increment As Double, counter As Integer
    start = -1
    endspot = 1
    increment = 0.02
    
    counter = 1
    For i = start To endspot + increment Step increment
    Cells(counter, 3).Value = i
    Cells(counter, 4).Value = Application.Run("Calculate", i, 1) - bsprice
    counter = counter + 1
    Next i
    
    Dim xaxis As Range
    Dim yaxis As Range
    Set xaxis = Range("$C$2", Range("$C$2").End(xlDown))
    Set yaxis = Range("$D$2", Range("$D$2").End(xlDown))
    
    Dim c1 As Chart
    Dim co As ChartObject
    Set co = ActiveWorkbook.Sheets("Sheet1").ChartObjects.Add(510, 110, 600, 400)
    Set c1 = co.Chart
    With c1
        .ChartType = xlXYScatterLines
        .HasLegend = False
        .HasTitle = True
        .ChartTitle.Text = "Discrepancy vs. Theta"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Theta Values"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Discrepancy"
        .Axes(xlCategory).MinimumScale = start
        .Axes(xlCategory).MaximumScale = endspot
    End With
    
    Dim s As Series
    Set s = c1.SeriesCollection.NewSeries
    With s
        .Values = yaxis
        .XValues = xaxis
    End With
    
End Sub
