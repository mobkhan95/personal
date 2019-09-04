Attribute VB_Name = "Module1"
Global Const Pi = 3.14159265358979

Option Explicit
Option Compare Text

Option Base 1

                    
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
                As Double, T As Date, E As Date, d As Double, v As Double, f As Double) As Double
                
    Dim d1 As Double, d2 As Double, tau As Double
    
    tau = (E - T) / 365
    d1 = (Log(s / K) + ((d - f + ((v ^ 2) / 2)) * tau)) / (v * Sqr(tau))
    d2 = d1 - (v * Sqr(tau))
    BlackScholes = (Exp(-f * tau) * s * CND(d1)) - (K * Exp(-d * tau) * CND(d2))
End Function

Public Function Delta(s As Double, K _
                As Double, T As Date, E As Date, d As Double, v As Double, f As Double) As Double
                
    Dim d1 As Double, d2 As Double, tau As Double
    
    tau = (E - T) / 365
    d1 = (Log(s / K) + ((d - f + ((v ^ 2) / 2)) * tau)) / (v * Sqr(tau))
    d2 = d1 - (v * Sqr(tau))
    Delta = (Exp(-f * tau) * CND(d1))
End Function


Public Function Gamma(s As Double, K _
                As Double, T As Date, E As Date, d As Double, v As Double, f As Double) As Double
                
    Dim d1 As Double, d2 As Double, tau As Double
    
    tau = (E - T) / 365
    d1 = (Log(s / K) + ((d - f + ((v ^ 2) / 2)) * tau)) / (v * Sqr(tau))
    d2 = d1 - (v * Sqr(tau))
    Gamma = (Exp(-f * tau) * ND(d1)) / (s * v * Sqr(tau))
End Function

Public Function Vega(s As Double, K _
                As Double, T As Date, E As Date, d As Double, v As Double, f As Double) As Double
                
    Dim d1 As Double, d2 As Double, tau As Double
    
    tau = (E - T) / 365
    d1 = (Log(s / K) + ((d - f + ((v ^ 2) / 2)) * tau)) / (v * Sqr(tau))
    d2 = d1 - (v * Sqr(tau))
    Vega = (Exp(-f * tau) * ND(d1) * s * Sqr(tau))
End Function

Public Function Theta(s As Double, K _
                As Double, T As Date, E As Date, d As Double, v As Double, f As Double) As Double
                
    Dim d1 As Double, d2 As Double, tau As Double
    
    tau = (E - T) / 365
    d1 = (Log(s / K) + ((d - f + ((v ^ 2) / 2)) * tau)) / (v * Sqr(tau))
    d2 = d1 - (v * Sqr(tau))
    Theta = (-1 * ((Exp(-f * tau) * ND(d1) * s * v) / (2 * Sqr(tau)))) - (d * K * Exp(-d * tau) * CND(d2)) + (f * s * Exp(-f * tau) * CND(d1))
End Function

Public Function Phi(s As Double, K _
                As Double, T As Date, E As Date, d As Double, v As Double, f As Double) As Double
                
    Dim d1 As Double, d2 As Double, tau As Double
    
    tau = (E - T) / 365
    d1 = (Log(s / K) + ((d - f + ((v ^ 2) / 2)) * tau)) / (v * Sqr(tau))
    d2 = d1 - (v * Sqr(tau))
    Phi = -1 * (Exp(-f * tau) * CND(d1) * s * tau)
End Function

Public Function Rho(s As Double, K _
                As Double, T As Date, E As Date, d As Double, v As Double, f As Double) As Double
                
    Dim d1 As Double, d2 As Double, tau As Double
    
    tau = (E - T) / 365
    d1 = (Log(s / K) + ((d - f + ((v ^ 2) / 2)) * tau)) / (v * Sqr(tau))
    d2 = d1 - (v * Sqr(tau))
    Rho = (Exp(-d * tau) * CND(d2) * K * tau)
End Function
Sub CalculateValues()

    Dim spot As Double, dom As Double, foreign As Double, vol As Double, strike As Double, today As Date, expiry As Date
    spot = Range("b1").Value
    dom = Range("b2").Value
    foreign = Range("b3").Value
    vol = Range("b4").Value
    strike = Range("b5").Value
    today = Range("b6").Value
    expiry = Range("b7").Value
    
    If Range("C25").Value = 1 Then
    
        Range("m1").Value = BlackScholes(spot, strike, today, expiry, dom, vol, foreign)
        Range("m2").Value = Delta(spot, strike, today, expiry, dom, vol, foreign)
        Range("m3").Value = Gamma(spot, strike, today, expiry, dom, vol, foreign)
        Range("m4").Value = Vega(spot, strike, today, expiry, dom, vol, foreign)
        Range("m5").Value = Theta(spot, strike, today, expiry, dom, vol, foreign)
        Range("m6").Value = Phi(spot, strike, today, expiry, dom, vol, foreign)
        Range("m7").Value = Rho(spot, strike, today, expiry, dom, vol, foreign)
    
        Range("n1").Value = "USD"
        Range("n2").Value = "USD change in Value for 1USD change in Spot"
        Range("n3").Value = "Change in Delta for 1USD change in Spot"
        Range("n4").Value = "USD change in Value for 100% change in Volatility"
        Range("n5").Value = "USD change in Value for 1 year change in time"
        Range("n6").Value = "USD change in Value for 100% change in foreign Interest rate"
        Range("n7").Value = "USD change in Value for 100% change in domestic Interest rate"
        
    ElseIf Range("C25").Value = 2 Then
    
        Range("m1").Value = BlackScholes(spot, strike, today, expiry, dom, vol, foreign)
        Range("m2").Value = Delta(spot, strike, today, expiry, dom, vol, foreign)
        Range("m3").Value = Gamma(spot, strike, today, expiry, dom, vol, foreign)
        Range("m4").Value = Vega(spot, strike, today, expiry, dom, vol, foreign) / 100
        Range("m5").Value = Theta(spot, strike, today, expiry, dom, vol, foreign) / 365
        Range("m6").Value = Phi(spot, strike, today, expiry, dom, vol, foreign) / 100
        Range("m7").Value = Rho(spot, strike, today, expiry, dom, vol, foreign) / 100
    
        Range("n1").Value = "USD"
        Range("n2").Value = "USD change in Value for 1USD change in Spot"
        Range("n3").Value = "Change in Delta for 1USD change in Spot"
        Range("n4").Value = "USD change in Value for 1% change in Volatility"
        Range("n5").Value = "USD change in Value for 1 day (1/365 year) change in time"
        Range("n6").Value = "USD change in Value for 1% change in foreign Interest rate"
        Range("n7").Value = "USD change in Value for 1% change in domestic Interest rate"
        
    End If
    
    
End Sub

Sub CreateGraphs()

    Application.Run ("Clear")
    
    Dim dom As Double, foreign As Double, vol As Double, strike As Double, today As Date, expiry As Date
    dom = Range("b2").Value
    foreign = Range("b3").Value
    vol = Range("b4").Value
    strike = Range("b5").Value
    today = Range("b6").Value
    expiry = Range("b7").Value
    
    Dim start As Double, endspot As Double, between As Double, increment As Double
    start = Range("b9").Value
    endspot = Range("b10").Value
    between = Range("b11").Value
    
    increment = (endspot - start) / (between + 1)
    
    Dim Greek1in As String, Greek2in As String, Greek1 As String, Greek2 As String, i As Integer
    Dim c As Collection
    Set c = New Collection
    c.Add "Delta", "1"
    c.Add "Gamma", "2"
    c.Add "Vega", "3"
    c.Add "Theta", "4"
    c.Add "Phi", "5"
    c.Add "Rho", "6"
     
    Greek1in = Range("D3").Value
    Greek2in = Range("E3").Value
    Greek1 = c.Item(Greek1in)
    Greek2 = c.Item(Greek2in)
    
    Range("H1").Value = Greek1
    Range("I1").Value = Greek2
    
    For i = 2 To (between + 3):
        Cells(i, 7).Value = start + (increment * (i - 2))
        Cells(i, 8).Value = Application.Run(Greek1, Cells(i, 7).Value, strike, today, expiry, dom, vol, foreign)
        Cells(i, 9).Value = Application.Run(Greek2, Cells(i, 7).Value, strike, today, expiry, dom, vol, foreign)
    Next i
    
    Application.Run ("Graph1")
    Application.Run ("Graph2")
    Application.Run ("Graph3")
End Sub

Sub Clear()
    Range("G:G").ClearContents
    Range("H:H").ClearContents
    Range("I:I").ClearContents
    Range("G1").Value = "Spot"
    Range("H1").Value = "Greek 1"
    Range("I1").Value = "Greek 2"
    
    Dim chtObj As ChartObject
    For Each chtObj In ActiveSheet.ChartObjects
        chtObj.Delete
    Next
End Sub


Sub Graph1()

    Dim Greek1name As String, start As Double, endspot As Double
    Greek1name = Range("H1").Value
    start = Range("b9").Value
    endspot = Range("b10").Value
    Dim xaxis As Range
    Dim yaxis As Range
    Set xaxis = Range("$G$2", Range("$G$2").End(xlDown))
    Set yaxis = Range("$H$2", Range("$H$2").End(xlDown))
    
    Dim c1 As Chart
    Dim co As ChartObject
    Set co = ActiveWorkbook.Sheets("Sheet1").ChartObjects.Add(510, 110, 300, 200)
    Set c1 = co.Chart
    With c1
        .ChartType = xlXYScatterLines
        .HasLegend = False
        .HasTitle = True
        .ChartTitle.Text = "Spot - " & Greek1name
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Spot Prices"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = Greek1name
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

Sub Graph2()

    Dim Greek2name As String, start As Double, endspot As Double
    Greek2name = Range("I1").Value
    start = Range("b9").Value
    endspot = Range("b10").Value
    Dim xaxis As Range
    Dim yaxis As Range
    Set xaxis = Range("$G$2", Range("$G$2").End(xlDown))
    Set yaxis = Range("$I$2", Range("$I$2").End(xlDown))
    
    Dim c1 As Chart
    Dim co As ChartObject
    Set co = ActiveWorkbook.Sheets("Sheet1").ChartObjects.Add(510, 320, 300, 200)
    Set c1 = co.Chart
    With c1
        .ChartType = xlXYScatterLines
        .HasLegend = False
        .HasTitle = True
        .ChartTitle.Text = "Spot - " & Greek2name
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Spot Prices"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = Greek2name
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

Sub Graph3()

    Dim Greek1name As String, Greek2name As String, start As Double, endspot As Double
    Greek1name = Range("H1").Value
    Greek2name = Range("I1").Value
    'start = Range("b9").Value
    'endspot = Range("b10").Value
    Dim xaxis As Range
    Dim yaxis As Range
    Set xaxis = Range("$H$2", Range("$H$2").End(xlDown))
    Set yaxis = Range("$I$2", Range("$I$2").End(xlDown))
    
    Dim c1 As Chart
    Dim co As ChartObject
    Set co = ActiveWorkbook.Sheets("Sheet1").ChartObjects.Add(850, 110, 300, 200)
    Set c1 = co.Chart
    With c1
        .ChartType = xlXYScatterLines
        .HasLegend = False
        .HasTitle = True
        .ChartTitle.Text = Greek1name & " - " & Greek2name
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = Greek1name
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = Greek2name
        '.Axes(xlCategory).MinimumScale = start
        '.Axes(xlCategory).MaximumScale = endspot
    End With
    
    Dim s As Series
    Set s = c1.SeriesCollection.NewSeries
    With s
        .Values = yaxis
        .XValues = xaxis
    End With
    
End Sub
