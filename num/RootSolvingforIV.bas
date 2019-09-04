Attribute VB_Name = "Module2"
Global Const Pi = 3.14159265358979
Public spot As Double, dom As Double, foreign As Double, strike As Double, today As Date, expiry As Date


'// The normal distribution function
Public Function ND(x As Double) As Double
    ND = 1 / Sqr(2 * Pi) * Exp(-x ^ 2 / 2)
End Function

'// The cumulative normal distribution function
Public Function CND(x As Double) As Double
    
    Dim L As Double, K As Double
    Const a1 = 0.31938153:  Const a2 = -0.356563782: Const a3 = 1.781477937:
    Const a4 = -1.821255978:  Const a5 = 1.330274429
    
    L = Abs(x)
    K = 1 / (1 + 0.2316419 * L)
    CND = 1 - 1 / Sqr(2 * Pi) * Exp(-L ^ 2 / 2) * (a1 * K + a2 * K ^ 2 + a3 * K ^ 3 + a4 * K ^ 4 + a5 * K ^ 5)
    
    If x < 0 Then
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
Public Function Vega(s As Double, K _
                As Double, T As Date, E As Date, d As Double, v As Double, f As Double) As Double
                
    Dim d1 As Double, d2 As Double, tau As Double
    
    tau = (E - T) / 365
    d1 = (Log(s / K) + ((d - f + ((v ^ 2) / 2)) * tau)) / (v * Sqr(tau))
    d2 = d1 - (v * Sqr(tau))
    Vega = (Exp(-f * tau) * ND(d1) * s * Sqr(tau))
End Function
Public Function it(x As Double, T As Double) As Double
it = x - ((BlackScholes(spot, strike, today, expiry, dom, x, foreign) - T) / Vega(spot, strike, today, expiry, dom, x, foreign))
End Function
Public Function NR2(g As Double, T As Double) As Variant
    Do Until (Abs(it(g, T) - g) < 0.00000001)
        g = it(g, T)
        If g >= 100000 Then
            MsgBox ("Newton's Method for the Call IV Does not Converge")
            NR2 = "Didn't Converge"
            Exit Function
        End If
    Loop
    NR2 = g
End Function
Public Function mu(a As Double, x As Double, T As Double) As Double
mu = (((a * (BlackScholes(spot, strike, today, expiry, dom, x, foreign) - T)) - (x * (BlackScholes(spot, strike, today, expiry, dom, a, foreign) - T))) / ((BlackScholes(spot, strike, today, expiry, dom, x, foreign) - T) - (BlackScholes(spot, strike, today, expiry, dom, a, foreign) - T)))
End Function
Public Function RF(a As Double, x As Double, T As Double) As Variant
Dim hold As Double

If ((BlackScholes(spot, strike, today, expiry, dom, a, foreign) - T) * (BlackScholes(spot, strike, today, expiry, dom, x, foreign) - T)) >= 0 Then
    MsgBox ("For Regula Falsi, the Initial Guesses must lead to Opposite Signs to Guarantee Convergence")
    If Range("i19").Value = 1 Then
        RF = "Terminated Immediately since Initial Guesses did not have opposite signs"
        Exit Function
    End If
End If

Do Until (Abs(BlackScholes(spot, strike, today, expiry, dom, mu(a, x, T), foreign) - T) < 0.0000001)
    If ((BlackScholes(spot, strike, today, expiry, dom, mu(a, x, T), foreign) - T) * (BlackScholes(spot, strike, today, expiry, dom, x, foreign) - T)) > 0 Then
        x = mu(a, x, T)
        a = a
    Else
        hold = a
        a = x
        x = mu(hold, x, T)
    End If
Loop
    RF = mu(a, x, T)
End Function
Public Function Secant(a As Double, x As Double, T As Double) As Variant
Dim hold As Double
If a = x Then
    MsgBox ("For Secant, the Initial Guesses cannot be the same")
    Secant = "Did not work"
    Exit Function
End If
Do Until (Abs(BlackScholes(spot, strike, today, expiry, dom, mu(a, x, T), foreign) - T) < 0.0000001)
        hold = a
        a = x
        x = mu(hold, x, T)
        If x >= 100000 Then
            MsgBox ("Secant Method for the Call IV Does not Converge")
            Secant = "Didn't Converge"
            Exit Function
        End If
Loop
    Secant = mu(a, x, T)
End Function


Sub runner()

Range("G:G").ClearContents

spot = Range("j1").Value
dom = Range("j2").Value
foreign = Range("j3").Value
strike = Range("j5").Value
today = Range("j6").Value
expiry = Range("j7").Value

Dim target As Double
Dim guessNR As Double, guessRFa As Double, guessRFb As Double, guessSa As Double, guessSb As Double

target = Range("b8").Value
guessNR = Range("j13").Value
guessRFa = Range("j15").Value
guessRFb = Range("k15").Value
guessSa = Range("j14").Value
guessSb = Range("k14").Value

If target < 0 Then
    MsgBox ("Call Price cannot be negative")
    Exit Sub
End If


Range("B12").Value = RF(guessRFa, guessRFb, target)
Range("B11").Value = NR2(guessNR, target)
Range("B13").Value = Secant(guessSa, guessSb, target)

End Sub

