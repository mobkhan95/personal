Attribute VB_Name = "Module1"

Global Const Pi = 3.14159265358979
Public start_cell As String
Public end_cell As String
Public q As Integer
'Adapted from stackexchange
Sub PrintArrayAsRange(myArr As Variant, cell As Range)
    Dim arr2D As Variant
    arr2D = Application.WorksheetFunction.Transpose(myArr)
    cell.Resize(UBound(arr2D, 1), UBound(arr2D, 2)) = arr2D
End Sub

Public Function Cubic(x As Double) As Double

Dim deltas() As Variant
deltas = Range(start_cell & ":" & Cells(Range(end_cell).Row, Range(end_cell).Column + 1).Address).Value
ReDim Lambda(1 To (Range(start_cell & ":" & end_cell).Rows.Count - 1)) As Double
ReDim d(1 To Range(start_cell & ":" & end_cell).Rows.Count) As Double
ReDim Mu(2 To Range(start_cell & ":" & end_cell).Rows.Count) As Double
ReDim M(1 To Range(start_cell & ":" & end_cell).Rows.Count) As Double
ReDim w(2 To Range(start_cell & ":" & end_cell).Rows.Count) As Double
ReDim h(2 To Range(start_cell & ":" & end_cell).Rows.Count) As Double
ReDim b(1 To Range(start_cell & ":" & end_cell).Rows.Count) As Double

b(LBound(b)) = 2
For i = 2 To UBound(h):
    h(i) = deltas(i, 1) - deltas(i - 1, 1)
    b(i) = 2
    Next i

If (Range("D13") = 1) Or (Range("D13") = 3) Then
    Lambda(LBound(Lambda)) = 0
    d(LBound(d)) = 2 * Range("D11").Value
End If

If (Range("D13") = 2) Then
    Lambda(LBound(Lambda)) = 1
    d(LBound(d)) = (6 / h(LBound(h))) * (((deltas(LBound(d) + 1, 2) - deltas(LBound(d), 2)) / h(LBound(h))) - Range("D11").Value)
End If

If (Range("E13") = 1) Or (Range("E13") = 3) Then
    Mu(UBound(Mu)) = 0
    d(UBound(d)) = 2 * Range("E11").Value
End If

If (Range("D13") = 2) Then
    Mu(UBound(Mu)) = 1
    d(UBound(d)) = (6 / h(UBound(h))) * (Range("E11").Value - ((deltas(UBound(d), 2) - deltas(UBound(d) - 1, 2)) / h(UBound(h))))
End If


For i = LBound(Mu) To (UBound(Mu) - 1)
    Lambda(i) = (h(i + 1) / (h(i) + h(i + 1)))
    Mu(i) = 1 - Lambda(i)
    d(i) = (6 / (h(i) + h(i + 1))) * (((deltas(i + 1, 2) - deltas(i, 2)) / h(i + 1)) - ((deltas(i, 2) - deltas(i - 1, 2)) / h(i)))
    Next i

ActiveWorkbook.Worksheets("Cubic Spline Matrix").Range("A1").Value = "Vector Lambda"
PrintArrayAsRange Lambda, ActiveWorkbook.Worksheets("Cubic Spline Matrix").Range("A2:A15")
ActiveWorkbook.Worksheets("Cubic Spline Matrix").Range("B1").Value = "Vector Mu"
PrintArrayAsRange Mu, ActiveWorkbook.Worksheets("Cubic Spline Matrix").Range("B2:B15")
ActiveWorkbook.Worksheets("Cubic Spline Matrix").Range("C1").Value = "Vector d"
PrintArrayAsRange d, ActiveWorkbook.Worksheets("Cubic Spline Matrix").Range("C2:C15")


For i = LBound(w) To UBound(w)
    w(i) = Mu(i) / b(i - 1)
    b(i) = b(i) - (Lambda(i - 1) * w(i))
    d(i) = d(i) - (d(i - 1) * w(i))
    Next i
M(UBound(M)) = d(UBound(d)) / b(UBound(b))

For i = (UBound(M) - 1) To LBound(M) Step -1
    M(i) = (d(i) - Lambda(i) * M(i + 1)) / b(i)
    Next i
    
ActiveWorkbook.Worksheets("Cubic Spline Matrix").Range("E1").Value = "Moments Vector"
PrintArrayAsRange M, ActiveWorkbook.Worksheets("Cubic Spline Matrix").Range("E2:E15")
    
ReDim Alpha(1 To (Range(start_cell & ":" & end_cell).Rows.Count - 1)) As Double
ReDim Beta(1 To (Range(start_cell & ":" & end_cell).Rows.Count - 1)) As Double
ReDim Gamma(1 To (Range(start_cell & ":" & end_cell).Rows.Count - 1)) As Double
ReDim Delta(1 To (Range(start_cell & ":" & end_cell).Rows.Count - 1)) As Double

For i = LBound(Alpha) To UBound(Alpha)
    Alpha(i) = deltas(i, 2)
    Gamma(i) = M(i) / 2
    Delta(i) = (M(i + 1) - M(i)) / (6 * h(i + 1))
    Beta(i) = ((deltas(i + 1, 2) - deltas(i, 2)) / h(i + 1)) - ((2 * M(i) + M(i + 1)) * (h(i + 1) / 6))
    Next i
    
ActiveWorkbook.Worksheets("Cubic Spline Matrix").Range("G1").Value = "Vector Alpha"
PrintArrayAsRange Alpha, ActiveWorkbook.Worksheets("Cubic Spline Matrix").Range("G2:G15")
ActiveWorkbook.Worksheets("Cubic Spline Matrix").Range("H1").Value = "Vector Beta"
PrintArrayAsRange Beta, ActiveWorkbook.Worksheets("Cubic Spline Matrix").Range("H2:H15")
ActiveWorkbook.Worksheets("Cubic Spline Matrix").Range("I1").Value = "Vector Gamma"
PrintArrayAsRange Gamma, ActiveWorkbook.Worksheets("Cubic Spline Matrix").Range("I2:I15")
ActiveWorkbook.Worksheets("Cubic Spline Matrix").Range("J1").Value = "Vector Delta"
PrintArrayAsRange Delta, ActiveWorkbook.Worksheets("Cubic Spline Matrix").Range("J2:J15")


Dim below As Double, b_index As Integer
b_index = Closest_Below(deltas, x)
If b_index = UBound(deltas, 1) Then
    Cubic = deltas(b_index, 2)
    Exit Function
End If
below = deltas(b_index, 1)

Cubic = Alpha(b_index) + (Beta(b_index) * (x - below)) + (Gamma(b_index) * (x - below) * (x - below)) + (Delta(b_index) * (x - below) * (x - below) * (x - below))







End Function

Sub Setup()


Application.Run ("Clear")

start_cell = Cells(2, 1).Address
end_cell = Cells(10, 1).Address

Range(start_cell & ":" & Cells(Range(end_cell).Row, Range(end_cell).Column + 1).Address).Sort Key1:=Range(start_cell), Order1:=xlAscending, Header:=xlNo
Range("a13").Value = Range(start_cell).Value
Range("b13").Value = Range(end_cell).Value

If Range("d13") = 1 Then
    Range("D11") = 0
End If

If Range("E13") = 1 Then
    Range("E11") = 0
End If

If IsEmpty(Range("d11")) Or IsEmpty(Range("e11")) Then
    MsgBox ("Enter Boundary Conditions in Cells D11 and E11")
End If


End Sub

Public Function Linear(x As Double) As Double

Dim deltas() As Variant, below As Double, above As Double, b_index As Integer, a_index As Integer
deltas = Range(start_cell & ":" & Cells(Range(end_cell).Row, Range(end_cell).Column + 1).Address).Value
b_index = Closest_Below(deltas, x)
If b_index = UBound(deltas, 1) Then
    Linear = deltas(b_index, 2)
    Exit Function
End If
a_index = b_index + 1
below = deltas(b_index, 1)
above = deltas(a_index, 1)

Linear = ((x - deltas(b_index, 1)) * ((deltas(a_index, 2) - deltas(b_index, 2)) / (deltas(a_index, 1) - deltas(b_index, 1)))) + deltas(b_index, 2)

End Function

Public Function Closest_Below(deltas As Variant, x As Double) As Double

Dim current As Double, index As Integer
current = Abs(x - deltas(LBound(deltas, 1), 1))
index = LBound(deltas, 1)
For i = (LBound(deltas, 1) + 1) To UBound(deltas, 1)
    If deltas(i, 1) <= x Then
        If Abs((x - deltas(i, 1))) < current Then
            current = Abs((x - deltas(i, 1)))
            index = i
        End If
    Else: Exit For
    
    End If
        
    Next i
    
    Closest_Below = index

End Function

Public Function Neville(x As Double, Optional flag As Integer = 0) As Double

Dim deltas() As Variant
Dim Arr() As Integer
deltas = Range(start_cell & ":" & Cells(Range(end_cell).Row, Range(end_cell).Column + 1).Address).Value
ReDim Arr(1 To UBound(deltas, 1)) As Integer
For i = 1 To UBound(deltas, 1)
    Arr(i) = i
    Next i
q = 1
If flag = 1 Then
    Neville = Recursion_Int(deltas, x, Arr)
    Exit Function
End If
Neville = Recursion(deltas, x, Arr)
End Function

Public Function Recursion_Int(deltas As Variant, x As Double, ByRef Arr As Variant) As Double

Dim val As String
If (UBound(Arr) - LBound(Arr) + 1) = 1 Then
    Recursion_Int = deltas(Arr(1), 2)
    Worksheets("Tree Diagram for Neville").Cells(q, 1).Value = Recursion_Int
    Worksheets("Tree Diagram for Neville").Cells(q, 3).Value = Arr
    q = q + 1
    Exit Function
End If

Dim F As Double, L As Double
F = deltas(Arr(LBound(Arr)), 1)
L = deltas(Arr(UBound(Arr)), 1)

ReDim Arr_F(LBound(Arr) To (UBound(Arr) - 1)) As Integer
ReDim Arr_L(LBound(Arr) To (UBound(Arr) - 1)) As Integer

For i = LBound(Arr) To (UBound(Arr) - 1)
    Arr_F(i) = Arr(i)
    Arr_L(i) = Arr(i + 1)
    Next i

Recursion_Int = (((x - F) * Recursion_Int(deltas, x, Arr_L)) - ((x - L) * Recursion_Int(deltas, x, Arr_F))) / (L - F)
Worksheets("Tree Diagram for Neville").Cells(q, 1).Value = Recursion_Int
For i = LBound(Arr) To UBound(Arr)
    Worksheets("Tree Diagram for Neville").Cells(q, 2 + i) = Arr(i)
    Next i
q = q + 1


End Function
Public Function Recursion(deltas As Variant, x As Double, ByRef Arr As Variant) As Double

Dim val As String
If (UBound(Arr) - LBound(Arr) + 1) = 1 Then
    Recursion = deltas(Arr(1), 2)
    Exit Function
End If

Dim F As Double, L As Double
F = deltas(Arr(LBound(Arr)), 1)
L = deltas(Arr(UBound(Arr)), 1)

ReDim Arr_F(LBound(Arr) To (UBound(Arr) - 1)) As Integer
ReDim Arr_L(LBound(Arr) To (UBound(Arr) - 1)) As Integer

For i = LBound(Arr) To (UBound(Arr) - 1)
    Arr_F(i) = Arr(i)
    Arr_L(i) = Arr(i + 1)
    Next i

Recursion = (((x - F) * Recursion(deltas, x, Arr_L)) - ((x - L) * Recursion(deltas, x, Arr_F))) / (L - F)


End Function


Sub Perform_Interpolations()

Application.Run ("Setup")

Range("B18").Value = Linear(Range("B15").Value)
Range("B19").Value = Neville(Range("B15").Value, 1)
Range("B20").Value = Cubic(Range("B15").Value)

End Sub

Sub Create_Graphs()
Application.Run ("Setup")


Dim c As Collection, int_type As String
Set c = New Collection
c.Add "Linear", "1"
c.Add "Neville", "2"
c.Add "Cubic", "3"

int_type = c.Item(Range("D16").Value)

Dim startval As Double, endval As Double, between As Double, increment As Double
    startval = Range(start_cell).Value
    endval = Range(end_cell).Value
    between = Range("b22").Value
    increment = (endval - startval) / (between + 1)
    
    For i = 2 To (between + 3):
        Cells(i, 7).Value = start + (increment * (i - 2))
        Cells(i, 8).Value = Application.Run(int_type, Cells(i, 7).Value)
    Next i

Range("$G$2", Range("$H$2").End(xlDown)).Sort Key1:=Range("$G$2"), Order1:=xlDescending, Header:=xlNo

Dim xaxis As Range
Dim yaxis As Range
Set xaxis = Range("$G$2", Range("$G$2").End(xlDown))
Set yaxis = Range("$H$2", Range("$H$2").End(xlDown))

Dim c1 As Chart
Dim co As ChartObject
Set co = ActiveWorkbook.Sheets("Sheet1").ChartObjects.Add(550, 150, 300, 200)
Set c1 = co.Chart
With c1
    .ChartType = xlXYScatterLines
    .HasLegend = False
    .HasTitle = True
    .ChartTitle.Text = "Implied Volatility vs. Delta using " & int_type & " Interpolation"
    .Axes(xlCategory, xlPrimary).HasTitle = True
    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Delta Values"
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Volatility"
    .Axes(xlCategory).ReversePlotOrder = True
End With

Dim s As Series
Set s = c1.SeriesCollection.NewSeries
With s
    .Values = yaxis
    .XValues = xaxis
End With



    
End Sub
Sub Clear()
    Range("G:G").ClearContents
    Range("H:H").ClearContents
    Range("G1").Value = "Delta"
    Range("H1").Value = "Vol"

    Dim chtObj As ChartObject
    For Each chtObj In ActiveSheet.ChartObjects
        chtObj.Delete
    Next
    
    Sheets("Tree Diagram for Neville").Cells.ClearContents
    
End Sub
