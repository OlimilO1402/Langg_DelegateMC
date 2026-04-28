Attribute VB_Name = "MMathFunctions"
Option Explicit
'Public C As List(Of Double)
Public c As Variant ' As New Collection ' Konstantenparameter

Public Function New_MathFunction(aDF As TDelegateFunction, ByVal pFncAddr As Long) As MathFunction 'ICallDoubleReturnDouble 'ICallLongReturnLong
    Set New_MathFunction = New_DelegateFunction(aDF, pFncAddr)
End Function

' hier nur zum Probieren vier Beispielfunktionen
Public Function Linear(ByVal x As Double) As Double
'Public Function Linear(ByVal x As Long) As Long

    Linear = c(0) * x + c(1)

End Function

Public Function Quadratic(ByVal x As Double) As Double
'Public Function Quadratic(ByVal x As Long) As Long

    Quadratic = Linear(x) * x + c(2)

End Function

Public Function Cubic(ByVal x As Double) As Double
'Public Function Cubic(ByVal x As Long) As Long

    Cubic = Quadratic(x) * x + c(3)

End Function

' ... noch mehr Polynome?

Public Function Sinus(ByVal x As Double) As Double
'Public Function Sinus(ByVal x As Long) As Long

    Sinus = Math.Sin(Linear(x)) * c(2) + c(3)

End Function

Public Function Exponent(ByVal x As Double) As Double
'Public Function Exponent(ByVal x As Long) As Long

    Exponent = Math.Exp(Linear(x)) * c(2) + c(3)

End Function

Public Function DamperedHarmonic(ByVal x As Double) As Double
'Public Function DamperedHarmonic(ByVal x As Long) As Long

    DamperedHarmonic = Exponent(x) * Math.Sin(c(4) * x + c(5)) + c(6)

End Function

' Man müßte hier noch viel freier sein dürfen,
' und die Funktionen verketten können
' y = f(f(f(x)))
' ...
