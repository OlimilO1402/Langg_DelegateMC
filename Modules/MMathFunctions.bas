Attribute VB_Name = "MMathFunctions"
Option Explicit
'Public C As List(Of Double)
Public c As Variant ' As New Collection ' Konstantenparameter

' hier nur zum Probieren ein paar Beispielfunktionen
Public Function Linear(ByVal x As Double) As Double

    Linear = c(0) * x + c(1)

End Function

Public Function Quadratic(ByVal x As Double) As Double

    Quadratic = Linear(x) * x + c(2)

End Function

Public Function Cubic(ByVal x As Double) As Double

    Cubic = Quadratic(x) * x + c(3)

End Function

' ... noch mehr Polynome?
Public Function Sinus(ByVal x As Double) As Double

    Sinus = Math.Sin(Linear(x)) * c(2) + c(3)

End Function

Public Function Exponent(ByVal x As Double) As Double

    Exponent = Math.Exp(Linear(x)) * c(2) + c(3)

End Function

Public Function DamperedHarmonic(ByVal x As Double) As Double

    DamperedHarmonic = Exponent(x) * Math.Sin(c(4) * x + c(5)) + c(6)

End Function

' Man müßte hier noch viel freier sein dürfen,
' und die Funktionen verketten können
' y = f(f(f(x)))
' ...
