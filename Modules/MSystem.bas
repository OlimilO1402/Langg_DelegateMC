Attribute VB_Name = "MSystem"
Option Explicit
Public Enum EMathFuncs
    None = 0
    Linear = 1
    Quadratic
    Cubic
    Sinus
    Exponent
    DamperedHarmonic
End Enum

Public Type Point
    x As Double 'Long
    y As Double 'Long
End Type

Public Function EMathFuncs_ToStr(e As EMathFuncs) As String
    Dim s As String
    Select Case e
    Case EMathFuncs.None:             s = "None"
    Case EMathFuncs.Linear:           s = "Linear"
    Case EMathFuncs.Quadratic:        s = "Quadratic"
    Case EMathFuncs.Cubic:            s = "Cubic"
    Case EMathFuncs.Sinus:            s = "Sinus"
    Case EMathFuncs.Exponent:         s = "Exponent"
    Case EMathFuncs.DamperedHarmonic: s = "DamperedHarmonic"
    End Select
    EMathFuncs_ToStr = s
End Function

Public Sub EMathFuncs_ToCBLB(CBLB)
    With CBLB
        .Clear
        Dim i As Long
        For i = 0 To 6
            .AddItem EMathFuncs_ToStr(i)
        Next
    End With
End Sub

Public Function EMathFuncs_ToMathFormula(e As EMathFuncs) As String
    Dim s As String
    Select Case e
    Case EMathFuncs.None:             s = "--"
    Case EMathFuncs.Linear:           s = "c1 * x + c0"
    Case EMathFuncs.Quadratic:        s = "c2 * x² + c1 * x + c0"
    Case EMathFuncs.Cubic:            s = "c3 * x³ + c2 * x² + c1 * x + c0"
    Case EMathFuncs.Sinus:            s = "Sin(c3 * x + c2) * c1 + c0"
    Case EMathFuncs.Exponent:         s = "Exp(c3 * x + c2) * c1 + c0"
    Case EMathFuncs.DamperedHarmonic: s = "(Exp(c6 * x + c5) * c4 + c3) * Math.Sin(c2 * x + c1) + c0"
    End Select
    EMathFuncs_ToMathFormula = s
End Function


'Public Function New_Point(ByVal x As Long, ByVal y As Long) As Point
Public Function New_Point(ByVal x As Double, ByVal y As Double) As Point
    With New_Point: .x = x: .y = y: End With
End Function
