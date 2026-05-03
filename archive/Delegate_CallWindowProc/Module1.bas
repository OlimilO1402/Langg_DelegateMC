Attribute VB_Name = "Module1"
Option Explicit
#If VBA7 = 0 Then
    Public Enum LongPtr: [_]: End Enum
#End If
#If VBA7 Then
    Public Declare PtrSafe Function MathFunc4Lng Lib "user32" Alias "CallWindowProcA" (ByVal lpFunc As LongPtr, ByVal Value1 As Long, ByVal Value2 As Long, ByVal Value3 As Long, ByVal Value4 As Long) As Long
    Public Declare PtrSafe Function MathFunc4SngRetSng Lib "user32" Alias "CallWindowProcA" (ByVal lpFunc As LongPtr, ByVal Value1 As Single, ByVal Value2 As Single, ByVal Value3 As Single, ByVal Value4 As Single) As Single
    Public Declare PtrSafe Function MathFunc4RefDblRetSng Lib "user32" Alias "CallWindowProcA" (ByVal lpFunc As LongPtr, ByRef Value1 As Double, ByRef Value2 As Double, ByRef Value3 As Double, ByRef Value4 As Double) As Single
    Public Declare PtrSafe Function MathFunc4RefDblRetDbl Lib "user32" Alias "CallWindowProcA" (ByVal lpFunc As LongPtr, ByRef Value1 As Double, ByRef Value2 As Double, ByRef Value3 As Double, ByRef Value4 As Double) As Double
    Public Declare PtrSafe Function MathFunc2ValDblRetDbl Lib "user32" Alias "CallWindowProcA" (ByVal lpFunc As LongPtr, ByVal Value1 As Double, ByVal Value2 As Double) As Double
    'the following crashes in x86 but works in x64:
    #If Win64 Then
        Public Declare PtrSafe Function MathFunc4ValDblRetDbl Lib "user32" Alias "CallWindowProcA" (ByVal lpFunc As LongPtr, ByVal Value1 As Double, ByVal Value2 As Double, ByVal Value3 As Double, ByVal Value4 As Double) As Double
    #End If
#Else
    Public Declare Function MathFunc4Lng Lib "user32" Alias "CallWindowProcA" (ByVal lpFunc As LongPtr, ByVal Value1 As Long, ByVal Value2 As Long, ByVal Value3 As Long, ByVal Value4 As Long) As Long
    Public Declare Function MathFunc4SngRetSng Lib "user32" Alias "CallWindowProcA" (ByVal lpFunc As LongPtr, ByVal Value1 As Single, ByVal Value2 As Single, ByVal Value3 As Single, ByVal Value4 As Single) As Single
    Public Declare Function MathFunc4RefDblRetSng Lib "user32" Alias "CallWindowProcA" (ByVal lpFunc As LongPtr, ByRef Value1 As Double, ByRef Value2 As Double, ByRef Value3 As Double, ByRef Value4 As Double) As Single
    Public Declare Function MathFunc4RefDblRetDbl Lib "user32" Alias "CallWindowProcA" (ByVal lpFunc As LongPtr, ByRef Value1 As Double, ByRef Value2 As Double, ByRef Value3 As Double, ByRef Value4 As Double) As Double
    Public Declare Function MathFunc2ValDblRetDbl Lib "user32" Alias "CallWindowProcA" (ByVal lpFunc As LongPtr, ByVal Value1 As Double, ByVal Value2 As Double) As Double
    'the following crashes in x86 but works in x64:
    #If Win64 Then
        Public Declare Function MathFunc4ValDblRetDbl Lib "user32" Alias "CallWindowProcA" (ByVal lpFunc As LongPtr, ByVal Value1 As Double, ByVal Value2 As Double, ByVal Value3 As Double, ByVal Value4 As Double) As Double
    #End If
#End If

Public Function FncPtr(ByVal pFunc As Long) As Long
    FncPtr = pFunc
End Function

Public Function Add4Lng(ByVal Value1 As Long, ByVal Value2 As Long, ByVal Value3 As Long, ByVal Value4 As Long) As Long
    Add4Lng = Value1 + Value2 + Value3 + Value4
End Function

Public Function Add4Sng(ByVal Value1 As Single, ByVal Value2 As Single, ByVal Value3 As Single, ByVal Value4 As Single) As Single
    Add4Sng = Value1 + Value2 + Value3 + Value4
End Function

Public Function Add4RefDblRetSng(ByRef Value1 As Double, ByRef Value2 As Double, ByRef Value3 As Double, ByRef Value4 As Double) As Single
    Add4RefDblRetSng = Value1 + Value2 + Value3 + Value4
End Function

Public Function Add4RefDblRetDbl(ByRef Value1 As Double, ByRef Value2 As Double, ByRef Value3 As Double, ByRef Value4 As Double) As Double
    Add4RefDblRetDbl = Value1 + Value2 + Value3 + Value4
End Function

Public Function Add2ValDblRetDbl(ByVal Value1 As Double, ByVal Value2 As Double) As Double '?
    Add2ValDblRetDbl = Value1 + Value2
End Function

'the following crashes in x86 but works in x64:
#If Win64 Then
    Public Function Add4ValDblRetDbl(ByVal Value1 As Double, ByVal Value2 As Double, ByVal Value3 As Double, ByVal Value4 As Double) As Double
        Add4ValDblRetDbl = Value1 + Value2 + Value3 + Value4
    End Function
#End If
