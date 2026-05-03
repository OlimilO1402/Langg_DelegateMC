VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    Dim pFunc As LongPtr
    Dim LngVal As Long, SngVal As Single, DblVal As Double
    
    pFunc = FncPtr(AddressOf Module1.Add4Lng)
    LngVal = Module1.MathFunc4Lng(pFunc, 1, 2, 3, 4)
    Debug.Print LngVal '10
    
    pFunc = FncPtr(AddressOf Module1.Add4Sng)
    SngVal = Module1.MathFunc4SngRetSng(pFunc, 1.111111!, 2.222222!, 3.333333!, 4.444444!)
    Debug.Print SngVal '11,11111
    
    pFunc = FncPtr(AddressOf Module1.Add4RefDblRetSng)
    SngVal = Module1.MathFunc4RefDblRetSng(pFunc, 1.11111111111111, 2.22222222222222, 3.33333333333333, 4.44444444444444)
    Debug.Print SngVal ' 11,11111
    
    pFunc = FncPtr(AddressOf Module1.Add4RefDblRetDbl)
    DblVal = Module1.MathFunc4RefDblRetDbl(pFunc, 1.11111111111111, 2.22222222222222, 3.33333333333333, 4.44444444444444)
    Debug.Print DblVal ' 11,1111111111111
    
    pFunc = FncPtr(AddressOf Module1.Add2ValDblRetDbl)
    DblVal = Module1.MathFunc2ValDblRetDbl(pFunc, 1.11111111111111, 2.22222222222222)
    Debug.Print DblVal ' 3,33333333333333
    
    'crashes in x86 but works in x64:
#If Win64 Then
    pFunc = FncPtr(AddressOf Module1.Add4ValDblRetDbl)
    DblVal = Module1.MathFunc4ValDblRetDbl(pFunc, 1.1111111111, 2.2222222222, 3.3333333333, 4.4444444444)
    Debug.Print DblVal ' 11,1111111111
#End If
End Sub

