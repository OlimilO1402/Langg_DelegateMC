VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   625
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3075
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   480
      Width           =   4935
   End
   Begin VB.ComboBox CmbMathFunctions 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label LblMathFormula 
      AutoSize        =   -1  'True
      Caption         =   "--"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   150
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mDF As TDelegateFunction
Private f   As MathFunction 'MathFunction ist in einer Typelibrary definiert
Private mCLin As Variant ' Konstantenparameter für eine lineare Funktion
Private mCQud As Variant ' Konstantenparameter für eine quadratische Funktion
Private mCCub As Variant ' Konstantenparameter für eine kubische Funktion
Private mCSin As Variant ' Konstantenparameter für eine Sinus-Funktion
Private mCExp As Variant ' Konstantenparameter für eine Expon-Funktion
Private mCDpH As Variant ' Konstantenparameter für eine ged. Schwingung
Private mView As New MathFunctionView

Private Sub Command1_Click()
'    Dim aGuid As Guid: Set aGuid = MNew.GuidCo
'    Dim s As String: s = aGuid.ToStr
'    s = InputBox("Guid:", "GUID", UCase(s))
End Sub

Private Sub Form_Load()
    
    ' Alle Namen der Enumkonstanten in die ComboBox eintragen
    EMathFuncs_ToCBLB CmbMathFunctions
    
    Set mView = MNew.MathFunctionView(Picture1)
    ' irgendwelche Konstantenparameter voreinstellen
    ' wer will, kann sie hier mit der Hand anpassen
    ' (C(0), C(1), C(2), C(3))
    mCLin = Array(0.5, 1.5) '15#)
    mCQud = Array(1, 0#, -2#)
    mCCub = Array(0.1, 0#, -2#, 0#)
    mCSin = Array(0.5, 1#, 2#, 0#)
    mCExp = Array(0.7, 0#, 5#, 0#)
    mCDpH = Array(-0.7, 0, 5, 0.2, 0.1, 10#, 0)
    'mCDpH = Array(-0.007, 0, 25, 0, 0.1, 10, 0)

End Sub

Private Sub Form_Resize()
    
    Dim br As Single: br = 8 '* Screen.TwipsPerPixelX
    Dim l As Single: l = br
    Dim t As Single: t = Picture1.Top
    Dim W As Single: W = Me.ScaleWidth - l - br
    Dim H As Single: H = Me.ScaleHeight - t - br
    
    If W > 0 And H > 0 Then
        Picture1.Move l, t, W, H ' br, PictureBox1.Top, .ScaleWidth - PictureBox1.Left - br, .ScaleHeight - PictureBox1.Top - br
    End If
    
    'Set mView.Size = PictureBox1
    mView.DrawFunction f
    'Call PictureBox1.Refresh

End Sub

Sub test()
    Dim f As MathFunction
    Dim res As Double: res = f.von(123)
    
End Sub
Private Sub CmbMathFunctions_Click()
    Dim e As EMathFuncs: e = CmbMathFunctions.ListIndex
    LblMathFormula.Caption = MSystem.EMathFuncs_ToMathFormula(e)
    Select Case CmbMathFunctions.ListIndex
    Case EMathFuncs.None:               Set f = Nothing
    Case EMathFuncs.Linear:             MMathFunctions.c = mCLin
                                        Set f = MNew.MathFunction(mDF, AddressOf MMathFunctions.Linear)
    Case EMathFuncs.Quadratic:          MMathFunctions.c = mCQud
                                        Set f = MNew.MathFunction(mDF, AddressOf MMathFunctions.Quadratic)
    Case EMathFuncs.Cubic:              MMathFunctions.c = mCCub
                                        Set f = MNew.MathFunction(mDF, AddressOf MMathFunctions.Cubic)
    Case EMathFuncs.Sinus:              MMathFunctions.c = mCSin
                                        Set f = MNew.MathFunction(mDF, AddressOf MMathFunctions.Sinus)
    Case EMathFuncs.Exponent:           MMathFunctions.c = mCExp
                                        Set f = MNew.MathFunction(mDF, AddressOf MMathFunctions.Exponent)
    Case EMathFuncs.DamperedHarmonic:   MMathFunctions.c = mCDpH
                                        Set f = MNew.MathFunction(mDF, AddressOf MMathFunctions.DamperedHarmonic)
    End Select
    mView.DrawFunction f
End Sub
