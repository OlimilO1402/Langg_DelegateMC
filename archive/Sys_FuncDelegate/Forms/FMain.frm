VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FuncDelegate"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9555
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
   ScaleHeight     =   448
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   637
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox PictureBox1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      ScaleHeight     =   5715
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   480
      Width           =   9015
   End
   Begin VB.ComboBox CmbMathFunctions 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   2895
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

Private Sub Form_Load()

    ' irgendwelche Konstantenparameter voreinstellen
    ' wer will, kann sie hier mit der Hand anpassen
    ' (C(0), C(1), C(2), C(3))
    mCLin = Array(0.5, 15)
    mCQud = Array(0.01, 0, -100)
    mCCub = Array(0.0001, 0, -2, 0)
    mCSin = Array(0.05, 100, 20, 0)
    mCExp = Array(0.007, 0, 50, 0)
    mCDpH = Array(-0.007, 0, 25, 0, 0.1, 10, 0)

    PictureBox1.AutoRedraw = True
    PictureBox1.BackColor = vbWhite
    PictureBox1.ForeColor = vbBlack
    ' Alle Namen der Enumkonstanten in die ComboBox eintragen
    EMathFuncs_ToCBLB CmbMathFunctions

End Sub

Private Sub Form_Resize()
    
    Dim br As Single
    br = 8 '* Screen.TwipsPerPixelX
    
    With Me

        Call PictureBox1.Move(br, PictureBox1.Top, .ScaleWidth - PictureBox1.Left - br, .ScaleHeight - PictureBox1.Top - br)

    End With

    Set mView.Size = PictureBox1
    Call mView.DrawFunction(f)
    Call PictureBox1.Refresh

End Sub

Private Sub CmbMathFunctions_Click()

    Select Case CmbMathFunctions.ListIndex

    Case EMathFuncs.None
        Set f = Nothing

    Case EMathFuncs.Linear
        MMathFunctions.c = mCLin
        Set f = New_MathFunction(mDF, AddressOf MMathFunctions.Linear)

    Case EMathFuncs.Quadratic
        MMathFunctions.c = mCQud
        Set f = New_MathFunction(mDF, AddressOf MMathFunctions.Quadratic)

    Case EMathFuncs.Cubic
        MMathFunctions.c = mCCub
        Set f = New_MathFunction(mDF, AddressOf MMathFunctions.Cubic)

    Case EMathFuncs.Sinus
        MMathFunctions.c = mCSin
        Set f = New_MathFunction(mDF, AddressOf MMathFunctions.Sinus)

    Case EMathFuncs.Exponent
        MMathFunctions.c = mCExp
        Set f = New_MathFunction(mDF, AddressOf MMathFunctions.Exponent)

    Case EMathFuncs.DamperedHarmonic
        MMathFunctions.c = mCDpH
        Set f = New_MathFunction(mDF, AddressOf MMathFunctions.DamperedHarmonic)

    End Select

    Call mView.DrawFunction(f)
    Call PictureBox1.Refresh

End Sub
