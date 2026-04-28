VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MathFunctionView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private gr As PictureBox
'Public Enum VBRUN.ScaleModeConstants
'    vbUser = 0
'    vbTwips = 1
'    vbPoints = 2
'    vbPixels = 3
'    vbCharacters = 4
'    vbInches = 5
'    vbMillimeters = 6
'    vbCentimeters = 7
'    vbHimetric = 8
'    vbContainerPosition = 9
'    vbContainerSize = 10
'End Enum

Public Property Set Size(Value As PictureBox)
    
    Set gr = Value
    Call DrawAxis
    
    With gr
        'gr.Scale (0, .ScaleHeight / 2)-(.ScaleWidth / 2, 0)
        'gr.Scale (0, .ScaleHeight)-(.ScaleWidth, 0)
        gr.Scale (0, 0)-(.ScaleWidth, .ScaleWidth) ' .ScaleHeight)
        'mMtrx.Translate(.Width \ 2, .Height \ 2)
        'mMtrx.Scale(1.0F, -1.0F) ' (2.0F, -2.0F)
    End With
    
    
End Property

Private Sub DrawAxis()

    With gr
        
        .Cls
        '.Scale
        .ScaleMode = ScaleModeConstants.vbPixels ' 3
        '.ScaleMode = ScaleModeConstants.vbInches ' 3
        
        Set .Picture = Nothing
        
        'gr.DrawLine(Pens.Black, New Point(0, .Height \ 2), New Point(.Width, .Height \ 2))
        gr.Line (0, .Height \ 2)-(.Width, .Height \ 2), vbBlack
        'gr.DrawLine(Pens.Black, New Point(.Width \ 2, 0), New Point(.Width \ 2, .Height))
        gr.Line (.Width \ 2, 0)-(.Width \ 2, .Height), vbBlack
        
        ' das Hintergrundbild, die Koordinatenachsen werden nur einmal
        ' gezeichnet und in Image festgehalten.
        gr.Picture = gr.Image
        
    End With

End Sub

Public Sub DrawFunction(f As MathFunction)

    Dim xMin As Double: xMin = -gr.ScaleWidth / 2
    'Dim xMin As Long: xMin = -gr.ScaleWidth \ 2
    Dim xMax As Double: xMax = Math.Abs(xMin)
    Dim yMin As Double: yMin = -gr.ScaleHeight / 2
    Dim yMax As Double: yMax = Math.Abs(yMin)
    'Dim xMax As Long: xMax = Math.Abs(xMin)
    Dim x As Double, y As Double
    'Dim x As Long, y As Long
    Dim newPt As Point, oldPt As Point
    'Dim gr As Drawing.Graphics = Graphics.FromImage(Buffer)
    'gr.Transform = mMtrx
    'gr.sca
    If f Is Nothing Then Exit Sub
    
    Dim c As Long
Try: On Error GoTo Catch
    gr.Cls
    For x = xMin To xMax
        c = c + 1
        y = f.von(x) ' so hier ist also jetzt endlich das y = f(x), einfacher wirds nicht
        'newPt = New_Point(CLng(x + gr.Width / 2), CLng(y))
        newPt = New_Point(x, y)

        If x = xMin Then oldPt = newPt        ' noch bevor gezeichnet wird!
        'gr.DrawLine(Pens.Blue, oldPt, newPt)
        gr.Line (xMax + oldPt.x, yMax + oldPt.y)-(xMax + newPt.x, yMax + newPt.y), vbBlue
        
        oldPt = newPt
    Next
    
    Exit Sub
Catch:
    Debug.Print "Error " & xMin & " " & x & " " & xMax & " " & c
    'hier nur der Vollständigkeit halber
End Sub

' Das Programm ist recht einfach zu erweitern, die Prozedur zum zeichnen
' der Funktion muß für reelle Funktionen nicht verändert werden:
' 1. in der Klasse MathFunctions eine verallgemeinerte Funktion hinzufügen
' 2. das Enum EMathFuncs um den Namen der neuen Funktion erweitern
' 3. in der Form-Klasse ein List-Member hinzufügen (f. Konstantenparameter)
' 4. in Form_Load Konstantenparameter in die Liste schreiben.
' 5. in der Prozedur Form1.ChangeFunction einen neuen Case-Fall hinzufügen

' es wäre auch eine erweiterung für Parameterfunktionen denkbar
' dabei muß man eigentlich nur eine zusätzliche Membervariable für eine Funktion für x
' einführen bsp: fx(phi) dann ist die Funktion für y eben nicht von x abhängig sondern
' von einem variablen Parameter, ebenso die Funktion für x
' Es muß dann aber einen Neue Prozedur DrawFunction mit zwei Funktionsparameter
' programmiert werden.
' Sub DrawFunction(ByVal fx As MathFunction(Of Double), ByVal fy As MathFunction(Of Double))
' Außerdem muß vereinbart werden wie der Paramterwertebereich definiert wird.

