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
Private WithEvents mCanvas As PictureBox
Attribute mCanvas.VB_VarHelpID = -1

Friend Sub New_(Canvas As PictureBox)
    Set mCanvas = Canvas
    mCanvas.AutoRedraw = True
    mCanvas.BackColor = vbWhite
    mCanvas.ForeColor = vbBlack
    'mCanvas.ScaleMode = ScaleModeConstants.vbMillimeters
    'mCanvas.ScaleMode = ScaleModeConstants.vbPixels
End Sub

Private Sub mCanvas_Resize()
'Public Property Set Size(Value As PictureBox)
    
    'With mCanvas
        'gr.Scale (0, .ScaleHeight / 2)-(.ScaleWidth / 2, 0)
        'gr.Scale (0, .ScaleHeight)-(.ScaleWidth, 0)
    '    mCanvas.Scale (0, 0)-(.ScaleWidth, .ScaleHeight) ' .ScaleHeight)
        'mMtrx.Translate(.Width \ 2, .Height \ 2)
        'mMtrx.Scale(1.0F, -1.0F) ' (2.0F, -2.0F)
    'End With
    'mCanvas.Cls
    If mCanvas.ScaleWidth = 0 Or mCanvas.ScaleHeight = 0 Then Exit Sub
    mCanvas.ScaleMode = ScaleModeConstants.vbCentimeters
    mCanvas.Scale (-mCanvas.ScaleWidth / 2, -mCanvas.ScaleHeight / 2)-(mCanvas.ScaleWidth / 2, mCanvas.ScaleHeight / 2)
    mCanvas.Cls
    
    DrawAxes
    
'End Property
End Sub

Private Sub DrawAxes()
    
    With mCanvas
        .Cls
        '.Scale
        '.ScaleMode = ScaleModeConstants.vbPixels ' 3
        '.ScaleMode = ScaleModeConstants.vbInches ' 3
        '.ScaleMode = ScaleModeConstants.vbMillimeters ' 3
        '.ScaleMode = ScaleModeConstants.vbCentimeters
        Set .Picture = Nothing
        .ForeColor = vbBlack
        
        'gr.DrawLine(Pens.Black, New Point(0, .Height \ 2), New Point(.Width, .Height \ 2))
        'mCanvas.Line (0, .ScaleHeight / 2)-(.ScaleWidth, .ScaleHeight / 2), vbBlack
        'gr.DrawLine(Pens.Black, New Point(.Width \ 2, 0), New Point(.Width \ 2, .Height))
        'mCanvas.Line (.ScaleWidth / 2, 0)-(.ScaleWidth / 2, .ScaleHeight), vbBlack
        mCanvas.Line (-mCanvas.ScaleWidth / 2, 0)-(mCanvas.ScaleWidth / 2, 0)
        mCanvas.Line (0, -mCanvas.ScaleHeight / 2)-(0, mCanvas.ScaleHeight / 2)
        Dim i As Long, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, s As String
        For i = -mCanvas.ScaleWidth / 2 To mCanvas.ScaleWidth / 2
            X1 = i:  Y1 = 0:  X2 = X1: Y2 = 0.2
            mCanvas.Line (X1, Y1)-(X2, Y2)
            s = i
            mCanvas.CurrentX = X1 - mCanvas.TextWidth(s) / 2
            If i = 0 Then
                mCanvas.CurrentX = -mCanvas.TextWidth(s) / 2
                mCanvas.CurrentY = -mCanvas.TextHeight(s) / 2
            End If
            mCanvas.Print s
        Next
        For i = -mCanvas.ScaleHeight / 2 To mCanvas.ScaleHeight / 2
            X1 = 0:   Y1 = i:  X2 = 0.2: Y2 = Y1
            mCanvas.Line (X1, Y1)-(X2, Y2)
            s = -i
            Y1 = Y1 - mCanvas.TextHeight(s) / 2
            mCanvas.CurrentY = Y1 ',Y1
            If i <> 0 Then
                mCanvas.Print s
            End If
        Next
        ' das Hintergrundbild, die Koordinatenachsen werden nur einmal
        ' gezeichnet und in Image festgehalten.
        Set .Picture = mCanvas.Image
    End With
End Sub

Public Sub DrawFunction(f As MathFunction)
    Dim xMin As Double: xMin = CLng(-mCanvas.ScaleWidth / 2)
    'Dim xMin As Long: xMin = -gr.ScaleWidth \ 2
    Dim xMax As Double: xMax = Math.Abs(xMin)
    Dim yMin As Double: yMin = CLng(-mCanvas.ScaleHeight / 2)
    Dim yMax As Double: yMax = Math.Abs(yMin)
    'Dim xMax As Long: xMax = Math.Abs(xMin)
    Dim x As Double, y As Double
    'Dim x As Long, y As Long
    Dim newPt As Point, oldPt As Point
    'Dim gr As Drawing.Graphics = Graphics.FromImage(Buffer)
    'gr.Transform = mMtrx
    'gr.sca
    
    Dim c As Long
Try: On Error GoTo Catch
    mCanvas.Cls
    If f Is Nothing Then Exit Sub
    For x = xMin To xMax Step 0.1
        c = c + 1
        y = f.von(x) ' so hier ist also jetzt endlich das y = f(x), einfacher wirds nicht
        If y < 1000 Then
            'newPt = New_Point(CLng(x + gr.Width / 2), CLng(y))
            newPt = New_Point(x, y)
    
            If x = xMin Then oldPt = newPt        ' noch bevor gezeichnet wird!
            'gr.DrawLine(Pens.Blue, oldPt, newPt)
            mCanvas.Line (oldPt.x, oldPt.y)-(newPt.x, newPt.y), vbBlue
            
            oldPt = newPt
        End If
    Next
    
    Exit Sub
Catch:
    Debug.Print "Error xMin:" & xMin & "; xMax: " & xMax & "; x:" & x & "; y:" & y '& c
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
