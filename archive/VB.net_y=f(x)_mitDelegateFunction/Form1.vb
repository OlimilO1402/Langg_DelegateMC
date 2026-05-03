Option Explicit On
Option Strict On

Public Enum EMathFuncs
    None = 0
    Linear = 1
    Quadratic
    Cubic
    Sinus
    Exponent
    DamperedHarmonic
End Enum

Public Delegate Function MathFunction(Of T)(ByVal x As T) As T

Public Class Form1

    Private f As MathFunction(Of Double)
    Private mMathfuncs As New MathFunctions
    Private mCLin As New List(Of Double) 'Konstantenparameter für eine lineare Funktion
    Private mCQud As New List(Of Double) 'Konstantenparameter für eine quadratische Funktion
    Private mCCub As New List(Of Double) 'Konstantenparameter für eine kubische Funktion
    Private mCSin As New List(Of Double) 'Konstantenparameter für eine Sinus-Funktion
    Private mCExp As New List(Of Double) 'Konstantenparameter für eine Expon-Funktion
    Private mCDpH As New List(Of Double) 'Konstantenparameter für eine ged. Schwingung
    Private mView As New MathFunctionView

    Private Sub Form1_Load(ByVal sender As System.Object, _
                           ByVal e As System.EventArgs) Handles MyBase.Load
        ' irgendwelche Konstantenparameter voreinstellen
        ' wer will, kann sie hier mit der Hand anpassen
        '                           {C(0), C(1), C(2), C(3)}
        mCLin.AddRange(New Double() {0.5, 15})
        mCQud.AddRange(New Double() {0.01, 0, -100})
        mCCub.AddRange(New Double() {0.0001, 0, -2, 0})
        mCSin.AddRange(New Double() {0.05, 100, 20, 0})
        mCExp.AddRange(New Double() {0.007, 0, 50, 0})
        mCDpH.AddRange(New Double() {-0.007, 0, 25, 0, 0.1, 10, 0})
        ' Alle Namen der Enumkonstanten in die ComboBox eintragen
        With CmbMathFunctions
            With .Items
                For Each s As String In [Enum].GetNames(GetType(EMathFuncs))
                    .Add(s)
                Next
            End With
            'das erste Element der ComboBox auswählen
            .SelectedIndex = 1
        End With
    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, _
                             ByVal e As System.EventArgs) Handles Me.Resize
        With Me.ClientSize
            PictureBox1.Size = New Size(.Width - PictureBox1.Left - 12, _
                                        .Height - PictureBox1.Top - 12)
        End With
        mView.Size = PictureBox1.Size
        mView.DrawFunction(f)
        PictureBox1.Refresh()
    End Sub

    Private Sub ChangeFunction(ByVal sender As System.Object, _
                               ByVal e As System.EventArgs) _
                               Handles CmbMathFunctions.SelectedIndexChanged
        Select Case DirectCast(CmbMathFunctions.SelectedIndex, EMathFuncs)
            Case EMathFuncs.None
                f = Nothing
            Case EMathFuncs.Linear
                mMathfuncs.C = mCLin
                f = New MathFunction(Of Double)(AddressOf mMathfuncs.Linear)
            Case EMathFuncs.Quadratic
                mMathfuncs.C = mCQud
                f = New MathFunction(Of Double)(AddressOf mMathfuncs.Quadratic)
            Case EMathFuncs.Cubic
                mMathfuncs.C = mCCub
                f = New MathFunction(Of Double)(AddressOf mMathfuncs.Cubic)
            Case EMathFuncs.Sinus
                mMathfuncs.C = mCSin
                f = New MathFunction(Of Double)(AddressOf mMathfuncs.Sinus)
            Case EMathFuncs.Exponent
                mMathfuncs.C = mCExp
                f = New MathFunction(Of Double)(AddressOf mMathfuncs.Exponent)
            Case EMathFuncs.DamperedHarmonic
                mMathfuncs.C = mCDpH
                f = New MathFunction(Of Double)(AddressOf mMathfuncs.DamperedHarmonic)
        End Select
        mView.DrawFunction(f)
        PictureBox1.Refresh()
    End Sub

    Private Sub PictureBox1_Paint(ByVal sender As Object, _
                                  ByVal e As System.Windows.Forms.PaintEventArgs) _
                                  Handles PictureBox1.Paint
        If mView.Buffer IsNot Nothing Then
            e.Graphics.DrawImage(mView.Buffer, 0, 0)
        Else
            e.Graphics.Clear(PictureBox1.BackColor)
        End If

    End Sub

End Class

Class MathFunctions

    Public C As List(Of Double) ' Konstantenparameter

    'hier nur zum Probieren vier Beispielfunktionen
    Public Function Linear(ByVal x As Double) As Double
        Return C(0) * x + C(1)
    End Function

    Public Function Quadratic(ByVal x As Double) As Double
        Return Linear(x) * x + C(2)
    End Function

    Public Function Cubic(ByVal x As Double) As Double
        Return Quadratic(x) * x + C(3)
    End Function
    '... noch mehr Polynome?

    Public Function Sinus(ByVal x As Double) As Double
        Return Math.Sin(Linear(x)) * C(2) + C(3)
    End Function

    Public Function Exponent(ByVal x As Double) As Double
        Return Math.Exp(Linear(x)) * C(2) + C(3)
    End Function

    Public Function DamperedHarmonic(ByVal x As Double) As Double
        Return Exponent(x) * Math.Sin(C(4) * x + C(5)) + C(6)
    End Function

    'Man müßte hier noch viel freier sein dürfen,
    'und die Funktionen verketten können
    'y = f(f(f(x)))
    '...

End Class

Class MathFunctionView

    Public Buffer As Bitmap 'wird neu gezeichnet
    Private mAxis As Bitmap 'statisch
    Private mMtrx As Drawing2D.Matrix
    Private mSize As Size

    Public WriteOnly Property Size() As Size
        Set(ByVal value As Size)
            mSize = value
            mMtrx = New System.Drawing.Drawing2D.Matrix()
            With mSize
                mMtrx.Translate(.Width \ 2, .Height \ 2)
                mMtrx.Scale(1.0F, -1.0F) ' (2.0F, -2.0F)
            End With
            DrawAxis()
        End Set
    End Property

    Private Sub DrawAxis()
        With mSize
            mAxis = New Bitmap(.Width, .Height)
            Dim gr As Drawing.Graphics = Graphics.FromImage(mAxis)
            gr.DrawLine(Pens.Black, New Point(0, .Height \ 2), _
                                    New Point(.Width, .Height \ 2))
            gr.DrawLine(Pens.Black, New Point(.Width \ 2, 0), _
                                    New Point(.Width \ 2, .Height))
        End With
    End Sub

    Public Sub DrawFunction(ByVal f As MathFunction(Of Double))
        Buffer = New Bitmap(mAxis)
        Dim xMin As Double = -mSize.Width \ 2
        Dim xMax As Double = Math.Abs(xMin)
        Dim x, y As Double
        Dim newPt, oldPt As Point
        Dim gr As Drawing.Graphics = Graphics.FromImage(Buffer)
        gr.Transform = mMtrx
        If f Is Nothing Then Exit Sub
        Try
            For x = xMin To xMax
                y = f(x) 'so hier ist also jetzt endlich das y = f(x), einfacher wirds nicht
                newPt = New Point(CInt(x), CInt(y))
                If x = xMin Then oldPt = newPt 'noch bevor gezeichnet wird!
                gr.DrawLine(Pens.Blue, oldPt, newPt)
                oldPt = newPt
            Next
        Catch
            '
        End Try
    End Sub

    'Das Programm ist recht einfach zu erweitern, die Prozedur zum zeichnen 
    'der Funktion muß für reelle Funktionen nicht verändert werden:
    '1. in der Klasse MathFunctions eine verallgemeinerte Funktion hinzufügen
    '2. das Enum EMathFuncs um den Namen der neuen Funktion erweitern
    '3. in der Form-Klasse ein List-Member hinzufügen (f. Konstantenparameter)
    '4. in Form_Load Konstantenparameter in die Liste schreiben.
    '5. in der Prozedur Form1.ChangeFunction einen neuen Case-Fall hinzufügen

    'es wäre auch eine erweiterung für Parameterfunktionen denkbar 
    'dabei muß man eigentlich nur eine zusätzliche Membervariable für eine Funktion für x 
    'einführen bsp: fx(phi) dann ist die Funktion für y eben nicht von x abhängig sondern 
    'von einem variablen Parameter, ebenso die Funktion für x
    'Es muß dann aber einen Neue Prozedur DrawFunction mit zwei Funktionsparameter 
    'programmiert werden.
    'Sub DrawFunction(ByVal fx As MathFunction(Of Double), ByVal fy As MathFunction(Of Double))
    'Außerdem muß vereinbart werden wie der Paramterwertebereich definiert wird.

End Class