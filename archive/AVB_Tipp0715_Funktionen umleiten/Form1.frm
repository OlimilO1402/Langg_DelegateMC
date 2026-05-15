VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Form1"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MessageBoxA"
      Height          =   435
      Left            =   3300
      TabIndex        =   4
      Top             =   900
      Width           =   2115
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MessageBoxA hooken"
      Height          =   435
      Left            =   3300
      TabIndex        =   3
      Top             =   180
      Width           =   2115
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Shift Left Bench"
      Height          =   435
      Left            =   300
      TabIndex        =   0
      Top             =   900
      Width           =   1995
   End
   Begin VB.Label lblWAsm 
      AutoSize        =   -1  'True
      Caption         =   "mit asm:"
      Height          =   195
      Left            =   585
      TabIndex        =   2
      Top             =   540
      Width           =   570
   End
   Begin VB.Label lblWOAsm 
      AutoSize        =   -1  'True
      Caption         =   "ohne asm:"
      Height          =   195
      Left            =   420
      TabIndex        =   1
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function MessageBoxA Lib "user32" (ByVal hwnd As Long, ByVal strMsg As String, ByVal strTitle As String, ByVal style As Long) As Long

Private m_udtMsgBoxAHook As HookData

Private Sub Command1_Click()
    'Const iters As Long = 1000000
    Command1.Enabled = False
    ' Optimierter Maschinencode für Shift Left
    Dim udtAsm As MachineCode: udtAsm = QuickHook.ASMStringToMemory("8B4424048B4C2408D3E0C20800")
    ' Unoptimierte VB SHL Funktion testen
    Dim i As Long, x As Long, y As Long
    Dim dt As Single: dt = Timer
    Do
        x = ShiftLeft(2, 8)
        i = i + 1
    Loop While Timer - dt < 1
    lblWOAsm.Caption = "ohne asm: " & i & " Calls/Sekunde"
    ' Umleitung in VB Funktion auf eigenen Maschinencode
    Dim udtHook As HookData: udtHook = QuickHook.RedirectFunction(AddressOf ShiftLeft, True, udtAsm.pAsm)
    ' Optimierte SHL Funktion testen
    dt = Timer
    i = 0
    Do
        y = ShiftLeft(2, 8)
        i = i + 1
    Loop While Timer - dt < 1
    lblWAsm.Caption = "mit asm: " & i & " Calls/Sekunde"
    ' Für zusätzliche Geschwindigkeit zu N-Code kompilieren,
    ' dann fällt der VB Stub vor dem Hook weg
    Debug.Print "VBSHL(2,8)=" & x, "ASMSHL(2,8)=" & y
    ' VB Funktion wiederherstellen und Maschinencodespeicher freigeben
    If Not RestoreFunction(udtHook) Then Debug.Print "RestoreFunction fehlgeschlagen!"
    If Not FreeASMMemory(udtAsm) Then Debug.Print "Konnte ASM Speicher nicht freigeben!"
    Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    If Command2.Tag = "" Then
        ' user32.MessageBoxA auf TestModule.MessageBoxAHook umleiten
        m_udtMsgBoxAHook = QuickHook.RedirectFunction(GetWinAPIFunction("user32", "MessageBoxA"), False, AddressOf TestModule.MessageBoxAHook)
        
        If Not m_udtMsgBoxAHook.valid Then
            MsgBox "Hook fehlgeschlagen!", vbExclamation
        Else
            Command2.Tag = "h"
            Command2.Caption = "MessageBoxA enthooken"
        End If
    Else
        If Not RestoreFunction(m_udtMsgBoxAHook) Then
            MsgBox "Hook konnte nicht entfernt werden!", vbRetryCancel
        Else
            Command2.Tag = ""
            Command2.Caption = "MessageBoxA hooken"
        End If
    End If
End Sub

Private Sub Command3_Click()
    Debug.Print "MsgBoxA Result: " & MessageBoxA(0, "test msg", "titel", vbInformation)
End Sub

Private Sub Command4_Click()
    QuickHook.IsZeroBad
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If m_udtMsgBoxAHook.valid Then
        RestoreFunction m_udtMsgBoxAHook
    End If
End Sub
