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
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Col As Collection
Private m_Obj As Test1

Private Sub Form_Load()
    Set m_Col = New Collection
    m_Col.Add "eins", "eins"
    Dim s As String: s = m_Col.Item(1)
    Set m_Obj = New Test1
End Sub

Private Sub Command1_Click()
    'Early Binding
    Dim Obj As Test1
    Set Obj = New Test1
    Obj.MyPublicLongValue = 123
    MsgBox Obj.MyPublicLongValue
    Obj.MyPublicSub1
    Obj.PrivateLongValue = 456
    MsgBox Obj.PrivateLongValue
    MsgBox Obj.ThisIsMyFunction(123456789)
End Sub

Private Sub Command2_Click()
    'Late Binding
    Dim Obj As Object
    Set Obj = New Test1 'CreateObject("Test1")
    Obj.MyPublicLongValue = 123
    MsgBox Obj.MyPublicLongValue
    Obj.MyPublicSub1
    Obj.PrivateLongValue = 456
    MsgBox Obj.PrivateLongValue
    MsgBox Obj.ThisIsMyFunction(123456789)
End Sub

Private Sub Command3_Click()
    'Dim pCmdBtnFirstFunction As LongPtr: pCmdBtnFirstFunction = MPtr.ObjectAddressOf(m_Col, 0)
    
    Dim membNam As String
    
    membNam = "Item":   MsgBox "Collection." & membNam & " = " & MIDispatch.GetMemberID(m_Col, membNam) ' 0
    membNam = "Add":    MsgBox "Collection." & membNam & " = " & MIDispatch.GetMemberID(m_Col, membNam) ' 1
    membNam = "Count":  MsgBox "Collection." & membNam & " = " & MIDispatch.GetMemberID(m_Col, membNam) ' 2
    membNam = "Remove": MsgBox "Collection." & membNam & " = " & MIDispatch.GetMemberID(m_Col, membNam) ' 3
    
End Sub


Private Sub Command4_Click()
    
    m_Obj.MyPublicLongValue = 123
    m_Obj.PrivateLongValue = 456
    
    Dim membNam As String
    
    membNam = "MyPublicSub1":     MsgBox TypeName(m_Obj) & "." & membNam & " = &H" & Hex(MIDispatch.GetMemberID(m_Obj, membNam))
    membNam = "ThisIsMyFunction": MsgBox TypeName(m_Obj) & "." & membNam & " = &H" & Hex(MIDispatch.GetMemberID(m_Obj, membNam))
    membNam = "PrivateLongValue": MsgBox TypeName(m_Obj) & "." & membNam & " = &H" & Hex(MIDispatch.GetMemberID(m_Obj, membNam))
    
End Sub

Private Sub Command5_Click()
    
    Dim id As IDispatchCallable: Set id = New_IDispatch(m_Obj)
    
    MsgBox id.GetTypeInfoCount(0)
    
    
End Sub

