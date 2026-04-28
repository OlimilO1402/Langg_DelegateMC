VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Test2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MyPublicLongValue As Long
Private m_PrivateLongValue As Long

Private Sub Class_Initialize()
    m_PrivateLongValue = 65536
End Sub

Public Sub MyPublicSub1()
    MsgBox "MyPublicSub1"
End Sub

Public Function ThisIsMyFunction(ByVal Param1 As Long) As Long
    ThisIsMyFunction = Param1
End Function

Public Property Get PrivateLongValue() As Long
    PrivateLongValue = m_PrivateLongValue
    MsgBox "PrivateLongValue = m_PrivateLongValue"
End Property

Public Property Let PrivateLongValue(ByVal Value As Long)
    m_PrivateLongValue = Value
    MsgBox "m_PrivateLongValue = " & Value
End Property


