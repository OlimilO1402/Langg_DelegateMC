Attribute VB_Name = "TestModule"
Option Explicit

Private m_lngShift(31) As Long
Private i              As Long

' VB behandelt lngIn als vorzeichenbehaftet.
' Der ASM Code, mit dem die Funktion überschrieben wird,
' interpretiert lngIn dagegen als unsigned.
Public Function ShiftLeft(ByVal lngIn As Long, ByVal lngBits As Long) As Long
    If m_lngShift(0) = 0 Then
        m_lngShift(0) = 1
        For i = 1 To 30
            m_lngShift(i) = m_lngShift(i - 1) * 2
        Next
        m_lngShift(i) = &H80000000
    End If
    
    ShiftLeft = lngIn * m_lngShift(lngBits)
End Function

Public Function MessageBoxAHook(ByVal hwnd As LongPtr, ByVal msg As Long, ByVal title As Long, ByVal style As Long) As Long
    
    MessageBoxAHook = MsgBox("Gehookt!", style, "Am Haken")
    
End Function
