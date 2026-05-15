Attribute VB_Name = "MDelegateFunction"
Option Explicit
Private Type TDelegateFunctionVTable
    pVTIUOK   As TIUnknownVTable
    pDlgASM1  As LongPtr
    pVTIUFail As TIUnknownVTable
    pDlgASM2  As LongPtr
End Type
Private mIUVTable    As TDelegateFunctionVTable
Private mpVTableOK   As LongPtr
Private mpVTableFail As LongPtr '= VarPtr(mIUVTable)

Public Type TDelegateFunction
    pVTable     As LongPtr 'must be at Offset 0
    pFunction   As LongPtr 'must be at Offset 4
    'pThisObject As IUnknown
End Type

Private Type VBGuid
    Data1 As Long          ' 4
    Data2 As Integer       ' 2
    Data3 As Integer       ' 2
    Data5(0 To 7) As Byte  ' 8
End Type              'Sum: 16

'das ist illegal, hier muß natürlich ein VirtualAlloc und VirtualProtect her!
'die kurze ASM-Funktion nur 8-Bytes insgesamt passt komplett in einen Currency
Private m_VMem As VirtualMemory
Private m_pVMemDelegateASM As LongPtr

#If Win64 Then
    Private Const C_DelegateASM As Long = &HCC0861FF
    Private Const SizeOf_ASM As Long = 4
    'Private mDelegateASM As Long
#Else
    Private Const C_DelegateASM As Currency = -368956918007638.6215@
    Private Const SizeOf_ASM As Long = 8
    'Private mDelegateASM As Currency
#End If
#If VBA7 Then
    Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef pGuid As Any, ByVal lpsz As LongPtr, ByVal cchMax As Long) As Long
#Else
    Private Declare Function StringFromGUID2 Lib "ole32" (ByRef pGuid As Any, ByVal lpsz As LongPtr, ByVal cchMax As Long) As Long
#End If

Public Function New_DelegateFunction(this As TDelegateFunction, ByVal pFnc As LongPtr) As IUnknown
    Dim pAR As LongPtr
    Dim pRe As LongPtr
    If mpVTableOK = 0 Then
        'so jetzt muß man zuerst mal die Funktionspointer ermitteln
        'das wird aber nur einmal im Projekt gemacht!
        With mIUVTable
            'pAR = MPtr.FncPtr(AddressOf AddRefRelease)
            pAR = MPtr.FncPtr(AddressOf AddRef)
            pRe = MPtr.FncPtr(AddressOf Release)
            With .pVTIUOK
                .PQueryInterface = MPtr.FncPtr(AddressOf QueryInterfaceOK)
                .PAddRef = pAR
                '.PRelease = pAR
                .PRelease = pRe
            End With
            With .pVTIUFail
                .PQueryInterface = MPtr.FncPtr(AddressOf QueryInterfaceFail)
                .PAddRef = pAR
                '.PRelease = pAR
                .PRelease = pRe
            End With
            mpVTableOK = VarPtr(.pVTIUOK)
            mpVTableFail = VarPtr(.pVTIUFail)
            If m_pVMemDelegateASM = 0 Then
                Set m_VMem = New VirtualMemory
                m_pVMemDelegateASM = m_VMem.Alloc(SizeOf_ASM) 'LenB(C_DelegateASM))
                If m_pVMemDelegateASM = 0 Then
                    Debug.Print "!!! Achtung: m_pVMemDelegateASM = 0; !!!"
                    'Exit Function
                Else
                    RtlMoveMemory ByVal m_pVMemDelegateASM, ByVal VarPtr(C_DelegateASM), SizeOf_ASM 'LenB(C_DelegateASM)
                End If
                'mDelegateASM = C_DelegateASM
            End If
            .pDlgASM1 = m_pVMemDelegateASM 'VarPtr(mDelegateASM)
            .pDlgASM2 = .pDlgASM1
        End With
    End If
    With this
        .pFunction = pFnc
        .pVTable = mpVTableOK
    End With
    RtlMoveMemory New_DelegateFunction, VarPtr(this), MPtr.SizeOf_LongPtr
End Function

Private Function QueryInterfaceOK(this As TDelegateFunction, riid As LongPtr, pvObj As LongPtr) As Long 'Ptr
    pvObj = VarPtr(this)
    'Dim sGuid(0 To 76) As Byte
    'Dim hr As Long
    'hr = StringFromGUID2(riid, VarPtr(sGuid(0)), 76)
    'Debug.Print sGuid '{803D4140-C4E3-4699-8847-F8C07AD202CA}
    'this.pVTable = mpVTableFail
End Function
Private Function QueryInterfaceFail(this As TIUnknownVTable, riid As LongPtr, pvObj As LongPtr) As Long 'Ptr
    pvObj = VarPtr(this)
    'pvObj = 0
    'QueryInterfaceFail = E_NOINTERFACE
End Function
'Private Function AddRefRelease(this As LongPtr) As LongPtr
'    'hier wird nichts gemacht
'End Function
Private Function AddRef(this As LongPtr) As LongPtr
    'hier wird nichts gemacht
    'Debug.Print "AddRef"
End Function
Private Function Release(this As LongPtr) As LongPtr
    'hier wird nichts gemacht
    'Debug.Print "Release"
End Function


