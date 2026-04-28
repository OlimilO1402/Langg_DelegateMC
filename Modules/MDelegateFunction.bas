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

'das ist illegal, hier muß natürlich ein VirtualAlloc und VirtualProtect her!
'die kurze ASM-Funktion nur 8-Bytes insgesamt passt komplett in einen Currency
Private Const C_DelegateASM As Currency = -368956918007638.6215@
Private mDelegateASM As Currency

Public Function New_DelegateFunction(this As TDelegateFunction, ByVal pFnc As LongPtr) As IUnknown
    Dim pAR As LongPtr
    If mpVTableOK = 0 Then
        'so jetzt muß man zuerst mal die Funktionspointer ermitteln
        'das wird aber nur einmal im Projekt gemacht!
        With mIUVTable
            pAR = MPtr.FncPtr(AddressOf AddRefRelease)
            With .pVTIUOK
                .PQueryInterface = MPtr.FncPtr(AddressOf QueryInterfaceOK)
                .PAddRef = pAR
                .PRelease = pAR
            End With
            With .pVTIUFail
                .PQueryInterface = MPtr.FncPtr(AddressOf QueryInterfaceFail)
                .PAddRef = pAR
                .PRelease = pAR
            End With
            mpVTableOK = VarPtr(.pVTIUOK)
            mpVTableFail = VarPtr(.pVTIUFail)
            mDelegateASM = C_DelegateASM
            .pDlgASM1 = VarPtr(mDelegateASM)
            .pDlgASM2 = .pDlgASM1
        End With
    End If
    With this
        .pFunction = pFnc
        .pVTable = mpVTableOK
    End With
    Call RtlMoveMemory(New_DelegateFunction, VarPtr(this), 4)
    
End Function

Private Function QueryInterfaceOK(this As TDelegateFunction, riid As LongPtr, pvObj As LongPtr) As LongPtr
    pvObj = VarPtr(this)
    this.pVTable = mpVTableFail
End Function
Private Function QueryInterfaceFail(this As TIUnknownVTable, riid As LongPtr, pvObj As LongPtr) As LongPtr
    pvObj = 0
    QueryInterfaceFail = E_NOINTERFACE
End Function
Private Function AddRefRelease(this As LongPtr) As LongPtr
    'hier wird nichts gemacht
End Function
