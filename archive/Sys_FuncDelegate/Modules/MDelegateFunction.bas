Attribute VB_Name = "MDelegateFunction"
Option Explicit
Private Type TDelegateFunctionVTable
    pVTIUOK   As TIUnknownVTable
    pDlgASM1  As Long
    pVTIUFail As TIUnknownVTable
    pDlgASM2  As Long
End Type
Private mIUVTable    As TDelegateFunctionVTable
Private mpVTableOK   As Long
Private mpVTableFail As Long '= VarPtr(mIUVTable)

Public Type TDelegateFunction
    pVTable     As Long 'must be at Offset 0
    pFunction   As Long 'must be at Offset 4
    'pThisObject As IUnknown
End Type

'das ist illegal, hier muß natürlich ein VirtualAlloc und VirtualProtect her!
'die kurze ASM-Funktion nur 8-Bytes insgesamt passt komplett in einen Currency
Private Const C_DelegateASM As Currency = -368956918007638.6215@
Private mDelegateASM As Currency

Public Function New_DelegateFunction(this As TDelegateFunction, ByVal pFnc As Long) As IUnknown
    Dim pAR As Long
    If mpVTableOK = 0 Then
        'so jetzt muß man zuerst mal die Funktionspointer ermitteln
        'das wird aber nur einmal im Projekt gemacht!
        With mIUVTable
            pAR = FncPtr(AddressOf AddRefRelease)
            With .pVTIUOK
                .PQueryInterface = FncPtr(AddressOf QueryInterfaceOK)
                .PAddRef = pAR
                .PRelease = pAR
            End With
            With .pVTIUFail
                .PQueryInterface = FncPtr(AddressOf QueryInterfaceFail)
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

Private Function QueryInterfaceOK(this As TDelegateFunction, riid As Long, pvObj As Long) As Long
    pvObj = VarPtr(this)
    this.pVTable = mpVTableFail
End Function
Private Function QueryInterfaceFail(this As TIUnknownVTable, riid As Long, pvObj As Long) As Long
    pvObj = 0
    QueryInterfaceFail = E_NOINTERFACE
End Function
Private Function AddRefRelease(this As Long) As Long
    'hier wird nichts gemacht
End Function
