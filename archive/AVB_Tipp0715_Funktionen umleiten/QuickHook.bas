Attribute VB_Name = "QuickHook"
Option Explicit

#If VBA7 = 0 Then
    Public Enum LongPtr
        [_]
    End Enum
#End If

Private Const MEM_COMMIT             As Long = &H1000&
Private Const MEM_DECOMMIT           As Long = &H4000&
Private Const PAGE_EXECUTE           As Long = &H10&
Private Const PAGE_EXECUTE_READ      As Long = &H20&
Private Const PAGE_EXECUTE_READWRITE As Long = &H40&

#If Win64 Then
    Private Const SizeOf_LongPtr     As Long = 8
    Private Const IDE_ADDROF_REL     As Long = 55
#Else
    Private Const SizeOf_LongPtr     As Long = 4
    Private Const IDE_ADDROF_REL     As Long = 22
#End If

Private Const ASMSIZE                As Long = 5

Public Type HookData
    pFunction   As LongPtr  ' Pointer zur umzuleitenden Stelle
    pNewFnc     As LongPtr  ' Umleitungsziel
    cHookSize   As Long     ' Größe des Hooks
    pBackup     As LongPtr  ' Pointer zu gesicherten Bytes
    cBackupSize As Long     ' Menge an gesicherten Bytes
    valid       As Boolean  ' Hook funktional?
End Type

Public Type MachineCode
    pAsm        As LongPtr  ' Pointer zum Code
    cSize       As Long     ' Größe des Codes in Bytes
    valid       As Boolean  ' gültig?
End Type

#If VBA7 Then
    Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
    Private Declare PtrSafe Function VirtualFree Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
    Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
    Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal cBytes As Long)
    Private Declare PtrSafe Sub RtlFillMemory Lib "kernel32" (pDst As Any, ByVal cBytes As Long, ByVal char As Byte)
    Private Declare PtrSafe Function IsBadCodePtr Lib "kernel32" (ByVal addr As Long) As Long
    Private Declare PtrSafe Function LoadLibraryA Lib "kernel32" (ByVal strPath As String) As Long
    Private Declare PtrSafe Function GetModuleHandleA Lib "kernel32" (ByVal strModule As String) As LongPtr
    Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal strName As String) As LongPtr
#Else
    Private Declare Function VirtualAlloc Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
    Private Declare Function VirtualFree Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
    Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
    Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal cBytes As Long)
    Private Declare Sub RtlFillMemory Lib "kernel32" (pDst As Any, ByVal cBytes As Long, ByVal char As Byte)
    Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal addr As LongPtr) As Long
    Private Declare Function LoadLibraryA Lib "kernel32" (ByVal strPath As String) As Long
    Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal strModule As String) As LongPtr
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal strName As String) As LongPtr
#End If

Public Function GetWinAPIFunction(ByVal strLib As String, ByVal strFncName As String) As LongPtr
    Dim hModule As LongPtr: hModule = GetModuleHandleA(strLib)
    If hModule = 0 Then
        hModule = LoadLibraryA(strLib)
        If hModule = 0 Then Exit Function
    End If
    GetWinAPIFunction = GetProcAddress(hModule, strFncName)
End Function

' alloziert ausführbaren Speicher und schreibt
' als Hex String übergebenen Maschinencode hinein
Public Function ASMStringToMemory(ByVal strAsm As String) As MachineCode
    Dim i As Long, u As Long: u = Len(strAsm) \ 2 - 1
    ReDim btAsm(0 To u) As Byte
    For i = 0 To u: btAsm(i) = CByte("&H" & Mid$(strAsm, i * 2 + 1, 2)): Next
    With ASMStringToMemory
        .cSize = UBound(btAsm) + 1
        .pAsm = VirtualAlloc(ByVal 0&, .cSize, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        If .pAsm <> 0 Then
            RtlMoveMemory ByVal .pAsm, btAsm(0), .cSize
            .valid = True
        End If
    End With
End Function

Public Function FreeASMMemory(asm As MachineCode) As Boolean
    If asm.valid Then
        asm.valid = False
        FreeASMMemory = VirtualFree(ByVal asm.pAsm, asm.cSize, MEM_DECOMMIT) <> 0
    End If
End Function

' Mit Jmp Instruktion überschriebene Funktion wiederherstellen
Public Function RestoreFunction(hook As HookData) As Boolean
    If hook.valid Then
        Dim lngRet As Long, lngOldProtection As Long
        lngRet = VirtualProtect(ByVal hook.pFunction, hook.cHookSize, PAGE_EXECUTE_READWRITE, lngOldProtection)
        If lngRet = 0 Then Exit Function
        RtlMoveMemory ByVal hook.pFunction, ByVal hook.pBackup, ByVal hook.cBackupSize
        VirtualProtect ByVal hook.pFunction, hook.cHookSize, lngOldProtection, 0&
        VirtualFree ByVal hook.pBackup, hook.cBackupSize, MEM_DECOMMIT
        hook.valid = False
        RestoreFunction = True
    End If
End Function

' Funktion mit Jmp Instruktion überschreiben,
' mit Unterstützung für VB 6 IDE
Public Function RedirectFunction(ByVal addr_in As LongPtr, ByVal isVBModule As Boolean, ByVal addr_out As LongPtr) As HookData
    
    
    Dim btAsm(ASMSIZE - 1)  As Byte
    
    If isVBModule Then addr_in = FncPtr(addr_in)
    Dim lpBackupMemory As Long: lpBackupMemory = VirtualAlloc(ByVal 0&, ASMSIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If lpBackupMemory = 0 Then Exit Function
    
    Dim lngOldInProtection  As Long
    Dim lngRet As Long: lngRet = VirtualProtect(ByVal addr_in, ASMSIZE, PAGE_EXECUTE_READWRITE, lngOldInProtection)
    If lngRet = 0 Then
        VirtualFree ByVal lpBackupMemory, ASMSIZE, MEM_DECOMMIT
        Exit Function
    End If
    
    RtlMoveMemory ByVal lpBackupMemory, ByVal addr_in, ASMSIZE
    
    Dim lngJmp As Long: lngJmp = addr_out - addr_in - ASMSIZE
    
    btAsm(0) = &HE9
    RtlMoveMemory btAsm(1), lngJmp, SizeOf_LongPtr
    
    RtlMoveMemory ByVal addr_in, btAsm(0), ASMSIZE
    
    lngRet = VirtualProtect(ByVal addr_in, ASMSIZE, lngOldInProtection, 0&)
'    If lngRet = 0 Then
'        VirtualFree ByVal lngBackupMemory, ASMSIZE, MEM_DECOMMIT
'        Exit Function
'    End If
    
    With RedirectFunction
        .pFunction = addr_in
        .pNewFnc = addr_out
        .pBackup = lpBackupMemory
        .cBackupSize = ASMSIZE
        .cHookSize = ASMSIZE
        .valid = True
    End With
End Function

Public Function FncPtrStub(ByVal pFunction As LongPtr) As LongPtr
    FncPtrStub = pFunction
End Function

Public Function FncPtr(ByVal pFunction As LongPtr) As LongPtr
    FncPtr = pFunction
    If Not IsInIDE() Then Exit Function
    ' Wird das Programm in der VB-IDE gestartet, befindet sich der eigentliche Zeiger auf eine Funktion
    ' bei (AddressOf X) + 22, AddressOf X selber zeigt nur auf einen Stub. (getestet mit VB 6)
    RtlMoveMemory FncPtr, ByVal pFunction + IDE_ADDROF_REL, SizeOf_LongPtr
    If IsBadCodePtr(FncPtr) Then FncPtr = pFunction
End Function

' http://www.activevb.de/tipps/vb6tipps/tipp0347.html
Public Function IsInIDE() As Boolean
Try: On Error GoTo Catch
    Debug.Print 1 / 0
    Exit Function
Catch: IsInIDE = True
End Function
