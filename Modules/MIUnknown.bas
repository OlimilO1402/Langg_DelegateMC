Attribute VB_Name = "MIUnknown"
Option Explicit
'wird von QueryInterface zurückgegeben, falls das Objekt kein Interface hat:
Public Const E_NOINTERFACE As Long = &H80004002

'dies ist der typische VTable der Schnittstelle IUnknown
Public Type TIUnknownVTable
    PQueryInterface As LongPtr
    PAddRef         As LongPtr
    PRelease        As LongPtr
End Type
'auch bekannt als Alias MoveMemory, Alias CopyMemory, Alias cpymem etc.
Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bytLength As Long)

'die folgenden drei Funktionsrümpfe kann man so übernehmen, und werden
'in jedem lightweight Objekt gebraucht.
Private Function QueryInterface(this As TIUnknownVTable, riid As LongPtr, pvObj As LongPtr) As LongPtr
    'pvObj = 0
    'bei Objekten die kein Interface haben:
    'QueryInterface = E_NOINTERFACE
End Function
Private Function AddRef(this As TIUnknownVTable) As LongPtr
    'hier wird eine Referenz hinzugefügt
End Function
Private Function Release(this As TIUnknownVTable) As LongPtr
    'hier wird eine Referenz abgezogen
End Function


