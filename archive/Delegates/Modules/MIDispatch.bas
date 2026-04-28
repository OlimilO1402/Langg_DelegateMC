Attribute VB_Name = "MIDispatch"
'***************************************************************
' (c) Copyright 2000 Matthew J. Curland
'
' This file is from the CD-ROM accompanying the book:
' Advanced Visual Basic 6: Power Techniques for Everyday Programs
'   Author: Matthew Curland
'   Published by: Addison-Wesley, July 2000
'   ISBN: 0-201-70712-8
'   http://www.PowerVB.com
'
' You are entitled to license free distribution of any application
'   that uses this file if you own a copy of the book, or if you
'   have obtained the file from a source approved by the author. You
'   may redistribute this file only with express written permission
'   of the author.
'
' This file depends on:
'   References:
'     VBoostTypes6.olb (VBoost Object Types (6.0))
'     VBoost6.Dll (VBoost Object Implementation (6.0)) (optional)
'   Files:
'     FunctionDelegator.bas
'     VBoost.Bas (optional)
'   Minimal VBoost conditionals:
'     VBOOST_INTERNAL = 1 : VBOOST_CUSTOM = 1
'   Conditional Compilation Values:
'     NOVBOOST = 1 'Removes VBoost dependency
'     FUNCTIONDELEGATOR_NOHEAP = 1 'Minimize FunctionDelegator impact
'
' This file is discussed in Chapter 4.
'***************************************************************

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDst As Any, ByRef pSrc As Any, ByVal BytLength As Long)
Private GUID_NULL As VBGUID
Private Const DISPID_PROPERTYPUT As Long = -3
Public Const VT_BYREF As Integer = &H4000

Public Enum DispInvokeFlags
    INVOKE_FUNC = 1
    INVOKE_PROPERTYGET = 2
    INVOKE_PROPERTYPUT = 4
    INVOKE_PROPERTYPUTREF = 8
End Enum

Public Function New_IDispatch(ByVal pObject As Object) As IDispatchCallable
    CopyMemory New_IDispatch, pObject, 4
    CopyMemory pObject, 0&, 4
End Function

Public Function GetMemberID(ByVal pObject As Object, Name As String) As Long
    
    Dim pCallDisp As IDispatchCallable
    CopyMemory pCallDisp, pObject, 4
    CopyMemory pObject, 0&, 4
    Dim hr As Long: hr = pCallDisp.GetIDsOfNames(GUID_NULL, VarPtr(Name), 1, 0, VarPtr(GetMemberID))
    If hr Then Err.Raise hr
    
End Function

Public Function CallInvoke(ByVal pObject As Object, ByVal MemberID As Long, ByVal InvokeKind As DispInvokeFlags, ParamArray ReverseArgList() As Variant) As Variant
    
    'Swap the ParamArray into a normal local variable. I'd like to use
    'DerefEBP to do this, but the stack offsets are 4 higher in the IDE
    'and in pcode than in a native Exe for functions in a standard module,
    'so you can't hard code the correct offset. The following line works
    'in native code, but not in the ide. DerefEBP.bas is required.
    'VBoost.AssignSwap ByVal VarPtrArray(pSAReverseArgList), ByVal DerefEBP.Call(24)
    
    Dim pArgList As Long: pArgList = (VarPtr(InvokeKind) Xor &H80000000) + 4 Xor &H80000000
    CopyMemory pArgList, ByVal pArgList, 4
    
    Dim pSAReverseArgList() As Variant
    CopyMemory ByVal ArrPtr(pSAReverseArgList), ByVal pArgList, 4
    CopyMemory ByVal pArgList, 0&, 4
    
    'Call the helper with pVarResult set to the address
    'of the return value of this function.
    
    CallInvokeHelper pObject, MemberID, InvokeKind, VarPtr(CallInvoke), pSAReverseArgList
    
End Function

Public Function CallInvokeArray(ByVal pObject As Object, ByVal MemberID As Long, ByVal InvokeKind As DispInvokeFlags, ReverseArgList() As Variant) As Variant
    
    'Call the helper with pVarResult set to the address
    'of the return value of this function.
    
    CallInvokeHelper pObject, MemberID, InvokeKind, VarPtr(CallInvokeArray), ReverseArgList
    
End Function

Public Sub CallInvokeSub(ByVal pObject As Object, ByVal MemberID As Long, ByVal InvokeKind As DispInvokeFlags, ParamArray ReverseArgList() As Variant)
    
    Dim pSAReverseArgList() As Variant
    'Swap the ParamArray into a normal local variable. I'd like to use
    'DerefEBP to do this, but the stack offsets are 4 higher in the IDE
    'and in pcode than in a native Exe for functions in a standard module,
    'so you can't hard code the correct offset. The following line works
    'in native code, but not in the ide. DerefEBP.bas is required.
    'VBoost.AssignSwap ByVal VarPtrArray(pSAReverseArgList), ByVal DerefEBP.Call(20)
    
    Dim pArgList As Long
    pArgList = (VarPtr(InvokeKind) Xor &H80000000) + 4 Xor &H80000000
    
    CopyMemory pArgList, ByVal pArgList, 4
    CopyMemory ByVal ArrPtr(pSAReverseArgList), ByVal pArgList, 4
    CopyMemory ByVal pArgList, 0&, 4
    
    'Call the helper with pVarResult set to 0
    CallInvokeHelper pObject, MemberID, InvokeKind, 0, pSAReverseArgList
End Sub

Public Sub CallInvokeSubArray(ByVal pObject As Object, ByVal MemberID As Long, ByVal InvokeKind As DispInvokeFlags, ReverseArgList() As Variant)
    
    'Call the helper with pVarResult set to 0
    CallInvokeHelper pObject, MemberID, InvokeKind, 0, ReverseArgList
    
End Sub

Private Sub CallInvokeHelper(pObject As Object, ByVal MemberID As Long, ByVal InvokeKind As Integer, ByVal pVarResult As Long, ReverseArgList() As Variant)
    
        
    'Fill in fields in the DISPPARAMS structure
    Dim lBoundArgs As Long: lBoundArgs = LBound(ReverseArgList)
    
    Dim Params As VBDISPPARAMS
    With Params
        
        .cArgs = UBound(ReverseArgList) - lBoundArgs + 1
        
        If .cArgs Then
            .rgvarg = VarPtr(ReverseArgList(lBoundArgs))
        End If
        
        If InvokeKind And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF) Then
            
            Dim dispidNamedArg As DISPID: dispidNamedArg = DISPID_PROPERTYPUT
            .cNamedArgs = 1
            .rgdispidNamedArgs = VarPtr(dispidNamedArg)
            'Make sure the RHS parameter is not VT_BYREF.
            VariantCopyInd ReverseArgList(lBoundArgs), ReverseArgList(lBoundArgs)
        End If
    End With
    
    Dim pCallDisp As IDispatchCallable
    'Get the incoming variable into a type we can call.
    CopyMemory pCallDisp, pObject, 4
    CopyMemory pObject, 0&, 4
    
    'Make the actual call
    Dim ExcepInfo As VBEXCEPINFO, uArgErr As UINT
    Dim hr As Long: hr = pCallDisp.Invoke(MemberID, GUID_NULL, 0, InvokeKind, Params, pVarResult, ExcepInfo, uArgErr)
    
    'Handle errors
    If hr = DISP_E_EXCEPTION Then
        
        'ExcepInfo has the information we need
        With ExcepInfo
            
            If .pfnDeferredFillIn Then
                Dim FDDeferred As FunctionDelegator
                Dim pFillExcepInfo As ICallDeferredFillIn
                Set pFillExcepInfo = InitDelegator(FDDeferred, .pfnDeferredFillIn)
                pFillExcepInfo.Fill ExcepInfo
                .pfnDeferredFillIn = 0
            End If
            Err.Raise .scode, .bstrSource, .bstrDescription, .bstrHelpFile, .dwHelpContext
        End With
    ElseIf hr Then
        
        Err.Raise hr
        
    End If
    
End Sub


