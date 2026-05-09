Attribute VB_Name = "MNew"
Option Explicit

Public Function MathFunction(aDF As TDelegateFunction, ByVal pFncAddr As LongPtr) As MathFunction 'ICallDoubleReturnDouble
    Set MathFunction = New_DelegateFunction(aDF, pFncAddr)
End Function

Public Function MathFunctionView(Canvas As PictureBox) As MathFunctionView
    Set MathFunctionView = New MathFunctionView: MathFunctionView.New_ Canvas
End Function

'Public Function Guidlp(ByVal lpGuid As LongPtr) As Guid
'    Set Guidlp = New Guid: RtlMoveMemory ByVal Guidlp.Ptr, ByVal lpGuid, Guidlp.Size
'End Function
'
'Public Function GuidCo() As Guid
'    Set GuidCo = New Guid: GuidCo.NewCo
'End Function

