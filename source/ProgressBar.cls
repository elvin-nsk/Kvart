VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements IProgressBar

'===============================================================================

Private Enum ProgressBarUnit
    pbDecimal
    pbNumeric
    pbPercent
End Enum

Private Type typeThis
    Cancelable As Boolean
    Canceled As Boolean
    Iteration As Long
    Max As Long
    Step As Long
    Unit As ProgressBarUnit
End Type
Private This As typeThis

Private WithEvents View As ProgressBarView
Attribute View.VB_VarHelpID = -1

'===============================================================================

'прогресс в виде десятичной дроби
Public Function CreateDecimal() As IProgressBar
    With New ProgressBar
        Set CreateDecimal = .Init(pbDecimal, 1, 0.1)
    End With
End Function

'прогресс в виде текущей итерации из максимальных
Public Function CreateNumeric( _
                    ByVal Max As Long, _
                    Optional ByVal Step As Long = 1 _
                ) As IProgressBar
    With New ProgressBar
        Set CreateNumeric = .Init(pbNumeric, Max, 1)
    End With
End Function

'прогресс в виде процентов
Public Function CreatePercent() As IProgressBar
    With New ProgressBar
        Set CreatePercent = .Init(pbPercent, 100, 1)
    End With
End Function

Friend Function Init( _
                    ByVal Unit As ProgressBarUnit, _
                    ByVal Max As Long, _
                    ByVal Step As Long _
                ) As IProgressBar
    With This
        .Unit = Unit
        If IsValid(Max) Then .Max = Max Else Throw
        .Step = Step
        .Iteration = 0
        Set View = New ProgressBarView
    End With
    Set Init = Me
End Function

Private Sub Class_Terminate()
    DisposeView
End Sub

'===============================================================================

Private Property Let IProgressBar_Cancelable(ByVal RHS As Boolean)
    This.Cancelable = RHS
    View.Cancelable = RHS
End Property
Private Property Get IProgressBar_Cancelable() As Boolean
    IProgressBar_Cancelable = This.Cancelable
End Property

Private Property Get IProgressBar_Canceled() As Boolean
    IProgressBar_Canceled = This.Canceled
End Property

Private Property Let IProgressBar_Caption(ByVal RHS As String)
    View.Caption = RHS
End Property
Private Property Get IProgressBar_Caption() As String
    IProgressBar_Caption = View.Caption
End Property

Private Property Let IProgressBar_CancelButtonCaption(ByVal RHS As String)
    View.CancelButtonCaption = RHS
End Property
Private Property Get IProgressBar_CancelButtonCaption() As String
    IProgressBar_CancelButtonCaption = View.CancelButtonCaption
End Property

Private Property Get IProgressBar_Max() As Long
    IProgressBar_Max = This.Max
End Property

Private Property Let IProgressBar_NumericMiddleText(ByVal RHS As String)
    View.NumericMiddleText = RHS
End Property
Private Property Get IProgressBar_NumericMiddleText() As String
    IProgressBar_NumericMiddleText = View.NumericMiddleText
End Property

Private Property Get IProgressBar_Step() As Long
    IProgressBar_Step = This.Step
End Property

Private Sub IProgressBar_Update()
    ShowView
    Dim CurrentValue As Long
    With This
        .Iteration = .Iteration + 1
        CurrentValue = .Iteration * .Step
        View.UpdateTo CurrentValue, .Max, Not (.Unit = pbNumeric)
    End With
End Sub

Private Sub IProgressBar_UpdateTo(ByVal RHS As Variant)
    ShowView
    With This
        If Not IsValid(RHS) Then Exit Sub
        Select Case .Unit
        Case pbDecimal
            View.UpdateTo VBA.CLng(RHS / 100), 100
        Case pbNumeric
            View.UpdateTo VBA.CLng(RHS), VBA.CLng(.Max), False
        Case pbPercent
            View.UpdateTo VBA.CLng(RHS), 100
        End Select
    End With
End Sub

Private Property Get IProgressBar_Self() As IProgressBar
    Set IProgressBar_Self = Me
End Property

'===============================================================================

Private Sub DisposeView()
    On Error Resume Next
    View.Hide
    On Error GoTo 0
    Set View = Nothing
End Sub

Private Function IsValid(ByVal Value As Variant) As Boolean
    If Not VBA.IsNumeric(Value) Then Exit Function
    With This
        Select Case .Unit
        Case pbDecimal
            If Value > 0 And Value <= 1 Then IsValid = True
        Case pbNumeric
            If Value > 0 Then IsValid = True
        Case pbPercent
            If Value > 0 And Value <= 100 Then IsValid = True
        End Select
    End With
End Function

Private Sub View_QueryCancel()
    With This
        If .Cancelable Then
            .Canceled = True
            DisposeView
        End If
    End With
End Sub

Private Sub ShowView()
    If Not View.Visible Then View.Show vbModeless
End Sub

Private Sub Throw()
    VBA.Err.Raise vbObjectError Or 32, "ProgressBar", "Wrong value in current unit context"
End Sub

Private Property Get Self() As IProgressBar
    Set Self = Me
End Property
