VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
Option Explicit

'===============================================================================

Public Property Get Canceled() As Boolean
End Property

Public Property Get Cancelable() As Boolean
End Property
Public Property Let Cancelable(ByVal RHS As Boolean)
End Property

Public Property Get CancelButtonCaption() As String
End Property
Public Property Let CancelButtonCaption(ByVal RHS As String)
End Property

Public Property Get Caption() As String
End Property
Public Property Let Caption(ByVal Value As String)
End Property

Public Property Get Max() As Long
End Property

Public Property Get NumericMiddleText() As String
End Property
Public Property Let NumericMiddleText(ByVal RHS As String)
End Property

Public Property Get Step() As Long
End Property

Public Sub Update()
End Sub

Public Sub UpdateTo(ByVal RHS As Variant)
End Sub

'===============================================================================

Public Property Get Self() As IProgressBar
End Property
