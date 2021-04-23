'Class Module: Extension

Public name As String
Public code As String
Private categories_ As Variant
Public dateOffset As Long


Public Property Let categories(cats As Variant)
    categories_ = cats
End Property

Public Property Get categories() As Variant
    categories = categories_
End Property