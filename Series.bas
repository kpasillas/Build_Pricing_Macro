'Class Module: Series

Public name As String
Public code As String
Private extensions_ As Variant


Public Property Let extensions(exts As Variant)
    extensions_ = exts
End Property

Public Property Get extensions() As Variant
    extensions = extensions_
End Property