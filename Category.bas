'Class Module: Category

Public name As String
Public code As String
Private departures_ As Variant


Public Property Let departures(deps As Variant)
    ReDim departures_(LBound(deps) To UBound(deps))
    departures_ = deps
End Property

Public Property Get departures() As Variant
    departures = departures_
End Property



'************************* For Debugging *************************
Public Sub printDebug()

    Debug.Print "Category.printDebug", "Name: " & name
    Dim i As Long
    For i = 0 To (UBound(departures_) - LBound(departures_))
        Debug.Print "Departure #: " & i
        Debug.Print departures_(i).printDebug
    Next i

End Sub
'*****************************************************************