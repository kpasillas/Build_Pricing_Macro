'Class Module: Departure

Public code As String
Public startDate As Date
Private originalCurrencyPrices_ As Variant
Public rateBandID As Long
Private extensionCurrencyPrices_ As Scripting.Dictionary


Public Property Let originalCurrencyPrices(Prices As Variant)
    originalCurrencyPrices_ = Prices
End Property

Public Property Get originalCurrencyPrices() As Variant
    originalCurrencyPrices = originalCurrencyPrices_
End Property


Public Property Let extensionCurrencyPrices(Prices As Variant)
    Set extensionCurrencyPrices_ = New Scripting.Dictionary
    
    Dim i As Long
    Dim price As CurrencyPricing
    For i = 0 To (UBound(Prices) - LBound(Prices))
        Set price = Prices(i)
        extensionCurrencyPrices_.Add price.code, price
    Next i
End Property

Public Property Get extensionCurrencyPrices() As Scripting.Dictionary
    Set extensionCurrencyPrices = extensionCurrencyPrices_
End Property



'************************* For Debugging *************************
Public Sub printDebug()

    Debug.Print "Departure.printDebug", "Code: " & code
    Debug.Print , , "Land Only", "Extension"
    Dim i As Long
    For i = 0 To (UBound(originalCurrencyPrices_) - LBound(originalCurrencyPrices_))
        Debug.Print "Currency #: " & i, originalCurrencyPrices_(i).code
        Debug.Print , "Twin:", originalCurrencyPrices_(i).roomTypePrices.twinPrice, extensionCurrencyPrices_(originalCurrencyPrices_(i).code).roomTypePrices.twinPrice
        Debug.Print , "Single:", originalCurrencyPrices_(i).roomTypePrices.singlePrice, extensionCurrencyPrices_(originalCurrencyPrices_(i).code).roomTypePrices.singlePrice
        Debug.Print , "Triple:", originalCurrencyPrices_(i).roomTypePrices.triplePrice, extensionCurrencyPrices_(originalCurrencyPrices_(i).code).roomTypePrices.triplePrice
        Debug.Print , "Child:", originalCurrencyPrices_(i).roomTypePrices.childPrice, extensionCurrencyPrices_(originalCurrencyPrices_(i).code).roomTypePrices.childPrice
    Next i

End Sub
'*****************************************************************