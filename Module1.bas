Option Explicit

Dim departuresStartRow As Long, departuresEndRow As Long
Dim rootPath As String
Dim originalWorksheet As Worksheet, currentWorksheet As Worksheet, seriesDataWorksheet As Worksheet
Dim Series As Series
Dim rateBands As Scripting.Dictionary
Dim rateBandsDict As Scripting.Dictionary, extensionDict As Scripting.Dictionary, categoriesDict As Scripting.Dictionary, columnsDict As Scripting.Dictionary, extensionPricesDict As Scripting.Dictionary


Sub buildPricingMacro()
    
    Application.ScreenUpdating = False
    
    If Application.ActiveWorkbook.name = "Build Pricing Macro.xlsm" Then
    
        MsgBox "Macro must be run from Pricing file. Please try again."
    
    Else
        
        rootPath = GetFolder

        If rootPath = "" Then
            
            MsgBox "Cancelled!"
            
        Else
            
            rootPath = rootPath & "\" & "Pricing Files - " & Format(Now, "dd-mmm-yy hh.mm.ss")
            MkDir rootPath
            Set originalWorksheet = Application.ActiveSheet
            
            Dim ws As Worksheet
            For Each ws In ActiveWorkbook.Worksheets
                
                If ws.name Like "*(PRC)" Then
                
                    Set currentWorksheet = ws
                    buildSeries
                    exportToCSV
                
                End If
                
            Next ws
            
            originalWorksheet.Activate
            MsgBox "Done!"
            
        End If
        
    End If
    
    Application.ScreenUpdating = True

End Sub


Private Sub buildSeries()
    
    getDeparturesStartAndEndRows
    
    currentWorksheet.Activate
    
    Set Series = New Series
    Series.name = UCase(Cells(1, 1).Value)
    Series.code = UCase(Cells(2, 1).Value)
    
    Set seriesDataWorksheet = Workbooks("Build Pricing Macro").Worksheets(Series.code)
    
    buildInfoDicts
    
    Series.extensions = getExtensions
    
    
'********** For Debugging **********
'    Debug.Print "buildSeries()"
'    Dim i As Long, j As Long
'    For i = 0 To (UBound(Series.extensions) - LBound(Series.extensions))
'        Debug.Print "i: " & i, "Name: " & Series.extensions(i).name
'        For j = 0 To (UBound(Series.extensions(i).categories) - LBound(Series.extensions(i).categories))
'            Debug.Print "j: " & j, "Name: " & Series.extensions(i).categories()(j).name
'        Next j
'    Next i
'***********************************


End Sub


Private Sub getDeparturesStartAndEndRows()

    currentWorksheet.Activate
    
    departuresStartRow = 0
    departuresEndRow = 0
    
    Dim i As Long
    For i = 1 To 1000
        If Cells(i, 1).Value = "DEPARTURE CODE" Then
            departuresStartRow = i
        ElseIf (Cells(i, 1).Value = "" And departuresStartRow <> 0) Then
            departuresEndRow = i - 1
            Exit For
        End If
    Next i

End Sub


Private Function getExtensions() As Variant

    currentWorksheet.Activate
    
    Dim extensions() As New Extension
    extensions = getLandOnlyExtensions
    ReDim Preserve extensions(UBound(extensions) + extensionDict.Count)
    
    Dim i As Long
    For i = 0 To (UBound(extensions) - LBound(extensions))
    
        If extensions(i).name = "" Then
            
            extensions(i).name = extensionDict(1)("Name")
            extensions(i).dateOffset = extensionDict(1)("Date Offset")
            
        End If
        
        extensions(i).categories = getCategories(extensions(i).name, extensions(i).code, extensions(i).dateOffset)
    
    Next i


'********** For Debugging **********
'    Debug.Print "getExtensions()"
'    Dim j As Long
'    For i = 0 To (UBound(extensions) - LBound(extensions))
'        Debug.Print "i: " & i, "Name: " & extensions(i).name
'        For j = 0 To (UBound(extensions(i).categories) - LBound(extensions(i).categories))
'            Debug.Print "j: " & j, "Name: " & extensions(i).categories(j).name, "Code: " & extensions(i).categories(j).code
'        Next j
'    Next i
'***********************************
    
    
    getExtensions = extensions

End Function


Private Function getLandOnlyExtensions() As Variant

    currentWorksheet.Activate
    
    Dim extensions() As New Extension
    Dim i As Long
    
    Do While Cells(i + 3, 1).Value <> ""
        ReDim Preserve extensions(i)
        extensions(i).name = "LAND ONLY " & i + 1
        extensions(i).code = UCase(Cells(i + 3, 1).Value)
        extensions(i).dateOffset = 0
        
        i = i + 1
    Loop
    
    getLandOnlyExtensions = extensions

End Function


Private Function getCategories(extensionName As String, extensionCode As String, dateOffset As Long) As Variant

    buildColumnsDict
    buildExtensionPricesDict
    
    currentWorksheet.Activate
    
    Dim categories() As New Category
    Dim departuresWithoutExtensionPricing() As New Departure
    departuresWithoutExtensionPricing = getDeparturesWithoutExtensionCurrencyPrices(dateOffset)
    
    If extensionName Like "LAND ONLY*" Then
        
        ReDim categories(0)
        categories(0).name = "LAND ONLY"
        categories(0).code = extensionCode
        categories(0).departures = departuresWithoutExtensionPricing

    Else
    
        ReDim categories(categoriesDict.Count - 1)
        
        Dim i As Long
        Dim key As Variant
        For Each key In categoriesDict.Keys
            categories(i).name = categoriesDict(key)("Name")
            categories(i).code = categoriesDict(key)("Code")
            categories(i).departures = getDeparturesWithExtensionCurrencyPrices(categoriesDict(key)("Name"), departuresWithoutExtensionPricing)
            
            i = i + 1
        Next key
        
    End If
    
    
'********** For Debugging **********
'    Debug.Print "getCategories(" & extensionName & ", " & extensionCode & ", " & dateOffset & ")"
'    Dim j As Long
'    For j = 0 To (UBound(categories) - LBound(categories))
'        categories(j).printDebug
'    Next j
'
'    For i = 0 To (UBound(categories) - LBound(categories))
'        Debug.Print "i: " & i, "Name: " & categories(i).name
'    Next i
'***********************************
    
    
    getCategories = categories

End Function


Private Function getDeparturesWithoutExtensionCurrencyPrices(dateOffset As Long) As Variant
    
    currentWorksheet.Activate
    
    Dim depArray() As New Departure
    ReDim depArray(departuresEndRow - departuresStartRow - 1)
    
    Dim i As Long, j As Long
    For i = (departuresStartRow + 1) To departuresEndRow
        
        If (Cells(i, columnsDict("BUILD SUPPORT")).Value <> "") And (Not Cells(i, columnsDict("BUILD SUPPORT")).Value Like "No*") Then
        
            ReDim Preserve depArray(UBound(depArray) - 1)
        
'TODO: Figure out how to handle Stampede departures separately, filter out for now.
        
        Else
        
            depArray(j).code = Cells(i, 1).Value
            depArray(j).startDate = Cells(i, 2).Value
            depArray(j).originalCurrencyPrices = getLandOnlyCurrencyPrices(i)
            depArray(j).rateBandID = getRateBand(Cells(i, 2).Value + dateOffset)
            
            If Cells(i, columnsDict("BUILD SUPPORT")).Value Like "No*" Then
                depArray(j).extensionOffered = False
            Else
                depArray(j).extensionOffered = True
            End If
            
            j = j + 1
            
        End If
        
    Next i
    

'********** For Debugging **********
'    Debug.Print "getDeparturesWithoutExtensionCurrencyPrices(" & dateOffset & ")"
'    For i = 0 To (UBound(depArray) - LBound(depArray))
'        Debug.Print "i: " & i, "Departure Code: " & depArray(i).code, "Start Date: " & depArray(i).startDate, "Extension Offered: " & depArray(i).extensionOffered, "Rate Band ID: " & depArray(i).rateBandID
'
'        For j = 0 To (UBound(depArray(i).originalCurrencyPrices) - LBound(depArray(i).originalCurrencyPrices))
'            Debug.Print "j: " & j, "Currency Code: " & depArray(i).originalCurrencyPrices(j).code, "Twin: " & depArray(i).originalCurrencyPrices(j).roomTypePrices.twinPrice, "Single: " & depArray(i).originalCurrencyPrices(j).roomTypePrices.singlePrice, "Triple: " & depArray(i).originalCurrencyPrices(j).roomTypePrices.triplePrice, "Child: " & depArray(i).originalCurrencyPrices(j).roomTypePrices.childPrice
'        Next j
'    Next i
'***********************************

    
    getDeparturesWithoutExtensionCurrencyPrices = depArray

End Function


Private Sub buildInfoDicts()

    seriesDataWorksheet.Activate

    Set rateBandsDict = New Scripting.Dictionary
    Set extensionDict = New Scripting.Dictionary
    Set categoriesDict = New Scripting.Dictionary
    Dim rowDict As Scripting.Dictionary
    
    Dim i As Long, j As Long
    For i = 1 To 100
    
        Select Case Cells(i, 1).Value
        
            Case "Rate Bands"
                j = 1
                Do While Cells(i + j, 2) <> ""
                    Set rowDict = New Scripting.Dictionary
                    
                    rowDict.Add "Start Date", Cells(i + j, 2).Value
                    rowDict.Add "End Date", Cells(i + j, 3).Value
                    rowDict.Add "Rate Band", Cells(i + j, 4).Value
                    
                    rateBandsDict.Add j, rowDict
                    j = j + 1
                Loop
            
            Case "Extension"
                j = 1
                Do While Cells(i + j, 2) <> ""
                    Set rowDict = New Scripting.Dictionary
            
                    rowDict.Add "Name", UCase(Cells(i + j, 2).Value)
                    rowDict.Add "Date Offset", CLng(Cells(i + j, 3).Value)
            
                    extensionDict.Add j, rowDict
                    j = j + 1
                Loop
            
            Case "Categories"
                j = 1
                Do While Cells(i + j, 2) <> ""
                    Set rowDict = New Scripting.Dictionary
                    
                    rowDict.Add "Name", UCase(Cells(i + j, 2).Value)
                    rowDict.Add "Code", UCase(Cells(i + j, 3).Value)
                    
                    categoriesDict.Add j, rowDict
                    j = j + 1
                Loop
        
        End Select
        
    Next i
    

'********** For Debugging **********
'    Debug.Print "buildInfoDicts()"
'    Dim outerKey As Variant, innerKey As Variant
'    For Each outerKey In rateBandsDict.Keys
'        Debug.Print outerKey
'        For Each innerKey In rateBandsDict(outerKey).Keys
'            Debug.Print innerKey, rateBandsDict(outerKey)(innerKey)
'        Next innerKey
'    Next outerKey
'***********************************


End Sub


Private Sub buildColumnsDict()

    currentWorksheet.Activate
    
    Set columnsDict = New Scripting.Dictionary
    
    Dim currencyCodeArray(8) As String
    currencyCodeArray(0) = "AUD"
    currencyCodeArray(1) = "CAD"
    currencyCodeArray(2) = "EUR"
    currencyCodeArray(3) = "GBP"
    currencyCodeArray(4) = "GET"
    currencyCodeArray(5) = "NZD"
    currencyCodeArray(6) = "SAR"
    currencyCodeArray(7) = "SIN"
    currencyCodeArray(8) = "USD"
    
    Dim i As Long, j As Long
    For i = 0 To (UBound(currencyCodeArray) - LBound(currencyCodeArray))
        
        Do While j <= 1000
            j = j + 1
            
            Select Case Cells(departuresStartRow, j).Value
            
                Case "BUILD SUPPORT"
                    columnsDict("BUILD SUPPORT") = j
                
                Case "BUILD " & currencyCodeArray(i)
                    columnsDict.Add "BUILD " & currencyCodeArray(i), j
            
                Case "SINGLE SUPP " & currencyCodeArray(i)
                    columnsDict.Add "SINGLE SUPP " & currencyCodeArray(i), j
                    
                Case "TRIPLE DISC " & currencyCodeArray(i)
                    columnsDict.Add "TRIPLE DISC " & currencyCodeArray(i), j
                    
                Case "YTD " & currencyCodeArray(i)
                    columnsDict.Add "YTD " & currencyCodeArray(i), j
                    Exit Do
            
            End Select
            
        Loop
        j = 0
        
    Next i
    
    columnsDict.Add "SINGLE SUPP GET", columnsDict("SINGLE SUPP USD")
    columnsDict.Add "TRIPLE DISC GET", columnsDict("TRIPLE DISC USD")
    columnsDict.Add "YTD GET", columnsDict("YTD USD")
    columnsDict.Add "SINGLE SUPP SIN", columnsDict("SINGLE SUPP USD")
    columnsDict.Add "TRIPLE DISC SIN", columnsDict("TRIPLE DISC USD")
    columnsDict.Add "YTD SIN", columnsDict("YTD USD")


'********** For Debugging **********
'    Debug.Print "buildColumnsDict()"
'    Dim key As Variant
'    For Each key In columnsDict.Keys
'        Debug.Print key, columnsDict(key)
'    Next key
'***********************************


End Sub


Private Sub buildExtensionPricesDict()

    Dim ws As Worksheet
    For Each ws In currentWorksheet.Parent.Worksheets
        If ws.name Like "Rocky Mountaineer-*" Then
            ws.Activate
            Exit For
        End If
    Next ws
    
    Dim currencyCodeArray() As String
    ReDim currencyCodeArray(10)
    currencyCodeArray(0) = "AUD"
    currencyCodeArray(1) = "CAD"
    currencyCodeArray(2) = "EUR"
    currencyCodeArray(3) = "GBP"
    currencyCodeArray(4) = "NZD"
    currencyCodeArray(5) = "SAR"
    currencyCodeArray(6) = "USD"
    currencyCodeArray(7) = "CODE"
    currencyCodeArray(8) = "CATEGORY"
    currencyCodeArray(9) = "SUPPORT"
    currencyCodeArray(10) = "TYPE"
    
    Dim i As Long, j As Long, k As Long, startRow As Long, startColumn As Long, tripRow As Long
    Dim cat As String
    Set extensionPricesDict = New Scripting.Dictionary
    Dim extensionColumnsDict As New Scripting.Dictionary, rbPricingDict As Scripting.Dictionary
    Dim currencyPricesArray() As New CurrencyPricing
    Dim roomTypePrices As Prices
    
    'Get start row(i)/column(j) by looking for "BRAND"
    For i = 1 To 1000
    
        If startColumn = 0 Then
        
            For j = 1 To 1000
            
                If Cells(i, j).Value = "BRAND" Then
                    startRow = i
                    startColumn = j
                    Exit For
                End If
            
            Next j
        
        ElseIf Cells(i, startColumn + 1).Value = Series.code Then
        
            tripRow = i
            Exit For
            
        End If
    
    Next i
    
    For i = 0 To (UBound(currencyCodeArray) - LBound(currencyCodeArray))
    
        For j = startColumn To 1000
        
            If Cells(startRow, j).Value = currencyCodeArray(i) Then
                extensionColumnsDict.Add currencyCodeArray(i), j
                Exit For
            End If
        
        Next j
    
    Next i
    
    extensionColumnsDict.Add "GET", extensionColumnsDict("USD")
    currencyCodeArray(7) = "GET"
    extensionColumnsDict.Add "SIN", extensionColumnsDict("USD")
    currencyCodeArray(8) = "SIN"
    ReDim Preserve currencyCodeArray(8)

    i = tripRow
    Do While Cells(i, extensionColumnsDict("CODE")).Value = Series.code
    
        Set rbPricingDict = New Scripting.Dictionary
        cat = Cells(i, extensionColumnsDict("CATEGORY")).Value
        
        Do While (Cells(i, extensionColumnsDict("CODE")).Value = Series.code And Cells(i, extensionColumnsDict("CATEGORY")).Value = cat)
        
            ReDim currencyPricesArray(8)
            
            For k = 0 To (UBound(currencyCodeArray) - LBound(currencyCodeArray))
        
                currencyPricesArray(k).code = currencyCodeArray(k)
                Set roomTypePrices = New Prices
                
                For j = 0 To 3
        
                    Select Case Cells(i + j, extensionColumnsDict("TYPE")).Value
                                
                        Case "DOUBLE"
                            roomTypePrices.twinPrice = Cells(i + j, extensionColumnsDict(currencyCodeArray(k))).Value
                        
                        Case "SINGLE"
                            roomTypePrices.singlePrice = Cells(i + j, extensionColumnsDict(currencyCodeArray(k))).Value
                        
                        Case "TRIPLE"
                            roomTypePrices.triplePrice = Abs(Cells(i + j, extensionColumnsDict(currencyCodeArray(k))).Value)
                        
                        Case "CHILD"
                            roomTypePrices.childPrice = Abs(Cells(i + j, extensionColumnsDict(currencyCodeArray(k))).Value)
                    
                    End Select
                
                Next j
                
                Set currencyPricesArray(k).roomTypePrices = roomTypePrices
                
            Next k
            
            rbPricingDict.Add Cells(i, extensionColumnsDict("SUPPORT")).Value, currencyPricesArray
        
            i = i + 5
        
        Loop
        
        extensionPricesDict.Add UCase(Cells(i - 5, extensionColumnsDict("CATEGORY")).Value), rbPricingDict
    
    Loop
    
    
'********** For Debugging **********
'    Debug.Print "buildExtensionPricesDict()"
'    Dim outerKey As Variant, innerKey As Variant
'    For Each outerKey In extensionPricesDict.Keys
'        Debug.Print "Category: " & outerKey
'        For Each innerKey In extensionPricesDict(outerKey).Keys
'            Debug.Print "Rate Band: " & innerKey
'            For i = 0 To (UBound(extensionPricesDict(outerKey)(innerKey)) - LBound(extensionPricesDict(outerKey)(innerKey)))
'                Debug.Print i, extensionPricesDict(outerKey)(innerKey)(i).code, "Twin: " & extensionPricesDict(outerKey)(innerKey)(i).roomTypePrices.twinPrice, "Single: " & extensionPricesDict(outerKey)(innerKey)(i).roomTypePrices.singlePrice, "Triple: " & extensionPricesDict(outerKey)(innerKey)(i).roomTypePrices.triplePrice, "Child: " & extensionPricesDict(outerKey)(innerKey)(i).roomTypePrices.childPrice
'            Next i
'        Next innerKey
'    Next outerKey
'***********************************


End Sub


Private Function getLandOnlyCurrencyPrices(currentRow As Long) As Variant

    Dim pricing(8) As New CurrencyPricing
    Dim roomTypePrice As Prices
    
    Dim currencyCodeArray(8) As String
    currencyCodeArray(0) = "AUD"
    currencyCodeArray(1) = "CAD"
    currencyCodeArray(2) = "EUR"
    currencyCodeArray(3) = "GBP"
    currencyCodeArray(4) = "GET"
    currencyCodeArray(5) = "NZD"
    currencyCodeArray(6) = "SAR"
    currencyCodeArray(7) = "SIN"
    currencyCodeArray(8) = "USD"
    
    Dim i As Long
    For i = 0 To (UBound(currencyCodeArray) - LBound(currencyCodeArray))
        
        Set roomTypePrice = New Prices
        
        pricing(i).code = currencyCodeArray(i)
        roomTypePrice.twinPrice = Cells(currentRow, columnsDict("BUILD " & currencyCodeArray(i))).Value
        roomTypePrice.singlePrice = Cells(currentRow, columnsDict("SINGLE SUPP " & currencyCodeArray(i))).Value
        roomTypePrice.triplePrice = Cells(currentRow, columnsDict("TRIPLE DISC " & currencyCodeArray(i))).Value
        roomTypePrice.childPrice = Cells(currentRow, columnsDict("YTD " & currencyCodeArray(i))).Value
        
        Set pricing(i).roomTypePrices = roomTypePrice
        
    Next i


'********** For Debugging **********
'    Debug.Print "getLandOnlyCurrencyPrices(" & currentRow & ")"
'    For i = 0 To (UBound(pricing) - LBound(pricing))
'        Debug.Print "i: " & i, "Currency Code: " & pricing(i).code, "Twin: " & pricing(i).roomTypePrices.twinPrice, "Single: " & pricing(i).roomTypePrices.singlePrice, "Triple: " & pricing(i).roomTypePrices.triplePrice, "Child: " & pricing(i).roomTypePrices.childPrice
'    Next i
'***********************************
    
    
    getLandOnlyCurrencyPrices = pricing

End Function


Private Function getRateBand(departureDate As Date) As Long

    Dim RateBand As Long
    
    Dim key As Variant
    For Each key In rateBandsDict.Keys
            
        If (departureDate >= rateBandsDict(key)("Start Date")) And (departureDate <= rateBandsDict(key)("End Date")) Then
            RateBand = rateBandsDict(key)("Rate Band")
            Exit For
        End If
            
    Next key
    
    getRateBand = RateBand

End Function


Private Function getDeparturesWithExtensionCurrencyPrices(extensionName As Variant, departures As Variant) As Variant

    Dim depArray() As New Departure
    ReDim depArray(UBound(departures))
    
    Dim i As Long, j As Long
    For i = 0 To (UBound(departures) - LBound(departures))
        
        If departures(i).extensionOffered = True Then
        
            depArray(j).extensionCurrencyPrices = getExtensionCurrencyPrices(extensionName, departures(i).rateBandID)
            depArray(j).code = departures(i).code
            depArray(j).startDate = departures(i).startDate
            depArray(j).extensionOffered = departures(i).extensionOffered
            depArray(j).rateBandID = departures(i).rateBandID
            depArray(j).originalCurrencyPrices = departures(i).originalCurrencyPrices
            
            j = j + 1
            
        Else
        
            ReDim Preserve depArray(UBound(depArray) - 1)
        
        End If
            
    Next i


'********** For Debugging **********
'    Debug.Print "getDeparturesWithExtensionCurrencyPrices(" & extensionName & ", departures)"
'    For i = 0 To (UBound(depArray) - LBound(depArray))
'        Debug.Print "i: " & i, "Code: " & depArray(i).code, "Start Date: " & depArray(i).startDate, "Rate Band: " & depArray(i).rateBandID
'        Dim key As Variant
'        For Each key In depArray(i).extensionCurrencyPrices.Keys
'            Debug.Print "key: " & key, "Code: " & depArray(i).extensionCurrencyPrices()(key).code, "Twin: " & depArray(i).extensionCurrencyPrices()(key).roomTypePrices.twinPrice, "Single: " & depArray(i).extensionCurrencyPrices()(key).roomTypePrices.singlePrice, "Triple: " & depArray(i).extensionCurrencyPrices()(key).roomTypePrices.triplePrice, "Child: " & depArray(i).extensionCurrencyPrices()(key).roomTypePrices.childPrice
'        Next key
'    Next i
'***********************************
    
    
    getDeparturesWithExtensionCurrencyPrices = depArray

End Function


Private Function getExtensionCurrencyPrices(extensionName As Variant, rateBandID As Long) As Variant

    Dim pricing(8) As New CurrencyPricing
    Dim roomTypePrice As Prices
    
    Dim i As Long
    For i = 0 To (UBound(extensionPricesDict(extensionName)(rateBandID)) - LBound(extensionPricesDict(extensionName)(rateBandID)))
        
        Set roomTypePrice = New Prices
        
        roomTypePrice.twinPrice = extensionPricesDict(extensionName)(rateBandID)(i).roomTypePrices.twinPrice
        roomTypePrice.singlePrice = extensionPricesDict(extensionName)(rateBandID)(i).roomTypePrices.singlePrice
        roomTypePrice.triplePrice = extensionPricesDict(extensionName)(rateBandID)(i).roomTypePrices.triplePrice
        roomTypePrice.childPrice = extensionPricesDict(extensionName)(rateBandID)(i).roomTypePrices.childPrice
        
        pricing(i).code = extensionPricesDict(extensionName)(rateBandID)(i).code
        Set pricing(i).roomTypePrices = roomTypePrice
        
    Next i
    
    
'********** For Debugging **********
'    Debug.Print "getExtensionCurrencyPrices(" & extensionName & ", " & rateBandID & ")"
'    For i = 0 To (UBound(pricing) - LBound(pricing))
'        Debug.Print "i: " & i, "Code: " & pricing(i).code, "Twin: " & pricing(i).roomTypePrices.twinPrice, "Single: " & pricing(i).roomTypePrices.singlePrice, "Triple: " & pricing(i).roomTypePrices.triplePrice, "Child: " & pricing(i).roomTypePrices.childPrice
'    Next i
'***********************************

    
    getExtensionCurrencyPrices = pricing

End Function


Private Sub exportToCSV()

    Dim sFilePath As String, tripNameFolderPath As String, duplicateFileExtension As String
    Dim fileNumber As Integer
    Dim startDate As String, code As String, singlePrice As String, twinPrice As String, triplePrice As String, childPrice As String
    Dim productCodeAndSellingOfficeDict As New Scripting.Dictionary
    Dim i As Long, j As Long, k As Long, l As Long
    
    tripNameFolderPath = rootPath & "\" & Series.name
    MkDir tripNameFolderPath
    
    For i = 0 To (UBound(Series.extensions) - LBound(Series.extensions))
        
        For j = 0 To (UBound(Series.extensions(i).categories) - LBound(Series.extensions(i).categories))
        
            Set productCodeAndSellingOfficeDict = getProductCodeAndSellingOffice(Series.extensions(i).categories()(j).code)
            MkDir tripNameFolderPath & "\" & Series.extensions(i).categories()(j).code
            
            For k = 0 To (UBound(Series.extensions(i).categories()(j).departures()(0).originalCurrencyPrices()) - LBound(Series.extensions(i).categories()(j).departures()(0).originalCurrencyPrices()))
                
                'Debug.Print "i: " & i, "j: " & j, "Category Code: " & series.extensions(i).categories()(j).code, "Currency Code: " & series.extensions(i).categories()(j).departures()(0).originalCurrencyPrices()(k).code
                sFilePath = tripNameFolderPath & "\" & Series.extensions(i).categories()(j).code & "\" & Series.extensions(i).categories()(j).departures()(0).originalCurrencyPrices()(k).code & ".csv"
                fileNumber = FreeFile
                Open sFilePath For Output As #fileNumber
        
                Write #fileNumber, "Product Code", productCodeAndSellingOfficeDict(Series.extensions(i).categories()(j).departures()(0).originalCurrencyPrices()(k).code)("Product Code"), "Selling Company", productCodeAndSellingOfficeDict(Series.extensions(i).categories()(j).departures()(0).originalCurrencyPrices()(k).code)("Selling Company"), "Default Room Type", "Twin"
                Write #fileNumber,
                Write #fileNumber, "Teenager discount (absolute)", "0", "Food Fund", "0", "Port Taxes-Adult", "0", "Port Taxes-Child", "0"
                Write #fileNumber,
                Write #fileNumber,
                Write #fileNumber, "Start Date", "Season Name", "Single(S)", "Twin", "Triple(R)", "Quad(R)", "Child(R)"
            
                For l = 0 To (UBound(Series.extensions(i).categories()(j).departures()) - LBound(Series.extensions(i).categories()(j).departures()))
            
                    startDate = Format(Series.extensions(i).categories()(j).departures()(l).startDate, "dd-mmm-yy")
                    code = Series.extensions(i).categories()(j).departures()(l).code
                    twinPrice = Series.extensions(i).categories()(j).departures()(l).originalCurrencyPrices()(k).roomTypePrices.twinPrice
                    
                    If Series.extensions(i).name Like "LAND ONLY*" Then
                    
                        singlePrice = Series.extensions(i).categories()(j).departures()(l).originalCurrencyPrices()(k).roomTypePrices.singlePrice
                        triplePrice = Series.extensions(i).categories()(j).departures()(l).originalCurrencyPrices()(k).roomTypePrices.triplePrice
                        childPrice = Series.extensions(i).categories()(j).departures()(l).originalCurrencyPrices()(k).roomTypePrices.childPrice
                    
                    Else
                        
                        singlePrice = Series.extensions(i).categories()(j).departures()(l).originalCurrencyPrices()(k).roomTypePrices.singlePrice + Series.extensions(i).categories()(j).departures()(l).extensionCurrencyPrices()(Series.extensions(i).categories()(j).departures()(0).originalCurrencyPrices()(k).code).roomTypePrices.singlePrice
                        triplePrice = Series.extensions(i).categories()(j).departures()(l).originalCurrencyPrices()(k).roomTypePrices.triplePrice + Series.extensions(i).categories()(j).departures()(l).extensionCurrencyPrices()(Series.extensions(i).categories()(j).departures()(0).originalCurrencyPrices()(k).code).roomTypePrices.triplePrice
                        childPrice = Series.extensions(i).categories()(j).departures()(l).originalCurrencyPrices()(k).roomTypePrices.childPrice + Series.extensions(i).categories()(j).departures()(l).extensionCurrencyPrices()(Series.extensions(i).categories()(j).departures()(0).originalCurrencyPrices()(k).code).roomTypePrices.childPrice
                    
                    End If
                    
                    Write #fileNumber, startDate, code, singlePrice, twinPrice, triplePrice, "NA", childPrice
                    'Debug.Print "k: " & k, "l: " & l, startDate, code, singlePrice, twinPrice, triplePrice, "NA", childPrice
                    
                Next l
                
                Close #fileNumber
                    
            Next k
            
        Next j
        
    Next i

End Sub


Private Function GetFolder() As String

    'https://stackoverflow.com/questions/26392482/vba-excel-to-prompt-user-response-to-select-folder-and-return-the-path-as-string/26392703

    Dim fldr As FileDialog
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then
            Set fldr = Nothing
            Exit Function
        Else
            GetFolder = .SelectedItems(1)
            Set fldr = Nothing
        End If
    End With

End Function


Private Function getProductCodeAndSellingOffice(opCode As String) As Scripting.Dictionary

    Dim ws As Worksheet
    For Each ws In seriesDataWorksheet.Parent.Worksheets
        If ws.name Like "*Codes" Then
            ws.Activate
            Exit For
        End If
    Next ws
    
    Dim productCodeAndSellingOfficeDict As New Scripting.Dictionary
    Dim codesDict As Scripting.Dictionary
    
    Dim i As Long, j As Long
    For i = 1 To 1000
        
        If Cells(i, 3).Value = opCode Then
        
            j = 0
            Do While Cells(i + j, 3).Value = opCode
            
                Set codesDict = New Scripting.Dictionary
                
                Select Case True
                
                    Case Cells(i + j, 5).Value Like "GEUSAS"
                        codesDict.Add "Product Code", Cells(i + j, 6).Value
                        codesDict.Add "Selling Company", Cells(i + j, 5).Value
                        productCodeAndSellingOfficeDict.Add "GET", codesDict
                    
                    Case Cells(i + j, 5).Value Like "*SYDS"
                        codesDict.Add "Product Code", Cells(i + j, 6).Value
                        codesDict.Add "Selling Company", Cells(i + j, 5).Value
                        productCodeAndSellingOfficeDict.Add "AUD", codesDict
                    
                    Case Cells(i + j, 5).Value Like "*AKLS"
                        codesDict.Add "Product Code", Cells(i + j, 6).Value
                        codesDict.Add "Selling Company", Cells(i + j, 5).Value
                        productCodeAndSellingOfficeDict.Add "NZD", codesDict
                        
                    Case Cells(i + j, 5).Value Like "*USAS"
                        codesDict.Add "Product Code", Cells(i + j, 6).Value
                        codesDict.Add "Selling Company", Cells(i + j, 5).Value
                        productCodeAndSellingOfficeDict.Add "USD", codesDict
                    
                    Case Cells(i + j, 5).Value Like "*CANS"
                        codesDict.Add "Product Code", Cells(i + j, 6).Value
                        codesDict.Add "Selling Company", Cells(i + j, 5).Value
                        productCodeAndSellingOfficeDict.Add "CAD", codesDict
                        
                    Case Cells(i + j, 5).Value Like "*JBGS"
                        codesDict.Add "Product Code", Cells(i + j, 6).Value
                        codesDict.Add "Selling Company", Cells(i + j, 5).Value
                        productCodeAndSellingOfficeDict.Add "SAR", codesDict
                        
                    Case Cells(i + j, 5).Value Like "*UKLS"
                        codesDict.Add "Product Code", Cells(i + j, 6).Value
                        codesDict.Add "Selling Company", Cells(i + j, 5).Value
                        productCodeAndSellingOfficeDict.Add "GBP", codesDict
                        
                    Case Cells(i + j, 5).Value Like "*SINS"
                        codesDict.Add "Product Code", Cells(i + j, 6).Value
                        codesDict.Add "Selling Company", Cells(i + j, 5).Value
                        productCodeAndSellingOfficeDict.Add "SIN", codesDict
                    
                    Case Cells(i + j, 5).Value Like "*EUOS"
                        codesDict.Add "Product Code", Cells(i + j, 6).Value
                        codesDict.Add "Selling Company", Cells(i + j, 5).Value
                        productCodeAndSellingOfficeDict.Add "EUR", codesDict
                
                End Select
            
                j = j + 1
            Loop
            
            Exit For
        
        End If
        
    Next i
    
    
'********** For Debugging **********
'    Debug.Print "getProductCodeAndSellingOffice(" & opCode & ")"
'    Dim debugKey As Variant
'    For Each debugKey In productCodeAndSellingOfficeDict.Keys
'        Debug.Print "Key: " & debugKey, "Product Code: " & productCodeAndSellingOfficeDict(debugKey)("Product Code"), "Selling Company: " & productCodeAndSellingOfficeDict(debugKey)("Selling Company")
'    Next debugKey
'***********************************

    
    Set getProductCodeAndSellingOffice = productCodeAndSellingOfficeDict

End Function