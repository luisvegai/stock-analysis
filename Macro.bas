Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Sub MacroCheck()

    Dim testMessage As String

    testMessage = "Hello World!"

    MsgBox (testMessage)

End Sub

Sub DQAnalysis_pre()

    Worksheets("DQ ANalysis").Activate
    
    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    'Create a header row
    
    Cells(3, 1).Value = "Year"
    
    Cells(3, 2).Value = "Total Daily Volume"
    
    Cells(3, 3).Value = "Return"
    

End Sub



Sub Sumrows()

    Worksheets("2018").Activate

    rowStart = 2
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    totalVolume = 0

    For i = rowStart To rowEnd
        'increase totalVolume
        If Cells(i, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(i, 8).Value
        End If

    Next i
    
    Worksheets("DQ Analysis").Activate

    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume

End Sub

Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    
    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    'Create a header row
    
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    

    Worksheets("2018").Activate
    
    rowStart = 2
    'find the number of rows to loop over
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    
    'set initial volume to zero
    totalVolume = 0
    
    Dim startingPrice As Double
    
    Dim endingPrice As Double

    'loop over all the rows
    For i = rowStart To rowEnd
    
        'increase totalVolume
        If Cells(i, 1).Value = "DQ" Then
        
            totalVolume = totalVolume + Cells(i, 8).Value
            
        End If
    
        If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
        
            startingPrice = Cells(i, 6).Value
        
        End If
        
    
        If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
        
            endingPrice = Cells(i, 6).Value
        
        End If
        
    Next i

    Worksheets("DQ Analysis").Activate
    
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = endingPrice / startingPrice - 1


End Sub


Sub AllStocksAnalysis()

    '1 Format the output sheet on the “All Stocks Analysis” worksheet.
    Worksheets("All Stocks Analysis").Activate
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    '2 Initialize an array of all tickers.
    
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    

'3a Initialize variables for the starting price and ending price.

    Dim startingPrice As Single
    
    Dim endingPrice As Single

'3b Activate the data worksheet.

    Worksheets(yearValue).Activate

'3c Find the number of rows to loop over.

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'4 Loop through the tickers.

    For i = 0 To 11

        ticker = tickers(i)
        totalVolume = 0

'5 Loop through rows in the data.

    Worksheets(yearValue).Activate
        For j = 2 To RowCount

'5a Find the total volume for the current ticker.

            If Cells(j, 1).Value = ticker Then
            
                totalVolume = totalVolume + Cells(j, 8).Value
                
            End If

'5b Find the starting price for the current ticker.
    
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
            
                startingPrice = Cells(j, 6).Value
            
            End If
        
'5c Find the ending price for the current ticker.

            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
            
                endingPrice = Cells(j, 6).Value
            
            End If
        
        Next j

'6 Output the data for the current ticker.
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        

    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Font.Color = RGB(0, 0, 255)
    Range("A3:C3").Font.Italic = True
    Range("A3:C3").Font.Underline = True
    
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "$#,###.00"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Change cell color to green
            Cells(i, 3).Interior.Color = vbGreen
            
        ElseIf Cells(i, 3) < 0 Then

            'Change cell color to red
            Cells(i, 3).Interior.Color = vbRed
            
        
        Else

            'Clear the cell color
            Cells(4, 3).Interior.Color = xlNone

        End If

    Next i

End Sub

Sub nestedLoop()

    Worksheets("SkillDrill").Activate
    
    indxStart = 1
    rowEnd = 5
    colEnd = 10
    
    
    For n = indxStart To rowEnd

        For m = indxStart To 10

            Cells(n, m).Value = n + m

        Next m

    Next n


End Sub

Sub nestedCheckerboard()

    Worksheets("SkillDrill2").Activate

    dataRowEnd = 8
    dataColEnd = 8
    
    For i = 1 To dataRowEnd
    
        If (i Mod 2) Then

            For j = 1 To dataColEnd
            
                If (j Mod 2 = False) Then

                    Cells(i, j).Interior.Color = vbBlack

                End If

            Next j
            
        Else

            For j = 1 To dataColEnd
            
                If (j Mod 2) Then

                    Cells(i, j).Interior.Color = vbBlack

                End If

            Next j

        End If
    
    Next i

End Sub


Sub ClearWorksheet()

    Cells.Clear

End Sub

Sub AllStocksAnalysis_challenge()
    yearValue = InputBox("What year would you like to run the analysis on?")

    Worksheets("All Stocks Analysis Challenge").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    Dim tickers(12) As String

    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"


    Dim totalVolume(12) As Single
    Dim startingPrice(12) As Single
    Dim endingPrice(12) As Single

    For k = 0 To 11

        totalVolume(k) = 0
        startingPrice(k) = 0
        endingPrice(k) = 0

    Next k

    Worksheets(yearValue).Activate

    'get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    Worksheets(yearValue).Activate

    'loop over all the rows

    tickerIndex = 0

    For j = 2 To RowCount

       If Cells(j, 1).Value <> tickers(tickerIndex) Then

          tickerIndex = tickerIndex + 1

       End If

       If Cells(j, 1).Value = tickers(tickerIndex) Then

           'increase totalVolume by the value in the current row
           totalVolume(tickerIndex) = totalVolume(tickerIndex) + Cells(j, 8).Value

       End If

       If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then

           startingPrice(tickerIndex) = Cells(j, 6).Value

       End If

       If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then

           endingPrice(tickerIndex) = Cells(j, 6).Value

       End If


    Next j

    Worksheets("All Stocks Analysis Challenge").Activate

    For i = 0 To 11

        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = totalVolume(i)
        Cells(4 + i, 3).Value = endingPrice(i) / startingPrice(i) - 1

    Next i

    'Formatting
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Font.Color = RGB(0, 0, 255)
    Range("A3:C3").Font.Italic = True
    Range("A3:C3").Font.Underline = True
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "$#,##0.00"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit


    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            Cells(i, 3).Interior.Color = vbGreen

        Else

            Cells(i, 3).Interior.Color = vbRed

        End If

    Next i

End Sub

