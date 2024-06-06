Attribute VB_Name = "Module4"

Sub Ticker()

   Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dict As Object
    Dim outputRow As Long
    Dim word As String

   ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Sheets

    ' Create a dictionary object to store unique words
    Set dict = CreateObject("Scripting.Dictionary")

    ' Find the last row with data in column A (change if your words are in a different column)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop through each row in column A
    For i = 1 To lastRow
        ' Get the word from column A
        word = ws.Cells(i, "A").value
        ' Check if the word is already in the dictionary
        If Not dict.Exists(word) Then
            ' Add the word to the dictionary
            dict.Add word, Nothing
        End If
    Next i

    ' Output the unique words to another column (e.g., column D)
    outputRow = 1
    For Each Key In dict.Keys
        ws.Cells(outputRow, "K").value = Key
        outputRow = outputRow + 1
    Next Key

    ' Clean up
    Set dict = Nothing

Next ws

End Sub


Sub Calculate_Quarterly_Change()
    
    Dim ws As Worksheet
    Dim lastRowA As Long, lastRowK As Long
    Dim i As Long, j As Long
    Dim Ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim startRow As Long
    Dim endRow As Long
    
    ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Sheets
    
    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowK = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    
    ' Loop through each ticker in column K
    For j = 2 To lastRowK
        Ticker = ws.Cells(j, 11).value ' Column K is the 11th column
        
        ' Find the first and last occurrence of the ticker in column A
        startRow = 0
        endRow = 0
        For i = 2 To lastRowA
            If ws.Cells(i, 1).value = Ticker Then
                If startRow = 0 Then startRow = i
                endRow = i
            End If
        Next i
        
        ' Get the first opening price and the last closing price
        If startRow > 0 And endRow > 0 Then
            openPrice = ws.Cells(startRow, 3).value
            closePrice = ws.Cells(endRow, 6).value
            
            ' Calculate the price change
            ws.Cells(j, 12).value = closePrice - openPrice
        End If
    
    Next j
    Next ws
End Sub


Sub CalculateQuarterlyPercentageChange()
    
    Dim ws As Worksheet
    Dim lastRowA As Long, lastRowK As Long
    Dim i As Long, j As Long
    Dim Ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim startRow As Long
    Dim endRow As Long
    
   ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Sheets
    
    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowK = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    
    ' Loop through each ticker in column K
    For j = 2 To lastRowK
        Ticker = ws.Cells(j, 11).value ' Column K is the 11th column
        
        ' Find the first and last occurrence of the ticker in column A
        startRow = 0
        endRow = 0
        For i = 2 To lastRowA
            If ws.Cells(i, 1).value = Ticker Then
                If startRow = 0 Then startRow = i
                endRow = i
            End If
        Next i
        
        ' Get the first opening price and the last closing price
        If startRow > 0 And endRow > 0 Then
            openPrice = ws.Cells(startRow, 3).value
            closePrice = ws.Cells(endRow, 6).value
            
            ' Calculate the price change
            ws.Cells(j, 12).value = closePrice - openPrice
            
            ' Calculate the percentage change
                If openPrice <> 0 Then
                    ws.Cells(j, 13).value = ((closePrice - openPrice) / openPrice)
                Else
                    ws.Cells(j, 13).value = 0
                End If
                
                ' Format as percentage
                ws.Cells(j, 13).NumberFormat = "0.00%"
            
                  
            
        End If
    Next j
    
    Next ws
    
End Sub

Sub LastStage2()
    Dim ws As Worksheet
    Dim lastRowA As Long, lastRowK As Long, lastRowM As Long, lastRowN As Long
    Dim i As Long, j As Long
    Dim Ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim startRow As Long
    Dim endRow As Long
    Dim dict As Object
    Dim tickerSum As Double
    Dim maxVal As Double
    Dim minVal As Double
    Dim maxVal2 As Double
    
    ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Find the last row with data in column A and column K
        lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        lastRowK = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
        lastRowM = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row
        lastRowN = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row
        
        ' Create a dictionary object to store sums of each ticker
        Set dict = CreateObject("Scripting.Dictionary")
        
        ' Loop through each row in column A to calculate sums
        For i = 2 To lastRowA
            Ticker = ws.Cells(i, 1).value ' Column A is the 1st column
            
            ' If the ticker is not in the dictionary, add it
            If Not dict.Exists(Ticker) Then
                dict.Add Ticker, 0
            End If
            
            ' Sum the values in column G for the ticker
            dict(Ticker) = dict(Ticker) + ws.Cells(i, 7).value ' Assuming column G has the values to sum
        Next i
        
        ' Output the sums to column N based on the tickers in column K
        For i = 2 To lastRowK
            Ticker = ws.Cells(i, 11).value ' Column K is the 11th column
            If dict.Exists(Ticker) Then
                ws.Cells(i, 14).value = dict(Ticker) ' Column N is the 14th column
            Else
                ws.Cells(i, 14).value = 0
            End If
        Next i
        
        ' Apply conditional formatting to column L
        With ws.Columns("L")
            .FormatConditions.Delete ' Clear existing formatting
            
            ' Format cells > 0 as green
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0"
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = RGB(144, 238, 144) ' Light green
            End With
            
            ' Format cells < 0 as red
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = RGB(255, 182, 193) ' Light red
            End With
        End With
        
        ' Find the highest, lowest values in column M and highest volume in column N
        maxVal = WorksheetFunction.Max(ws.Range("M2:M" & lastRowM))
        minVal = WorksheetFunction.Min(ws.Range("M2:M" & lastRowM))
        maxVal2 = WorksheetFunction.Max(ws.Range("N2:N" & lastRowN))
        
        ' AutoFit column N
        ws.Columns("N").AutoFit
        ws.Columns("T").AutoFit
        
        ' Output the highest value to cell T2
        ws.Cells(2, 20).value = maxVal ' T2 is the 20th column
        
        ' Output the lowest value to cell T3
        ws.Cells(3, 20).value = minVal ' T3 is the 20th column
        
        ' Output the highest total volume to cell T4
        ws.Cells(4, 20).value = maxVal2 ' T4 is the 20th column
        
        ' Format as percentage
        ws.Cells(2, 20).NumberFormat = "0.00%"
        ws.Cells(3, 20).NumberFormat = "0.00%"
        
        
        
        
        ' Loop through each row in column K to find the corresponding tickers
        For i = 2 To lastRowK
            ' Check for max value in column M
            If ws.Cells(i, 13).value > maxVal Then
                maxVal = ws.Cells(i, 13).value
                maxValTicker = ws.Cells(i, 11).value
            End If
            
            ' Check for min value in column M
            If ws.Cells(i, 13).value < minVal Then
                minVal = ws.Cells(i, 13).value
                minValTicker = ws.Cells(i, 11).value
            End If
            
            ' Check for max volume in column N
            If ws.Cells(i, 14).value > maxvol Then
                maxvol = ws.Cells(i, 14).value
                maxVolTicker = ws.Cells(i, 11).value
            End If
        Next i
        
        ' Output the highest value and its ticker to cells S2 and T2
        ws.Cells(2, 19).value = maxValTicker ' S2 is the 19th column
        
        
        ' Output the lowest value and its ticker to cells S3 and T3
        ws.Cells(3, 19).value = minValTicker ' S3 is the 19th column
       
        
        ' Output the highest total volume and its ticker to cells S4 and T4
        ws.Cells(4, 19).value = maxVolTicker ' S4 is the 19th column
        ws.Cells(4, 20).value = maxvol ' T4 is the 20th column
        
        
        ' AutoFit column N
        ws.Columns("N").AutoFit
        ws.Columns("T").AutoFit
        ' Clean up
        Set dict = Nothing
    Next ws
End Sub




