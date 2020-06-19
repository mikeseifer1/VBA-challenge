Attribute VB_Name = "Module2"
Sub Stocks3()
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
Dim i As Long
Dim j As Integer
Dim numrow As Long
Dim column As Integer
Dim outputrow As Integer
Dim openprice As Double
Dim closeprice As Double
Dim volume As Double
Dim pctchng As Double
Dim numerator As Double
Dim maxpct As Double
Dim mymax As Double
Dim mymin As Double
Dim maxvol As Double
Dim maxcell As Range
Dim maxticker As String
Dim strAddress As String
Dim maxvolticker As String


column = 1
outputrow = 2
volume = 0

Cells(1, 9) = "Ticker"
Cells(1, 15) = "Ticker"
Cells(1, 16) = "Value"
Cells(1, 10) = "Yearly change in Price"
Cells(1, 11) = "Yearly Percentage Change in Price"
Cells(1, 12) = "Total Volume"
Cells(2, 14) = "Greatest % Increase"
Cells(3, 14) = "Greatest % Decrease"
Cells(4, 14) = "Grestest Volume"
Range("I:P").EntireColumn.AutoFit


'Establish total number of rows

numrow = Cells(Rows.Count, 1).End(xlUp).Row

'Start loop

For i = 1 To numrow



    If Cells(i + 1, column).Value <> Cells(i, column).Value Then
    Cells(outputrow, column + 8).Value = Cells(i + 1, column).Value
        'Volume
    
            'if i = 1 then this will skip to end and establish the first open price
            If i <> 1 Then
            Cells(outputrow - 1, column + 11).Value = volume
            volume = 0
        
            closeprice = Cells(i, column + 5).Value
            numerator = closeprice - openprice
            Cells(outputrow - 1, column + 9) = numerator
        
            'Calculate Percentage change
            If numerator = openprice Then
            pctchng = 0
            ElseIf (closeprice <> 0 And openprice = 0) Then
            pctchng = 1
            Else: pctchng = numerator / openprice
            End If
            
                    'format cells either Green or Red
            If pctchng >= 0 Then
            Cells(outputrow - 1, column + 10).Interior.Color = vbGreen
            Else: Cells(outputrow - 1, column + 10).Interior.Color = vbRed
            End If
        
       
            Cells(outputrow - 1, column + 10) = pctchng
            Cells(outputrow - 1, column + 10).NumberFormat = "0.00%"
        
            End If
   
        openprice = Cells(i + 1, column + 2).Value
        outputrow = outputrow + 1
    


    End If
    'Record Sum of volume
volume = volume + CDbl(Cells(i + 1, column + 6).Value)

Next i


    
mymax = Cells(2, 11).Value
maxticker = Cells(2, 9).Value
mymin = Cells(2, 11).Value
minticker = Cells(2, 9).Value
maxvol = Cells(2, 12).Value
maxvolticker = Cells(2, 9).Value

For j = 1 To outputrow
    
    If Cells(j + 2, 11).Value > mymax Then
    mymax = Cells(j + 2, 11).Value
    maxticker = Cells(j + 2, 9).Value
    Else
    mymax = mymax
    maxticker = maxticker
    End If
    
    If Cells(j + 2, 11).Value < mymin Then
    mymin = Cells(j + 2, 11).Value
    minticker = Cells(j + 2, 9).Value
    Else
    mymin = mymin
    minticker = minticker
    End If
    
    If Cells(j + 2, 12).Value > maxvol Then
    maxvol = Cells(j + 2, 12).Value
    maxvolticker = Cells(j + 2, 9).Value
    Else
    maxvol = maxvol
    maxvolticker = maxvolticker
    End If


Next j

Cells(2, 16).Value = mymax
Cells(2, 16).NumberFormat = "0.00%"
Cells(2, 15).Value = maxticker
Cells(3, 16).Value = mymin
Cells(3, 16).NumberFormat = "0.00%"
Cells(3, 15).Value = minticker
Cells(4, 15).Value = maxvolticker
Cells(4, 16).Value = maxvol



ws.Cells(1, 1) = "ticker" 'this sets cell A1 of each sheet to "ticker"
Next

starting_ws.Activate 'activate the worksheet that was originally active


End Sub
