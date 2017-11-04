' 1) prints caclulations in same worksheet
' 2) creates new worksheet
' 3) Copy calculation and pased on new worksheet

Sub stock_analysis():

'Format Cells calculation, 1ft row. FontBold and Underline
Range("I1").Value = "Ticker"
Range("J1").Value = "Total Change"
Range("K1").Value = "% of Change"
Range("L1").Value = "Avg. Daily Change"
Range("M1").Value = "Total vol."
Range("I1:M1").Font.Bold = True
Range("I1:M1").Font.Underline = True


Dim ticker As String
Dim totalchange As Double
Dim i As Integer
'Dim change As Double
Dim j As Integer
Dim start As Double
Dim rowCount As Double
Dim percentChange As Double
'Dim days As Integer
Dim averagedailyChange As Double
Dim totalvol As Double

' Set initial values
j = 0
totalchange = 0
percentChange = 0
start = 2
averagedailyChange = 0
totalvol = 0

' get the row number of the last row with data
rowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

' If ticker changes then print results
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
         ' Stores results in variables
         ticker = Cells(i, 1).Value
         totalchange = totalchange + (Cells(i, 3) - Cells(i, 6))
         percentChange = Round(percentChange + ((Cells(i, 6) / Cells(i, 3)) * 100), 2)
         averagedailChange = (averagedailyChange + (Cells(i, 4) - Cells(i, 5))) / ((i - start) + 1)
         totalvol = totalvol + Cells(i, 7).Value
         
        ' start of the next stock ticker
            start = i + 1
            
        'print results in same worksheet
        Range("I" & 2 + j).Value = ticker
        Range("J" & 2 + j).Value = totalchange
        Range("K" & 2 + j).Value = "%" & percentChange
        Range("L" & 2 + j).Value = averagedailChange
        Range("M" & 2 + j).Value = totalvol
        
       ' colors positives green and negatives red
        Select Case Change
        Case Is > 0
            Range("J" & 2 + j).Interior.ColorIndex = 4
            Case Is < 0
                Range("J" & 2 + j).Interior.ColorIndex = 3
            Case Else
                Range("J" & 2 + j).Interior.ColorIndex = 0
        End Select
        
        ' reset variables for new stock ticker
        j = j + 1
        totalchange = 0
        percentChange = 0
        averagedailyChange = 0
        totalvol = 0
        
' If ticker is still the same add results
    Else
    ticker = Cells(i, 1).Value
    totalchange = totalchange + (Cells(i, 3) - Cells(i, 6))
    percentChange = Round(percentChange + ((Cells(i, 6) / Cells(i, 3)) * 100), 2)
    averagedailChange = (averagedailyChange + (Cells(i, 4) - Cells(i, 5))) / ((i - start) + 1)
    totalvol = totalvol + Cells(i, 7).Value
        
     End If
Next i

'Add New sheet and name as last worksheet
'it will ask to enter the name for new sheet

Dim NewName As String
'get new name
NewName = InputBox("Nwe Name of WorkSheet")
If NewName <> "" Then
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = NewName
End If

'Copy & past results to new sheet
'Get Last row for calculated results
rowCount2 = Cells(Rows.Count, "I").End(xlUp).Row
'Range Copy - past
Worksheets("Stock_data_2016").Range("I1:M1" & rowCount2).Copy Worksheets("NewName").Range("A1:E1" & rowCount2)

End Sub

