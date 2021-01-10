
# VBA Homework - The VBA of Wall Street

## Background

You are well on your way to becoming a programmer and Excel master! In this homework assignment you will use VBA scripting to analyze real stock market data. Depending on your comfort level with VBA, you may choose to challenge yourself with a few of the challenge tasks.

### Before You Begin

1. Create a new repository for this project called `VBA-challenge`. **Do not add this homework to an existing repository**.

2. Inside the new repository that you just created, add any VBA files you use for this assignment. These will be the main scripts to run for each analysis.

### Files

* [Test Data](Resources/alphabetical_testing.xlsx) - Use this while developing your scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Run your scripts on this data to generate the final homework report.

### Stock market analyst

![stock Market](Images/stockmarket.jpg)

## Instructions

* Create a script that will loop through all the stocks for one year and output the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* You should also have conditional formatting that will highlight positive change in green and negative change in red.

* The result should look as follows.

![moderate_solution](Images/moderate_solution.png)

## BONUS

* Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". The solution will look as follows:

![hard_solution](Images/hard_solution.png)

* Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

## Other Considerations

* Use the sheet `alphabetical_testing.xlsx` while developing your code. This data set is smaller and will allow you to test faster. Your code should run on this file in less than 3-5 minutes.

* Make sure that the script acts the same on each sheet. The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.

* Some assignments, like this one, contain a bonus. It is possible to achieve mastery on this assignment without completing the bonus. The bonus adds an opportunity to further develop you skills and be rewarded extra points for doing so.

## Submission

* To submit please upload the following to Github:

  * A screen shot for each year of your results on the Multi Year Stock Data.

  * VBA Scripts as separate files.

* Ensure you commit regularly to your repository and it contains a README.md file.

* After everything has been saved, create a sharable link and submit that to <https://bootcampspot-v2.com/>.

- - -

### Copyright

Trilogy Education Services © 2020. All Rights Reserved.

Sub Vba_Hw()

Dim Ticker
Dim YearlyChange
Dim PercentChange
Dim TotalStockVolume As Double
Dim OpenPrice
Dim ClosePrice
Dim SummaryTableRow
Dim YearStart
Dim wsCount As Integer


wsCount = ActiveWorkbook.Worksheets.Count

For ws = 1 To wsCount

    Worksheets(ws).Range("I1") = "Ticker"
    Worksheets(ws).Range("J1") = "Yearly Change"
    Worksheets(ws).Range("K1") = "Percent Change"
    Worksheets(ws).Range("L1") = "Total Stock Volume"
    

    SummaryTableRow = 2

    For i = 2 To Worksheets(ws).Cells(Rows.Count, 1).End(xlUp).Row

        Ticker = Worksheets(ws).Cells(i, 1)
        TotalStockVolume = TotalStockVolume + Cells(i, 7)

        If OpenPrice = "" Then
            OpenPrice = Worksheets(ws).Cells(i, 3)
        End If
   
        If Ticker <> Worksheets(ws).Cells((i + 1), 1) Then
        
 
        ClosePrice = Worksheets(ws).Cells(i, 6)
 
        YearlyChange = OpenPrice - ClosePrice

        Worksheets(ws).Range("I" & SummaryTableRow).Value = Ticker
        Worksheets(ws).Range("J" & SummaryTableRow).Value = YearlyChange
        
        If YearlyChange > 0 Then
        Worksheets(ws).Range("J" & SummaryTableRow).Interior.ColorIndex = 4
        Else
        Worksheets(ws).Range("J" & SummaryTableRow).Interior.ColorIndex = 3
        End If
        
        If startPrice <> ClosePrice Then
            PercentChange = YearlyChange / ClosePrice
        Else
            PercentChange = 0
        End If
        Worksheets(ws).Range("K" & SummaryTableRow).Value = PercentChange
        Worksheets(ws).Range("K" & SummaryTableRow).NumberFormat = "0.00%"
        
       
        Worksheets(ws).Range("L" & SummaryTableRow).Value = TotalStockVolume
        
       
        
        SummaryTableRow = SummaryTableRow + 1
        TotalStockVolume = 0
        
        End If
        
      
    Next i
  

    Worksheets(ws).Cells(2, 15).Value = "Greatest % Increase"
    Worksheets(ws).Cells(3, 15).Value = "Greatest % Decrease"
    Worksheets(ws).Cells(4, 15).Value = "Greatest Total Volume"
    Worksheets(ws).Cells(1, 16).Value = "Ticker"
    Worksheets(ws).Cells(1, 17).Value = "Value"
    
    lastrow = Worksheets(ws).Cells(Rows.Count, 9).End(xlUp).Row
    

    Dim best_stock As String
    Dim best_value As Double
    best_value = Worksheets(ws).Cells(2, 11).Value
    
    Dim worst_stock As String
    Dim worst_value As Double
    worst_value = Worksheets(ws).Cells(2, 11).Value
    
    Dim most_vol_stock As String
    Dim most_vol_value As Double
    most_vol_value = Worksheets(ws).Cells(2, 12).Value
    
    For o = 2 To lastrow
        If Worksheets(ws).Cells(o, 11).Value > best_value Then
        best_value = Worksheets(ws).Cells(o, 11).Value
        best_stock = Worksheets(ws).Cells(o, 9).Value
        End If
        If Worksheets(ws).Cells(o, 11).Value < worst_value Then
        worst_value = Worksheets(ws).Cells(o, 11).Value
        worst_stock = Worksheets(ws).Cells(o, 9).Value
        End If
        If Worksheets(ws).Cells(o, 12).Value > most_vol_value Then
        most_vol_value = Worksheets(ws).Cells(o, 12).Value
        most_vol_stock = Worksheets(ws).Cells(o, 9).Value
        End If
        'Move all data to performance table
        Worksheets(ws).Cells(2, 16).Value = best_stock
        Worksheets(ws).Cells(2, 17).Value = best_value
        Worksheets(ws).Cells(2, 17).NumberFormat = "0.00%"
        Worksheets(ws).Cells(3, 16).Value = worst_stock
        Worksheets(ws).Cells(3, 17).Value = worst_value
        Worksheets(ws).Cells(3, 17).NumberFormat = "0.00%"
        Worksheets(ws).Cells(4, 16).Value = most_vol_stock
        Worksheets(ws).Cells(4, 17).Value = most_vol_value
        Worksheets(ws).Columns("I:L").EntireColumn.AutoFit
        Worksheets(ws).Columns("O:Q").EntireColumn.AutoFit
    Next o
  
  
Next ws


End Sub