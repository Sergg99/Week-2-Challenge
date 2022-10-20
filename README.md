# Week-2-Challenge (1st attempt)

## Overview of Project:

### Purpose:

The purpose of this week's challenge is to refactor the code provided by our friend Steve on VBA Excel. We are helping Steve analyze stocks to offer his parents well-analyzed options in the stocks they should invest in relating to green energy, as well as reduce the time it takes for the code to run. The code analyzes the stocks of two different years, it oversights the volume and the percent return, as well as provides visual representations of the stocks analysis perform. 

## Results:

In “Picture 1”, we can observe the results from the original code for “2018”, it runs in 0.234375 seconds. As seen in “Picture 2”, we can see that our run-time has been increased by 0.0546875 seconds for the same year (2018). The "2017" also seemed to take longer than the original code (see Image 3 & 4 attached below). The factored code has seemed to increase the code run-time slightly, let’s take a look at the refactored code:

   ### See Image 1 (Original code) attached below:
   https://github.com/Sergg99/Week-2-Challenge/blob/44ea40a1dd71e5440378e37cff341a1744a944c7/Challenge%202/Resources/Original%20Run-times/2018%20Run-time%20Original.jpg
   
   ![Image 1](https://github.com/Sergg99/Week-2-Challenge/blob/44ea40a1dd71e5440378e37cff341a1744a944c7/Challenge%202/Resources/Original%20Run-times/2018%20Run-time%20Original.jpg)
   
   ### See Image 2 (Refactored code) attached below:
   https://github.com/Sergg99/Week-2-Challenge/blob/44ea40a1dd71e5440378e37cff341a1744a944c7/Challenge%202/Resources/Results%20Run-time/2018%20Run-time%20Result.jpg
   
   ![Image 2](https://github.com/Sergg99/Week-2-Challenge/blob/44ea40a1dd71e5440378e37cff341a1744a944c7/Challenge%202/Resources/Results%20Run-time/2018%20Run-time%20Result.jpg)
   
   ### See Image 3 (Original code) attached below:
   https://github.com/Sergg99/Week-2-Challenge/blob/44ea40a1dd71e5440378e37cff341a1744a944c7/Challenge%202/Resources/Original%20Run-times/2017%20Run-time%20Original.jpg
   
   ![Image 3](https://github.com/Sergg99/Week-2-Challenge/blob/44ea40a1dd71e5440378e37cff341a1744a944c7/Challenge%202/Resources/Original%20Run-times/2017%20Run-time%20Original.jpg)
   
   ### See Image 4 (Refactored code) attached below:
   https://github.com/Sergg99/Week-2-Challenge/blob/44ea40a1dd71e5440378e37cff341a1744a944c7/Challenge%202/Resources/Results%20Run-time/2017%20Run-time%20Result.jpg
   
   ![Image 4](https://github.com/Sergg99/Week-2-Challenge/blob/44ea40a1dd71e5440378e37cff341a1744a944c7/Challenge%202/Resources/Results%20Run-time/2017%20Run-time%20Result.jpg)
        
   #### Refactored Code:
   
   We have compared stocks, performance, volumes, and more using the data Steve has provided, but in this case, we have refactored to see if we can improve it!
      
  
   

                                                Sub AllStocksAnalysisRefactored()


                                                Dim startTime As Single
                                                Dim endTime  As Single

                                                yearValue = InputBox("What year would you like to run the analysis                       on?")

                                                startTime = Timer

                                                'Format the output sheet on All Stocks Analysis worksheet
                                                Worksheets("All Stocks Analysis").Activate

                                                Range("A1").Value = "All Stocks (" + yearValue + ")"

                                                'Create a header row
                                                Cells(3, 1).Value = "Ticker"
                                                Cells(3, 2).Value = "Total Daily Volume"
                                                Cells(3, 3).Value = "Return"

                                                'Initialize array of all tickers
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

                                                'Activate data worksheet
                                                Worksheets(yearValue).Activate

                                                'Get the number of rows to loop over
                                                RowCount = Cells(Rows.Count, "A").End(xlUp).Row

                                                RowStart = 2

                                                '1a) Create a ticker Index



                                                    tickerIndex = 0


                                                '1b) Create three output arrays

                                                    Dim tickerVolumes(12) As Long

                                                    Dim tickerStartingPrices(12) As Single

                                                    Dim tickerEndingPrices(12) As Single


                                                ''2a) Create a for loop to initialize the tickerVolumes to zero.            

                                                    For i = 0 To 11
                                                        tickerVolumes(i) = 0
                                                    Next i



                                                ''2b) Loop over all the rows in the spreadsheet.



                                                        For i = 2 To RowCount


                                                            '3a) Increase volume for current ticker

                                                          tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value


                                                    '3b) Check if the current row is the first row with the selected tickerIndex.

                                                            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then

                                                            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

                                                            'End If
                                                            End If





                                                    '3c) check if the current row is the last row with the selected ticker
                                                     'If the next rows ticker doesnt match, increase the tickerIndex.
                                                    'If  Then

                                                        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then

                                                                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value



                                                        '3d Increase the tickerIndex.

                                                        tickerIndex = tickerIndex + 1




                                                    'End If

                                                        End If

                                                Next i

                                                '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
                                                For i = 0 To 11

                                                    Sheets("All Stocks Analysis").Activate
                                                    Cells(4 + i, 1).Value = tickers(i)
                                                    Cells(4 + i, 2).Value = tickerVolumes(i)
                                                    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

                                                Next i



                                                'Formatting
                                                Worksheets("All Stocks Analysis").Activate
                                                Range("A3:C3").Font.FontStyle = "Bold"
                                                Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
                                                Range("B4:B15").NumberFormat = "#,##0"
                                                Range("C4:C15").NumberFormat = "0.0%"
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

                                                endTime = Timer
                                                MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

                                                End Sub

   

## Summary: 
  
 ### Deliverable and detail analysis: 
 
 We have reached to a point where I tought that would be more effective to run the analysis for Steve. Our findings suggest that the changes we have performed, result in the addition of run-time, which indicates our code requires more computing power to deliverate. 
 
  ### Disadvantages encountered: 
  
  - We ended right where we started, except now it takes 0.0546875 seconds more than before. In this case, it is not much difference but it is still adding more computing time to run.
  - We came across multiple errors along the way, in the end, with the help of my peers and tutors, I was able to get to our finish line. Not the way we anticipated, but it got us the same results with a different approach!
  - This was not easy, trying to break down lines into simpler code can result in modification of results by mistake. Luckily I was able to find my mistakes before submitting it. This is always a very important step.

    
   ### Advantages encountered: 
   
 - After this project, I have a better understanding of VBA's structure and rules. 
 - Troubleshooting is not always easy and quick as we would like, but it gives you challanging opportuinities to get a better understanding for future encounters. You become knowlageable once you are able to figure it out. This is how you learn to code!
 -  This project has helped me have a better understanding on real-life applications for VBA, and how it can help me in my profetional life. 



### Advantages of refactoring the code: 

I think of refactoring code as a way to clean it up and organize it in a pleasently way, it is also an oopportunity to go through the code and understand what it's doing and how it's doing it. Maybe you can find a simpler and quicker way of saving time and/or computing power. 

### Disavantages of refactoring the code: 

In this case I was not able to save computing power. I even came across 10+ errors along the way, I got frustrated, disencouraged, angry, and more. It is important to ask for help if you can, if not, take a break and clear your mind (it always helps to come back to it if you're stuck). With that said, It is not easy and it can be time consuming. The important thing is to remember it can always be done differently and/or easier. 

### See Table 1 Below (Stocks: 2018): 
https://github.com/Sergg99/Week-2-Challenge/blob/44ea40a1dd71e5440378e37cff341a1744a944c7/Challenge%202/Resources/Results%20Run-time/Table%20Results%20for%20All%20Stocks%202018.jpg

![Image 5](https://github.com/Sergg99/Week-2-Challenge/blob/44ea40a1dd71e5440378e37cff341a1744a944c7/Challenge%202/Resources/Results%20Run-time/Table%20Results%20for%20All%20Stocks%202018.jpg)

### See Table 1 Below (Stocks: 2017): 
https://github.com/Sergg99/Week-2-Challenge/blob/44ea40a1dd71e5440378e37cff341a1744a944c7/Challenge%202/Resources/Results%20Run-time/Table%20Results%20for%20All%20Stocks%202017.jpg

![Image 6](https://github.com/Sergg99/Week-2-Challenge/blob/44ea40a1dd71e5440378e37cff341a1744a944c7/Challenge%202/Resources/Results%20Run-time/Table%20Results%20for%20All%20Stocks%202017.jpg)
