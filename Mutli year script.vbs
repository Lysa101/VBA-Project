Attribute VB_Name = "Module1"
Sub mutliyear()
Dim ticker As String
 Dim yearlyopen As Double
 Dim yearlyclose As Double
 Dim yearlypricechange As Double
 Dim percentyearlychange As Double
 Dim totalstockvolume As Double
 Dim greatestpercincrease As Double
 Dim greatestpercdecrease As Double
 Dim greatesttotalvolume As Double
 Dim ws As Worksheet
 
 For Each ws In ThisWorkbook.Worksheets
 
     
 'Find out how many rows in the sheet
 
 RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
 
 'consolidate tickers
 yearlyopen = ws.Cells(2, 3)
 totalstockvolume = 0
 
 ws.Cells(1, 9).Value = "ticker"
 ws.Cells(1, 10).Value = "yearlypricechange"
 ws.Cells(1, 11).Value = "percentyearlychange"
 ws.Cells(1, 12).Value = "totalstockvolume"
 
 ws.Cells(4, 15).Value = "greatest per inc"
 ws.Cells(6, 15).Value = "greatest per dec"
 ws.Cells(8, 15).Value = "greatest volume"
 
  
 Dim Thing As Integer
 Thing = 2
 greatestpercincrease = 0
 greatestpercdecrease = 1000
 greatesttotalvolume = 0
 
 For Z = 2 To RowCount - 1
    If ws.Cells(Z, 1) <> ws.Cells(Z + 1, 1) Then
       ws.Cells(Thing, "l").Value = totalstockvolume
   
         
 'change yearly from open price to end price
  
    yearlyclose = ws.Cells(Z, 6).Value
    
    yearlypricechange = yearlyclose - yearlyopen
     If yearlyopen = 0 Then
     percentyearlychange = 0
     Else
     
     percentyearlychange = Round(yearlypricechange / yearlyopen * 100, 2)
    
    
    If percentyearlychange > greatestpercincrease Then
    greatestpercincrease = percentyearlychange
    
    greatestpercticker = ws.Cells(Z, 1)
    
    End If
    
    If percentyearlychange < greatestpercdecrease Then
    greatestpercdecrease = percentyearlychange
    
    greatestpercticker2 = ws.Cells(Z, 1)
    
    End If
    End If
    
    If totalstockvolume > greatesttotalvolume Then
    greatesttotalvolume = totalstockvolume
    maxticker = ws.Cells(Z, 1)
    
    End If
    
      
    yearlyopen = ws.Cells(Z + 1, 3).Value
    
        ws.Cells(Thing, "I").Value = ws.Cells(Z, 1).Value
        ws.Cells(Thing, "J").Value = yearlypricechange
        ws.Cells(Thing, "K").Value = percentyearlychange
        ws.Cells(Thing, "L").Value = totalstockvolume
                       
        Thing = Thing + 1
        totalstockvolume = 0
                 
        If ws.Cells(2, "J").Value >= 0 Then
        ws.Cells(2, "J").Interior.ColorIndex = 4
        
        ElseIf ws.Cells(2, "J").Value <= 0 Then
        ws.Cells(2, "J").Interior.ColorIndex = 3
        
        End If
        
            
    'Percentage change open to ending
     
     Else
     totalstockvolume = totalstockvolume + ws.Cells(Z, 7)
     
               
     End If
     
 Next Z
 
 ws.Cells(4, 16) = greatestpercincrease
 ws.Cells(5, 16) = greatestpercticker
 ws.Cells(6, 16) = greatestpercdecrease
 ws.Cells(7, 16) = greatestpercticker2
 ws.Cells(8, 16).Value = greatesttotalvolume
 ws.Cells(9, 16).Value = maxticker
 
 Next ws
   
'calculate total stock volume
    
  
 End Sub

       
 

