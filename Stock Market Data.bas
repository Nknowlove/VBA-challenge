Attribute VB_Name = "Module1"
Sub StockMarket()
 'For remove dunplicate in ticker
 
    ' Inserting Data Names Via Ranges
      Range("I1").Value = "Ticker"
      Range("J1").Value = "Quarterly Change"
      Range("K1").Value = "Percentage Change"
      Range("L1").Value = "Total Stock Volume"
      Range("O2").Value = "Greatest % increase"
      Range("O3").Value = "Greatest % decrease"
      Range("O4").Value = "Greatest total volume"
      Range("P1").Value = "Ticker"
      Range("Q1").Value = "Value"
    
    'Dim lastrow
     Dim lastrow As Long
     lastrow = Cells(Rows.Count, "A").End(xlUp).Row
     
    'Dim Dictionary Key and Item
     Dim dict As Object
     Set dict = CreateObject("Scripting.Dictionary")
     
    'Dim variable
     Dim cell As Range
     
    'For range "A2:A" & lastrow loop, if key "cell" not exist and not zero, then add to dictionary
     For Each cell In Range("A2:A" & lastrow)
         
         If Not dict.Exists(cell.Value) And Not IsEmpty(cell.Value) Then
         dict.Add cell.Value, ""
         
         End If
         
         Next cell
         
     'Put value into Column I, and tranpose from hirizon to virtical
      Range("I2").Resize(dict.Count, 1).Value = WorksheetFunction.Transpose(dict.Keys)
       
 ' For Quarterly Changes & Percentage Changes Calculation
     
     'Dim Dictionary Key and Item
      Dim dictopenprice As Object
      Set dictopenprice = CreateObject("Scripting.Dictionary")
      
      Dim dictcloseprice As Object
      Set dictcloseprice = CreateObject("Scripting.Dictionary")
      
     'Dim variables
      Dim currentTicker As String
      Dim openPrice As Double
      Dim closePrice As Double
      Dim currentDate As Date
      Dim i As Long
    
     'Loop through each row from 2 to the last row
      For i = 2 To lastrow
          currentTicker = Cells(i, 1).Value
            If IsDate(Cells(i, 2).Value) Then
               currentDate = Cells(i, 2).Value
            
            Else
            
              'Format Date
               currentDate = DateSerial(Left(Cells(i, 2).Value, 4), Mid(Cells(i, 2).Value, 5, 2), Right(Cells(i, 2).Value, 2))
            End If
            
      openPrice = Cells(i, 3).Value
      closePrice = Cells(i, 6).Value
            
           'If openPrice not exists in dictionary, then adds and equals to openPrice
            If Not dictopenprice.Exists(currentTicker) Then
               dictopenprice(currentTicker) = openPrice
            End If
            
           'Update closePrice
            dictcloseprice(currentTicker) = closePrice
            
      Next i
      
    'Calculate Quarterly Changes and Percentage Changes
     Dim resultRow As Long
     resultRow = 2
     
    'Dim Variables
     Dim key As Variant
     Dim startOpenPrice As Double
     Dim endClosePrice As Double
     Dim QuarterChange As Double
     Dim PercentageChange As Double
     
          'Loop for key to inital OpenPrice and ClosePrice
           For Each key In dictopenprice.Keys
        
               startOpenPrice = dictopenprice(key)
        
               endClosePrice = dictcloseprice(key)
        
               QuarterChange = endClosePrice - startOpenPrice
        
           If startOpenPrice <> 0 Then
            PercentageChange = QuarterChange / startOpenPrice
           
           Else
           
           'In Case Denominator Is Zero
            PercentageChange = 0
           
           End If
           
                
        'Calculate and formate
      
           Cells(resultRow, 10).Value = QuarterChange
           Cells(resultRow, 11).Value = PercentageChange
           Cells(resultRow, 11).Value = Format(PercentageChange, "0.00%")
           
           resultRow = resultRow + 1
    
    Next key
    
    'Dim variables
     Dim lastii As Long
     lastii = Cells(Rows.Count, "J").End(xlUp).Row

     Dim qc As Double
          'Loop Through to last cell
           For i = 2 To lastii
              qc = Cells(i, 10).Value
              
               'If Qc >0, Mark "Green", Qc < 0, Mark "Red"
                If qc > 0 Then
                   Cells(i, 10).Interior.ColorIndex = 4
 
           ElseIf qc < 0 Then
 
                   Cells(i, 10).Interior.ColorIndex = 3
 
    End If
 
    Next i
    
    'Dim variables
          Dim lastiii As Long
     lastiii = Cells(Rows.Count, "K").End(xlUp).Row

     Dim pc As Double
          'Loop Through to last cell
           For i = 2 To lastiii
              pc = Cells(i, 11).Value
              
               'If pc >0, Mark "Green", pc < 0, Mark "Red"
                If pc > 0 Then
                   Cells(i, 11).Interior.ColorIndex = 4
 
           ElseIf pc < 0 Then
 
                   Cells(i, 11).Interior.ColorIndex = 3
 
    End If
 
    Next i
            
       
 'For Total Stock Volume Calculation
   
   'Dim dictTotalVolume As Object
    Set dictTotalVolume = CreateObject("Scripting.Dictionary")
    
   'Dim variables
    Dim volume As Double
    Dim ticker As String
    
          'Loop through to last row
           For i = 2 To lastrow
               ticker = Cells(i, 1).Value
               volume = Cells(i, 7).Value
        
          'If Volumn Exists, Then Calculate Total Volume
           If ticker <> "" Then
              If dictTotalVolume.Exists(ticker) Then
                 dictTotalVolume(ticker) = dictTotalVolume(ticker) + volume
           
           Else
               'If Volume not exists, Add Volume into Dictionary
                dictTotalVolume(ticker) = volume
            
           End If
        
    End If
    
    Next i
   
          'Put Value into Ranges
           resultRow = 2
    
           For Each key In dictTotalVolume.Keys
               Cells(resultRow, 12).Value = dictTotalVolume(key)
            
           resultRow = resultRow + 1
            
    Next key
      
 'For Greatest Increase, Dercease And Total Volume
    
   'Dim Variables
    Dim lastline As Long
    lastline = Cells(Rows.Count, "k").End(xlUp).Row
    Dim maxv As Double
    Dim minv As Double
  
   'Find MaxValue and MinValue
    maxv = Application.WorksheetFunction.Max(Range("k2:k" & lastline))
    minv = Application.WorksheetFunction.Min(Range("k2:k" & lastline))
    
   'Formate Maxv And Minv
    Range("Q2") = maxv
    Range("Q2") = Format(maxv, "0.00%")

    Range("Q3") = minv
    Range("Q3") = Format(minv, "0.00%")
    
   'Loop Through to Last Cell
           For i = 2 To lastline
    
           Set cell = Cells(i, 11)
          
          'If Cell Value Equal to Maxv, record stock name
           If cell.Value = maxv Then
           Range("P2").Value = cell.Offset(0, -2).Value
    
    End If
    
          'If Cell Value Equal to Minv, record stock name
           If cell.Value = minv Then
           Range("P3").Value = cell.Offset(0, -2).Value
    
    End If
    
    
    Next i
    
   'Dim Variable
   
    
    lastwin = Cells(Rows.Count, "L").End(xlUp).Row
    
    Dim maxtv As String
    
   'Finding Max Total Value
    maxtv = Application.WorksheetFunction.Max(Range("L2:L" & lastwin))
    
    Range("Q4") = maxtv
    
    searchValue = Range("q4")
    
          'Loop Through To Last Cell, If Have Value And Equal to Search Value, Record Stock Name
           For Each cell In Range("l2:l" & lastwin)
               If IsNumeric(cell.Value) Then
               If cell.Value = searchValue Then
                  Range("P4").Value = cell.Offset(0, -3).Value
           Exit For
            
           End If
        
           End If
    
    Next cell
    
    
End Sub
