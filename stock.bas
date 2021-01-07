Sub Process_the_Data()


Dim Sheets_Count As Integer
Dim Current_Sheet As Integer

' Declare last row of Data, Current Row, Counter for different ticker names
Dim Last_Row As Long
Dim Current_Row As Long
Dim Ticker_Names_Found As Integer

' Declare where the Ticker types start, end, create range to encompass first to last
Dim First_Ticker_Row As Long
Dim Last_Ticker_Row As Long
Dim Ticker_Range As Range

' Declare variables to set up Yearly Change and Percent Change
Dim Ticker_Open As Double
Dim Ticker_Close As Double
Dim Ticker_Change As Double
Dim Ticker_Percent_Change As Double


' This makes the code run without moving the screen and finishes faster, False
Application.ScreenUpdating = False

Sheets_Count = ActiveWorkbook.Sheets.Count

For Current_Sheet = 1 To Sheets_Count

    Sheets(Current_Sheet).Select
    

'Find the last row of data
Last_Row = Range("A1").End(xlDown).Row

' Set Counter at 0
Ticker_Names_Found = 0


    ' Define Current_Row
    For Current_Row = 2 To Last_Row
        
        ' Create ActiveCell
        Cells(Current_Row, 1).Select
        
        ' Create Loop, ActiveCell does not equal one cell UP Then
        If ActiveCell <> ActiveCell.Offset(-1, 0) Then
            
            'Ticker Counter goes up 1
            Ticker_Names_Found = Ticker_Names_Found + 1
            
            ' Input to Column I ticker name, Uses Ticker_Names_Found to input in correct column cell
            Range("I1").Offset(Ticker_Names_Found, 0) = ActiveCell
                
            ' Save Ticker Open for end of year so can count Yearly Change later
            Ticker_Open = ActiveCell.Offset(0, 2)
            
            ' Save starting row of ticker type so can count Volume later
            First_Ticker_Row = ActiveCell.Row
        
        End If
        
        ' Create Loop, ActiveCell does not equal one cell DOWN Then
        If ActiveCell <> ActiveCell.Offset(1, 0) Then
        
            ' Save ending row of ticker type so can count Volume later
            Last_Ticker_Row = ActiveCell.Row
            
            ' Save Ticker Close for end of year so can count Yearly Change later
            Ticker_Close = ActiveCell.Offset(0, 5)
            
            ' Find Ticker Change
            Ticker_Change = Ticker_Close - Ticker_Open
            
                If Ticker_Open = 0 Then
                Ticker_Percent_Change = 0
                
                ' Find Ticker Percent Change
                Else: Ticker_Percent_Change = Ticker_Change / Ticker_Open
                
                End If
            
            ' Input to Column J Ticker Change, Uses Offset to move down the column
            Range("J1").Offset(Ticker_Names_Found, 0) = Ticker_Change
            
            ' Input to Column K Ticker Percent Change, Uses Offset to move down the column
            Range("K1").Offset(Ticker_Names_Found, 0) = Ticker_Percent_Change
            
            ' Set Ticker Range so we can find Volume. Goes from First to Last Ticker Row in the same type
            Set Ticker_Range = Range(Cells(First_Ticker_Row, 7), Cells(Last_Ticker_Row, 7))
            
            ' Input to Column L the sum of Volume for all the same ticker type
            Range("L1").Offset(Ticker_Names_Found, 0) = Application.WorksheetFunction.Sum(Ticker_Range)
              
        End If
    
    Next Current_Row
    
Next Current_Sheet

Sheets(1).Select

Range("A1").Select

' This makes the code run without moving the screen and finishes faster, True
Application.ScreenUpdating = True


End Sub
