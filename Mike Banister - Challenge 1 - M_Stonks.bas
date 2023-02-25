Attribute VB_Name = "M_Stonks"
Sub Stonks()

'Worksheet References
    Dim WB As Workbook:             Set WB = Application.ActiveWorkbook
    Dim WS As Worksheet:            Set WS = WB.ActiveSheet
    
'Make sure data is sorted
    WS.UsedRange.Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlYes
    WS.UsedRange.Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
   
'Variables for Loop
    Dim First_Row As Long:          First_Row = 2
    Dim Last_Row As Long:           Last_Row = WS.UsedRange.End(xlDown).Row

    Dim Ticker As String
    Dim Ticker_Previous As String
    Dim Ticker_Next As String
    
    Dim Opening As Double
    Dim Closing As Variant
    Dim Volume As Double
    Dim Volume_Total As Double

'Variables for Max / Min
    Dim Max_Change As Double:       Max_Change = 0
    Dim Min_Change As Double:       Min_Change = 0
    Dim Max_Volume As Double:       Max_Volume = 0
    
    Dim Max_Change_Ticker As String
    Dim Min_Change_Ticker As String
    Dim Max_Volume_Ticker As String

'Variables for Summary Loop
    Dim Summary_Last_Row As Long

'Variables for Output Table
    Dim Summary_Col As Integer:     Summary_Col = 9
    Dim Summary_Row As Integer:     Summary_Row = 2

'Summary Headers & Formats Because I Can't Help Myself
    WS.Range("I1").Value = "Ticker"
    WS.Range("J1").Value = "Yearly Change"
    WS.Range("K1").Value = "Percent Change"
    WS.Range("L1").Value = "Total Volume"

    With WS.Range("A1:G1, I1:L1, N1:P1")
        .Interior.Color = RGB(54, 96, 146)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True

    End With
    
    With WS.Range("I1:L1, N1:P1")
        .ColumnWidth = 18
    End With
    
    WS.Range("N:N").ColumnWidth = 36
    WS.Range("I:I").HorizontalAlignment = xlLeft
    WS.Range("J:L").HorizontalAlignment = xlRight

'Main Loop
    For a = First_Row To Last_Row
    
    'Row Variables
        Ticker = WS.Cells(a, 1)
        Ticker_Previous = WS.Cells(a - 1, 1)
        Ticker_Next = WS.Cells(a + 1, 1)
        Volume = WS.Cells(a, 7)

    'First Row of Block
        If Ticker <> Ticker_Previous And Ticker = Ticker_Next Then
        
            Opening = WS.Cells(a, 3)
            Volume_Total = Volume_Total + Volume
            Cells(Summary_Row, Summary_Col).Value = Ticker
        
        End If
        
    'Mid Block
        If Ticker = Ticker_Previous And Ticker = Ticker_Next Then
            Volume_Total = Volume_Total + Volume
        End If
        
    'Last Row of Block
        If Ticker = Ticker_Previous And Ticker <> Ticker_Next Then
        
            Closing = WS.Cells(a, 6)
            Volume_Total = Volume_Total + Volume
            Cells(Summary_Row, Summary_Col + 1).Value = Closing - Opening
            Cells(Summary_Row, Summary_Col + 2).Value = (Closing - Opening) / Opening
            Cells(Summary_Row, Summary_Col + 3).Value = Volume_Total

        'Increment Row
            Summary_Row = Summary_Row + 1
            
        'Reset Volume
            Volume_Total = 0
        
        End If

    Next a

'Format Columns
    With WS.Range("J2:J" & Summary_Row)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        .FormatConditions(1).Interior.Color = RGB(255, 100, 100)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        .FormatConditions(2).Interior.Color = RGB(100, 255, 100)
    End With
    
    With WS.Range("K2:K" & Summary_Row)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        .FormatConditions(1).Interior.Color = RGB(255, 100, 100)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        .FormatConditions(2).Interior.Color = RGB(100, 255, 100)
    End With

    WS.Range("J2:J" & Summary_Row).NumberFormat = "#,##0.00;(#,##0.00)"
    WS.Range("K2:K" & Summary_Row).NumberFormat = "0.00%;(0.00%)"
    WS.Range("L2:L" & Summary_Row).NumberFormat = "#,##0;(#,##0)"
    
'Max / Min Loop

    Summary_Last_Row = WS.Range("J1").End(xlDown).Row
      
    For a = First_Row To Summary_Last_Row
    
        If WS.Cells(a, 11) > Max_Change Then
            Max_Change = WS.Cells(a, 11)
            Max_Change_Ticker = WS.Cells(a, 9)
        End If
        If WS.Cells(a, 11) < Min_Change Then
            Min_Change = WS.Cells(a, 11)
            Min_Change_Ticker = WS.Cells(a, 9)
        End If
        If WS.Cells(a, 12) > Max_Volume Then
            Max_Volume = WS.Cells(a, 12)
            Max_Volume_Ticker = WS.Cells(a, 9)
        End If
    
    Next a

    WS.Cells(1, Summary_Col + 5).Value = "Category"
    WS.Cells(1, Summary_Col + 6).Value = "Ticker"
    WS.Cells(1, Summary_Col + 7).Value = "Value"
    
    WS.Cells(2, Summary_Col + 5).Value = "Greatest % Increase"
    WS.Cells(3, Summary_Col + 5).Value = "Greatest % Decrease"
    WS.Cells(4, Summary_Col + 5).Value = "Greatest Total Volume"
    
    WS.Cells(2, Summary_Col + 6).Value = Max_Change_Ticker
    WS.Cells(3, Summary_Col + 6).Value = Min_Change_Ticker
    WS.Cells(4, Summary_Col + 6).Value = Max_Volume_Ticker
    
    With WS.Cells(2, Summary_Col + 7)
        .Value = Max_Change
        .NumberFormat = "0.00%;(0.00%)"
    End With
    With WS.Cells(3, Summary_Col + 7)
        .Value = Min_Change
        .NumberFormat = "0.00%;(0.00%)"
    End With
    With WS.Cells(4, Summary_Col + 7)
        .Value = Max_Volume
        .NumberFormat = "#,##0;(#,##0)"
    End With
    
    'Debug.Print Max_Change_Ticker; Max_Change
    'Debug.Print Min_Change_Ticker; Min_Change
    'Debug.Print Max_Volume_Ticker; Max_Volume
    
End Sub
