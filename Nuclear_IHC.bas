Attribute VB_Name = "Nuclear_IHC"
Sub Nuclear_IHC_DataHandler()
'
' Nuclear_IHC_DataHandler Macro
'
'Delete unused data
'
Range("A:A,D:E,H:K,M:AA").Delete
'Column Headers
'
    Range("B1").Value = "Algorithm"
    Range("C1").Value = "Analysis ROI"
    Range("D1").Value = "Total Nuclei"
    Range("E1").Value = "IHC Positive"
    Range("F1").Value = "#Confirmed"
    Range("G1").Value = "%IHC Positive"
   
 '
 'Add Filter for slide types
    Range("A1", Range("A1").End(xlDown)).Select
    Selection.AutoFilter
 'Add individual % positive
    Range(Range("D2"), Range("D2").End(xlDown)).Select
    Selection.Offset(0, 3).Value = "=F2/D2*100"
    
 'Add formulas for summary statistics
 '
 'Number of slides
    Range("C1").End(xlDown).Offset(2, 0).Select
    Selection.Value = "Total Slides"
    ActiveCell.Offset(1, 0).FormulaR1C1 = "=SUBTOTAL(103,R2c3:r[-3]c)"
 'sum total nuclei
    ActiveCell.Offset(0, 1).Value = "Total Nuclei"
    ActiveCell.Offset(1, 1).FormulaR1C1 = "=SUBTOTAL(109,R2C4:R[-3]C)"
 'sum AI IHC positive nuclei
    ActiveCell.Offset(0, 2).Value = "#IHC positive"
    ActiveCell.Offset(1, 2).FormulaR1C1 = "=SUBTOTAL(109,R2C5:R[-3]C)"
 'sum confirmed positive nuclei
    ActiveCell.Offset(0, 3).Value = "#Confirmed"
    ActiveCell.Offset(1, 3).FormulaR1C1 = "=SUBTOTAL(109,R2C6:R[-3]C)"
 'percent positive nuclei
    ActiveCell.Offset(0, 4).Value = "%IHC Positive"
    ActiveCell.Offset(1, 4).FormulaR1C1 = "=RC[-1]/RC[-3]*100"
    
 'Repeat data labels for individual slide types
    ActiveCell.Offset(3, 0).Value = "#Slides"
    ActiveCell.Offset(3, 1).Value = "Total Nuclei"
    ActiveCell.Offset(3, 2).Value = "#IHC Positive"
    ActiveCell.Offset(3, 3).Value = "#Confirmed"
    ActiveCell.Offset(3, 4).Value = "%IHC Positive"
     
 'Results category labels
   'light blue
    Range("A1").End(xlDown).Offset(6, 0).Select
    Selection.Value = "IsoType Control"
    ActiveCell.Interior.Color = RGB(217, 225, 242)
   'dark blue
    Range("A1").End(xlDown).Offset(7, 0).Select
    Selection.Value = "Test"
    ActiveCell.Interior.Color = RGB(180, 198, 231)
   'light lavender
    Range("A1").End(xlDown).Offset(8, 0).Select
    Selection.Value = "0.01% Control"
    ActiveCell.Interior.Color = RGB(206, 205, 233)
   'dark lavender
    Range("A1").End(xlDown).Offset(9, 0).Select
    Selection.Value = "Control Test"
    ActiveCell.Interior.Color = RGB(183, 181, 221)
    
 'Reportable Results
 
    Range("A1").End(xlDown).Offset(12, 0).Select
    Selection.Font.Bold = True
    Selection.Value = "Test Nuclei>2M"
    ActiveCell.Offset(0, 1).FormulaR1C1 = "=IF(r[-5]c[2]>2000000,""YES"",""NO"")"
    ActiveCell.Offset(-5, 3).Font.Bold = True
    ActiveCell.Offset(0, 1).Font.Bold = True
    
    Range("A1").End(xlDown).Offset(13, 0).Select
    Selection.Font.Bold = True
    Selection.Value = "0.01% Control in Reference Range"
    ActiveCell.Offset(0, 1).FormulaR1C1 = "=IF(AND(r[-5]c[5]>r[1]c[3],r[-5]c[5]<r[1]c[4]),""YES"",""NO"")"
    ActiveCell.Offset(-5, 6).Font.Bold = True
    ActiveCell.Offset(0, 1).Font.Bold = True
    
    Range("A1").End(xlDown).Offset(14, 0).Select
    Selection.Font.Bold = True
    Selection.Value = "OCT4 Test Result"
    ActiveCell.Offset(0, 1).FormulaR1C1 = "=IF(r[-7]c[5]<0.0004,""<LLOD"",IF(r[-7]c[5]<0.0011,""<0.0011"",r[-7]c[5]))"
    ActiveCell.Offset(-7, 6).Font.Bold = True
    ActiveCell.Offset(0, 1).Font.Bold = True
    
 'Fill field for Current Reference range
    Range("D1").End(xlDown).Offset(12, 0).Select
    Selection.Value = "Reference Range for 0.01% Control"
    Selection.Offset(1, 0).Value = "Control"
    Selection.Offset(1, 1).Value = "lower"
    Selection.Offset(1, 2).Value = "upper"
 
 'add border box
    Range("C1").End(xlDown).Offset(2, 0).Select
    Selection.Resize(Selection.Rows.Count + 1, Selection.Columns.Count + 4).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Columns("A:G").Select
    Columns("A:G").EntireColumn.AutoFit
    
    
   'Colorfill image tag
    Range(Range("A1"), Range("A1").End(xlDown)).Select
    
   'color OCT4 test
    Selection.FormatConditions.Add Type:=xlTextString, String:="LOT71", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(180, 198, 231)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
   'color ISO control
    Selection.FormatConditions.Add Type:=xlTextString, String:="ISO", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(217, 225, 242)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
   'color 0.01%
    Selection.FormatConditions.Add Type:=xlTextString, String:="_0.01-", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(206, 205, 233)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    'color 0_01%
    Selection.FormatConditions.Add Type:=xlTextString, String:="_0_01", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(206, 205, 233)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    'color 0-01%
    Selection.FormatConditions.Add Type:=xlTextString, String:="_0-01", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(206, 205, 233)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    'color 0.01% test
    Selection.FormatConditions.Add Type:=xlTextString, String:="test", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(183, 181, 221)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
   'color 1%
    Selection.FormatConditions.Add Type:=xlTextString, String:="1pct", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(166, 166, 166)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    
   'Make fill color static
    Dim r
    For Each r In Selection
        r.Interior.Color = r.DisplayFormat.Interior.Color
    Next r
    Selection.FormatConditions.Delete
    
   'expand fill colour to data columns
    Range(Range("A2"), Range("A2").End(xlDown)).Select
    Selection.Copy
    Range(Range("B2"), Range("B2").End(xlDown)).Select
    Selection.Resize(Selection.Rows.Count + 0, Selection.Columns.Count + 5).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
       
   'expand fill colour to reporting cells
    Range("A2").End(xlDown).Offset(6, 0).Select
    Selection.Resize(Selection.Rows.Count + 3, Selection.Columns.Count + 0).Select
    Selection.Copy
    
    Range("C2").End(xlDown).Offset(6, 0).Select
    Selection.Resize(Selection.Rows.Count + 3, Selection.Columns.Count + 4).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
  
    
   'formatting
    Range(Range("D2"), Range("D2").End(xlDown)).Select
    Selection.Resize(Selection.Rows.Count + 0, Selection.Columns.Count + 3).Select
    Selection.NumberFormat = "#,##0"
        
    Range("D2").End(xlDown).Offset(3, 0).Select
    Selection.Resize(Selection.Rows.Count + 0, Selection.Columns.Count + 3).Select
    Selection.NumberFormat = "#,##0"
        
    Range("D2").End(xlDown).Offset(6, 0).Select
    Selection.Resize(Selection.Rows.Count + 3, Selection.Columns.Count + 3).Select
    Selection.NumberFormat = "#,##0"
    
    Range(Range("D2"), Range("D2").End(xlDown)).Select
    Selection.Offset(0, 3).NumberFormat = ".0000"
    
    Range("D2").End(xlDown).Offset(3, 0).Select
    Selection.Offset(0, 3).NumberFormat = ".0000"
    
    Range("D2").End(xlDown).Offset(6, 0).Select
    Selection.Resize(Selection.Rows.Count + 3).Select
    Selection.Offset(0, 3).NumberFormat = ".0000"
    
   'bold key results
    Range("D2").End(xlDown).Offset(7, 0).Font.Bold = True
    Range("D2").End(xlDown).Offset(7, 3).Font.Bold = True
    Range("D2").End(xlDown).Offset(8, 3).Font.Bold = True
             
   'filter results
   '0.01% dark lavender new control test results
    Range(Range("A1"), Range("A1").End(xlDown)).AutoFilter Field:=1, Criteria1:=RGB(183, _
        181, 221), Operator:=xlFilterCellColor
    Range(Cells.Address).Find(What:="Total Slides").Offset(1, 0).Select
    Selection.Resize(Selection.Rows.Count + 0, Selection.Columns.Count + 4).Select
    Selection.Copy
    Selection.Offset(6, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    '0.01% light lavender results
    Range(Range("A1"), Range("A1").End(xlDown)).AutoFilter Field:=1, Criteria1:=RGB(206, _
        205, 233), Operator:=xlFilterCellColor
    Range(Cells.Address).Find(What:="Total Slides").Offset(1, 0).Select
    Selection.Resize(Selection.Rows.Count + 0, Selection.Columns.Count + 4).Select
    Selection.Copy
    Selection.Offset(5, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    'OCT4 dark blue test results
    Range(Range("A1"), Range("A1").End(xlDown)).AutoFilter Field:=1, Criteria1:=RGB(180, _
        198, 231), Operator:=xlFilterCellColor
    Range(Cells.Address).Find(What:="Total Slides").Offset(1, 0).Select
    Selection.Resize(Selection.Rows.Count + 0, Selection.Columns.Count + 4).Select
    Selection.Copy
    Selection.Offset(4, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    'Isotype control light blue results
    Range(Range("A1"), Range("A1").End(xlDown)).AutoFilter Field:=1, Criteria1:=RGB(217, _
        225, 242), Operator:=xlFilterCellColor
    Range(Cells.Address).Find(What:="Total Slides").Offset(1, 0).Select
    Selection.Resize(Selection.Rows.Count + 0, Selection.Columns.Count + 4).Select
    Selection.Copy
    Selection.Offset(3, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    ActiveSheet.ShowAllData
  
    'Signature box
    
     Range(Cells.Address).Find(What:="OCT4 Test Result").Offset(2, 0).Select
     Selection.Resize(Selection.Rows.Count + 1, Selection.Columns.Count + 0).Select
     Selection.Merge
     
     Selection.Offset(1, 0).Select
     Selection.Resize(Selection.Rows.Count + 1, Selection.Columns.Count + 0).Select
     Selection.Merge
     
     Selection.Offset(1, 0).Select
     Selection.Resize(Selection.Rows.Count + 1, Selection.Columns.Count + 0).Select
     Selection.Merge
     
     Selection.Offset(0, 1).Select
     Selection.Resize(Selection.Rows.Count + 1, Selection.Columns.Count + 1).Select
     Selection.Merge
     
     Selection.Offset(-2, 0).Select
     Selection.Resize(Selection.Rows.Count + 1, Selection.Columns.Count + 1).Select
     Selection.Merge
     
     Selection.Offset(-2, 0).Select
     Selection.Resize(Selection.Rows.Count + 1, Selection.Columns.Count + 1).Select
     Selection.Merge
     
     Selection.Offset(0, -1).Select
     Selection.Resize(Selection.Rows.Count + 4, Selection.Columns.Count + 1).Select
   
     
     
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Range(Cells.Address).Find(What:="OCT4 Test Result").Offset(2, 0).Select
    Selection.Value = "Performed By"
    Selection.Offset(1, 0).Value = "Data Verification"
    Selection.Offset(3, 0).Value = "Quality Review"
    Selection.Offset(0, 1).Value = "Date"
    Selection.Offset(1, 1).Value = "Date"
    Selection.Offset(3, 1).Value = "Date"
    Selection.Resize(Selection.Rows.Count + 2, Selection.Columns.Count + 1).Select
    
    Range("A1").Select
End Sub

   





