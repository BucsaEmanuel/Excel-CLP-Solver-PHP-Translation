Attribute VB_Name = "Module1"
'This work is licensed under the Creative Commons Attribution 4.0 International License. To view a copy of this license, visit http://creativecommons.org/licenses/by/4.0/.

Option Explicit
Const epsilon As Double = 0.0001
Const column_offset As Long = 11
Const width_limit As Double = 300
Const height_limit As Double = 300
Const displacement_multiplier As Double = 0.353

Function CheckWorksheetExistence(sheetName As String) As Boolean
    
    Dim WorksheetExists As Boolean
    WorksheetExists = False
    Dim Sht As Worksheet
    For Each Sht In ThisWorkbook.Worksheets
        If (Sht.Name = sheetName) Then
            WorksheetExists = True
        End If
    Next Sht
    
    CheckWorksheetExistence = WorksheetExists
End Function

Sub SetupConsoleWorksheet()

    Dim WorksheetExists As Boolean
    WorksheetExists = CheckWorksheetExistence("CLP Solver Console")
    
    Dim reply As Integer
    If WorksheetExists = False Then
        ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = "CLP Solver Console"
    End If
    
    ThisWorkbook.Worksheets("CLP Solver Console").Activate
    
    'Problem parameters
    
    Cells(1, 1).Value = "Sequence"
    Cells(1, 1).Select
    Selection.Font.Bold = True
    
    Cells(1, 2).Value = "Parameter"
    Cells(1, 3).Value = "Value"
    Cells(1, 4).Value = "Remarks"
    
    Range(Cells(1, 1), Cells(1, 4)).Select
    Selection.Interior.ColorIndex = 1
    Selection.Font.ColorIndex = 2
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    
    'GIS License
    
    Cells(2, 1).Value = "1.Items"
    Cells(2, 1).Select
    Selection.Font.Bold = True
    
    Cells(2, 2).Value = "Number of types of items"
    
    Cells(2, 3).Validation.Delete
    With Cells(2, 3).Validation
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="100"
        .ErrorMessage = "Please enter an integer from 1 to 100"
    End With
    
    If Cells(2, 3).Value = "" Then
        Cells(2, 3).Value = 1
    End If
    
    Cells(2, 3).Font.ColorIndex = 2
    Cells(2, 3).Interior.ColorIndex = 50
    Cells(2, 4).Value = "[1,100]"
    
    Range(Cells(2, 1), Cells(2, 4)).Select
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    
    Cells(4, 1).Value = "2.Containers"
    Cells(4, 1).Select
    Selection.Font.Bold = True
    
    Cells(4, 2).Value = "Number of types of containers"
    Cells(4, 3).Validation.Delete
    With Cells(4, 3).Validation
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="10"
        .ErrorMessage = "Please enter an integer from 1 to 10"
    End With
    
    If Cells(4, 3).Value = "" Then
        Cells(4, 3).Value = 1
    End If
    
    Cells(4, 3).Font.ColorIndex = 2
    Cells(4, 3).Interior.ColorIndex = 50
    Cells(4, 4).Value = "[1,10]"
    
    Range(Cells(4, 1), Cells(4, 4)).Select
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    
    Cells(6, 1).Value = "3.Solution"
    Cells(6, 1).Select
    Selection.Font.Bold = True
    
    Cells(6, 2).Value = "Front side support required?"
    
    Cells(6, 3).Validation.Delete
    With Cells(6, 3).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Yes, No"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    Cells(6, 3).Value = "No"
    Cells(6, 3).Font.ColorIndex = 2
    Cells(6, 3).Interior.ColorIndex = 50
    Cells(6, 4).Value = ""
    
    Range(Cells(6, 1), Cells(6, 4)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    
    'Visualization parameters
    
    Cells(8, 1).Value = "4.Visualization"
    Cells(8, 1).Select
    Selection.Font.Bold = True
    
    Cells(8, 2).Value = "Item labels"
    
    Cells(8, 3).Validation.Delete
    With Cells(8, 3).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Yes, No"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    Cells(8, 3).Value = "Yes"
    Cells(8, 3).Font.ColorIndex = 2
    Cells(8, 3).Interior.ColorIndex = 50

    Cells(9, 2).Value = "Container labels"
        
    Cells(9, 3).Validation.Delete
    With Cells(9, 3).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Yes, No"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    Cells(9, 3).Value = "No"
    Cells(9, 3).Font.ColorIndex = 2
    Cells(9, 3).Interior.ColorIndex = 50
    
    Cells(10, 2).Value = "Animation advances by:"
    
    Cells(10, 3).Validation.Delete
    With Cells(10, 3).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="One second intervals, Message box between every two items"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    Cells(10, 3).Value = "One second intervals"
    Cells(10, 3).Font.ColorIndex = 2
    Cells(10, 3).Interior.ColorIndex = 50
    
    Range(Cells(8, 1), Cells(10, 4)).Select
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    
    'Algorithmic parameters
    
    Cells(12, 1).Value = "5.Solver"
    Cells(12, 1).Select
    Selection.Font.Bold = True
    
    Cells(12, 2).Value = "First-Fit-Decreasing based on:"
    
    Cells(12, 3).Validation.Delete
    With Cells(12, 3).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Volume, Weight, Max {width; height; length}"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    Cells(12, 3).Value = "Volume"
    Cells(12, 3).Font.ColorIndex = 2
    Cells(12, 3).Interior.ColorIndex = 50
    
    Cells(13, 2).Value = "Constructive heuristic:"
    
    Cells(13, 3).Validation.Delete
    With Cells(13, 3).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Wall-building, Layer-building"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    Cells(13, 3).Value = "Wall-building"
    Cells(13, 3).Font.ColorIndex = 2
    Cells(13, 3).Interior.ColorIndex = 50
       
    Cells(13, 4).Value = ""
        
    Cells(14, 2).Value = "CPU time limit (seconds)"
    
    Cells(14, 3).Validation.Delete
    With Cells(14, 3).Validation
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="10"
        .ErrorMessage = "Please enter a value greater than or equal to 10 seconds."
    End With
    
    If Cells(14, 3).Value = "" Then
        Cells(14, 3).Value = 60
    End If
    Cells(14, 3).Font.ColorIndex = 2
    Cells(14, 3).Interior.ColorIndex = 50
    
    Cells(14, 4).Value = "Recommendation: At least one second per item."
            
    Range(Cells(12, 1), Cells(14, 4)).Select
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    
    Columns("A").AutoFit
    Columns("B").AutoFit
    'Columns("C").AutoFit
    Columns("D").AutoFit
    
    Columns("C").EntireColumn.ColumnWidth = 35
     
    Rows("1").Select
    Selection.Font.Bold = True
    Cells(1, 1).Select
End Sub
Sub SetupItemsWorksheet()
   
    Dim WorksheetExists As Boolean
    WorksheetExists = CheckWorksheetExistence("1.Items")
    
    Dim reply As Integer
    If WorksheetExists = True Then
        reply = MsgBox("This will overwrite existing item data, and erase container, solution, and visualization data. Do you want to continue?", vbYesNo, "CLP Spreadsheet Solver")
        If reply = vbNo Then
            Exit Sub
        Else
            Application.DisplayAlerts = False
            
            ThisWorkbook.Worksheets("1.Items").Delete
            ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = "1.Items"
                        
            WorksheetExists = CheckWorksheetExistence("1.3.Item-Item Compatibility")
            If WorksheetExists = True Then
                ThisWorkbook.Worksheets("1.3.Item-Item Compatibility").Delete
            End If
            
            WorksheetExists = CheckWorksheetExistence("2.Containers")
            If WorksheetExists = True Then
                ThisWorkbook.Worksheets("2.Containers").Delete
            End If
            
            WorksheetExists = CheckWorksheetExistence("2.3.Container-ItemCompatibility")
            If WorksheetExists = True Then
                ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Delete
            End If
            
            WorksheetExists = CheckWorksheetExistence("3.Solution")
            If WorksheetExists = True Then
                ThisWorkbook.Worksheets("3.Solution").Delete
            End If
            
            WorksheetExists = CheckWorksheetExistence("4.Visualization")
            If WorksheetExists = True Then
                ThisWorkbook.Worksheets("4.Visualization").Delete
            End If
            
            Application.DisplayAlerts = True
        End If
    Else
        ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "1.Items"
    End If
    
    ThisWorkbook.Worksheets("1.Items").Activate
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim formula_delimiter As String
    If InStr(CStr(WorksheetFunction.Pi), ".") > 0 Then
        formula_delimiter = ","
    Else
        formula_delimiter = ";"
    End If
    
    Dim num_item_types As Long
    num_item_types = ThisWorkbook.Worksheets("CLP Solver Console").Cells(2, 3).Value
    
    Dim formulaText As String
    
    Cells(1, 8).Value = "Can be placed on:"
    
    Cells(2, 1).Value = "Item Type ID"
    Cells(2, 2).Value = "Name"
    Cells(2, 3).Value = "Colour / Image file name"
    Cells(2, 4).Value = "Width (x)"
    Cells(2, 5).Value = "Height (y)"
    Cells(2, 6).Value = "Length (z)"
    Cells(2, 7).Value = "Volume         "
    Cells(2, 8).Value = "x/z surface?"
    Cells(2, 9).Value = "x/y surface?"
    Cells(2, 10).Value = "y/z surface?"
    Cells(2, 11).Value = "Weight         "
    Cells(2, 12).Value = "Heavy item?"
    Cells(2, 13).Value = "Fragile item?"
    Cells(2, 14).Value = "Must be packed?"
    Cells(2, 15).Value = "Profit"
    Cells(2, 16).Value = "Number of items"
    
    Range(Cells(1, 1), Cells(2, 16)).Select
    
    Selection.Font.Bold = True
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    
    Range(Cells(1, 8), Cells(2, 10)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    Selection.Interior.ColorIndex = 15
    
    Dim i As Long
    
    For i = 1 To num_item_types
    
        Range(Cells(2 + i, 1), Cells(2 + i, 16)).Select
        Selection.Font.ColorIndex = 2
        Selection.Interior.ColorIndex = 50
        
        If i Mod 2 = 0 Then
            Selection.Interior.TintAndShade = 0.02
        Else
            Selection.Interior.TintAndShade = -0.02
        End If
        
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 16
            .TintAndShade = 0
            .weight = xlThin
        End With
        
        Cells(2 + i, 1).Value = i
        Cells(2 + i, 1).Font.ColorIndex = 2
        Cells(2 + i, 1).Interior.ColorIndex = 1
        
        Cells(2 + i, 1).FormatConditions.Add Type:=xlExpression, Formula1:="=ISBLANK(" & Cells(2 + i, 1).Address & ")"
        Cells(2 + i, 1).FormatConditions(1).Interior.ColorIndex = 3
        
        Cells(2 + i, 2).Value = "Item type " & i
        
        Cells(2 + i, 2).FormatConditions.Add Type:=xlExpression, Formula1:="=ISBLANK(" & Cells(2 + i, 2).Address & ")"
        Cells(2 + i, 2).FormatConditions(1).Interior.ColorIndex = 3
        
        Cells(2 + i, 3).Value = ""
        Cells(2 + i, 3).Interior.ColorIndex = (3 + i) Mod 57
        
        Cells(2 + i, 4).NumberFormat = "0.00"
        Cells(2 + i, 4).Value = 1
        
        With Cells(2 + i, 4).Validation
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:=0
            .ErrorMessage = "Please enter a positive value"
        End With
        
        Cells(2 + i, 5).NumberFormat = "0.00"
        Cells(2 + i, 5).Value = 1
        
        With Cells(2 + i, 5).Validation
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:=0
            .ErrorMessage = "Please enter a positive value"
        End With
        
        Cells(2 + i, 6).NumberFormat = "0.00"
        Cells(2 + i, 6).Value = 1
        
        With Cells(2 + i, 6).Validation
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:=0
            .ErrorMessage = "Please enter a positive value"
        End With
        
        Cells(2 + i, 7).NumberFormat = "0.00"
        Cells(2 + i, 7).Formula = "=" & Cells(2 + i, 4).Address(False, False) & "*" & Cells(2 + i, 5).Address(False, False) & "*" & Cells(2 + i, 6).Address(False, False)
        Cells(2 + i, 7).Interior.ColorIndex = 36
        Cells(2 + i, 7).Font.ColorIndex = 1
        
        Cells(2 + i, 8).Value = "Yes"
        Cells(2 + i, 8).Interior.ColorIndex = 1
        Cells(2 + i, 8).Font.ColorIndex = 2
        
        With Cells(2 + i, 9).Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Yes, No"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        Cells(2 + i, 9).Value = "Yes"
        
        Cells(2 + i, 9).FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="No"
        Cells(2 + i, 9).FormatConditions(1).Interior.ColorIndex = 43
        
        With Cells(2 + i, 10).Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Yes, No"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        Cells(2 + i, 10).Value = "Yes"
        
        Cells(2 + i, 10).FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="No"
        Cells(2 + i, 10).FormatConditions(1).Interior.ColorIndex = 43
        
        Cells(2 + i, 11).NumberFormat = "0.00"
        Cells(2 + i, 11).Value = 0
        
        With Cells(2 + i, 11).Validation
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:=-epsilon
            .ErrorMessage = "Please enter a nonnegative value"
        End With
        
        With Cells(2 + i, 12).Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Yes, No"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        Cells(2 + i, 12).Value = "No"
        
        Cells(2 + i, 12).FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="Yes"
        Cells(2 + i, 12).FormatConditions(1).Interior.ColorIndex = 4
        
        With Cells(2 + i, 13).Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Yes, No"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        Cells(2 + i, 13).Value = "No"
        
        Cells(2 + i, 13).FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="Yes"
        Cells(2 + i, 13).FormatConditions(1).Interior.ColorIndex = 4
        
        With Cells(2 + i, 14).Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Must be packed, May be packed, Do not pack"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        Cells(2 + i, 14).Value = "Must be packed"
        
        Cells(2 + i, 14).FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="May be packed"
        Cells(2 + i, 14).FormatConditions(1).Interior.ColorIndex = 4
        
        Cells(2 + i, 14).FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="Do not pack"
        Cells(2 + i, 14).FormatConditions(2).Interior.ColorIndex = 43
        
        Cells(2 + i, 15).NumberFormat = "0.00"
        Cells(2 + i, 15).Value = 0
        
        With Cells(2 + i, 15).Validation
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:=-epsilon
            .ErrorMessage = "Please enter a nonnegative value"
        End With
                
        Cells(2 + i, 16).Value = 1
        
        With Cells(2 + i, 16).Validation
            .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="0"
            .ErrorMessage = "Please enter a positive integer value"
        End With
    Next i
    
'    Cells(2, 15).Value = "Total volume requirement:"
'    Cells(3, 15).Formula = "=SUMPRODUCT(" & Cells(3, 7).Address(False, False) & ":" & Cells(2 + num_item_types, 7).Address(False, False) & "," & Cells(3, 14).Address(False, False) & ":" & Cells(2 + num_item_types, 14).Address(False, False) & ")"
'
'    Cells(3, 15).Interior.ColorIndex = 36
'    Cells(3, 15).Font.ColorIndex = 1
'    Cells(3, 15).NumberFormat = "0.00"
    
    ' Display all duplicate values in red
    Dim uv As UniqueValues
    Set uv = Range(Cells(2, 2), Cells(1 + num_item_types, 2)).FormatConditions.AddUniqueValues
    uv.DupeUnique = xlDuplicate
    uv.Interior.ColorIndex = 3
    uv.StopIfTrue = False
    
    Dim item_width As Double
    Dim item_height As Double
    Dim item_length As Double
    
    Dim nw_x As Double
    Dim nw_y As Double
    
    item_width = Columns("D").width + Columns("E").width + Columns("F").width
    item_height = item_width
    item_length = item_width
    
    nw_x = Columns("A").width + Columns("B").width + Columns("C").width
    nw_y = ((num_item_types + 3) * Rows(1).height) + item_length * displacement_multiplier
    
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, nw_x, nw_y, item_width, item_height).Select 'type, left, top, width, height
    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).Fill.Transparency = 1
    
    With Selection.ShapeRange.ThreeD
        .Visible = True
        .depth = item_length
        .ContourWidth = 0.2
        .ExtrusionColor.RGB = RGB(255, 255, 255) 'for containers
        .SetExtrusionDirection msoExtrusionTopRight
        .PresetLightingDirection = msoLightingTopRight
        .PresetMaterial = msoMaterialTranslucentPowder 'for containers
    End With
    
    
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, nw_x + item_width / 2, nw_y + item_height, Columns("D").width, 2 * Rows(1).height).TextFrame.Characters.Text = "x"
    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).TextFrame.Characters.Font.Size = 11
    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).TextFrame.Characters.Font.Bold = True
    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).Fill.Visible = msoFalse
    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).Line.Visible = msoFalse
    
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, nw_x, nw_y + item_height / 2, Columns("D").width, 2 * Rows(1).height).TextFrame.Characters.Text = "y"
    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).TextFrame.Characters.Font.Size = 11
    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).TextFrame.Characters.Font.Bold = True
    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).Fill.Visible = msoFalse
    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).Line.Visible = msoFalse
            
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, nw_x + item_width * (1 + displacement_multiplier / 2), nw_y + item_height * 0.75, Columns("D").width, 2 * Rows(1).height).TextFrame.Characters.Text = "z"
    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).TextFrame.Characters.Font.Size = 11
    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).TextFrame.Characters.Font.Bold = True
    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).Fill.Visible = msoFalse
    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).Line.Visible = msoFalse
    
    Columns("A:P").AutoFit
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Cells(1, 1).Select
End Sub
Sub SetupContainersWorksheet()
   
    Dim WorksheetExists As Boolean
    WorksheetExists = CheckWorksheetExistence("2.Containers")
    
    Dim reply As Integer
    If WorksheetExists = True Then
        reply = MsgBox("This will overwrite existing container data, and erase solution and visualization data. Do you want to continue?", vbYesNo, "CLP Spreadsheet Solver")
        If reply = vbNo Then
            Exit Sub
        Else
            Application.DisplayAlerts = False
            
            ThisWorkbook.Worksheets("2.Containers").Delete
            ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = "2.Containers"
            
            WorksheetExists = CheckWorksheetExistence("2.3.Container-ItemCompatibility")
            If WorksheetExists = True Then
                ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Delete
            End If
            
            WorksheetExists = CheckWorksheetExistence("3.Solution")
            If WorksheetExists = True Then
                ThisWorkbook.Worksheets("3.Solution").Delete
            End If
            
            WorksheetExists = CheckWorksheetExistence("4.Visualization")
            If WorksheetExists = True Then
                ThisWorkbook.Worksheets("4.Visualization").Delete
            End If
            
            Application.DisplayAlerts = True
        End If
    Else
        ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "2.Containers"
    End If
    
    ThisWorkbook.Worksheets("2.Containers").Activate
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim num_container_types As Long
    num_container_types = ThisWorkbook.Worksheets("CLP Solver Console").Cells(4, 3).Value
    Dim formulaText As String
    
    Cells(1, 1).Value = "Container Type ID"
    Cells(1, 2).Value = "Name"
    Cells(1, 3).Value = "Width (x)"
    Cells(1, 4).Value = "Height (y)"
    Cells(1, 5).Value = "Length (z)"
    Cells(1, 6).Value = "Volume capacity"
    Cells(1, 7).Value = "Weight capacity"
    Cells(1, 8).Value = "May be used?"
    Cells(1, 9).Value = "Cost"
    Cells(1, 10).Value = "Number of containers"
    
    Range(Cells(1, 1), Cells(1, 10)).Select
    
    Selection.Font.Bold = True
    
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    
    Dim i As Long
    
    For i = 1 To num_container_types
    
        Range(Cells(1 + i, 1), Cells(1 + i, 10)).Select
        Selection.Font.ColorIndex = 2
        Selection.Interior.ColorIndex = 50
        
        If i Mod 2 = 0 Then
            Selection.Interior.TintAndShade = 0.02
        Else
            Selection.Interior.TintAndShade = -0.02
        End If
        
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 16
            .TintAndShade = 0
            .weight = xlThin
        End With
        
        Cells(1 + i, 1).Value = i
        Cells(1 + i, 1).Font.ColorIndex = 2
        Cells(1 + i, 1).Interior.ColorIndex = 1
        
        Cells(1 + i, 1).FormatConditions.Add Type:=xlExpression, Formula1:="=ISBLANK(" & Cells(1 + i, 1).Address(False, False) & ")"
        Cells(1 + i, 1).FormatConditions(1).Interior.ColorIndex = 3
        
        Cells(1 + i, 2).Value = "Container type " & i
        
        Cells(1 + i, 2).FormatConditions.Add Type:=xlExpression, Formula1:="=ISBLANK(" & Cells(1 + i, 2).Address(False, False) & ")"
        Cells(1 + i, 2).FormatConditions(1).Interior.ColorIndex = 3
        
        Cells(1 + i, 3).NumberFormat = "0.00"
        Cells(1 + i, 3).Value = 1
        
        With Cells(1 + i, 3).Validation
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:=0
            .ErrorMessage = "Please enter a positive value"
        End With
        
        Cells(1 + i, 4).NumberFormat = "0.00"
        Cells(1 + i, 4).Value = 1
        
        With Cells(1 + i, 4).Validation
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:=0
            .ErrorMessage = "Please enter a positive value"
        End With
        
        Cells(1 + i, 5).NumberFormat = "0.00"
        Cells(1 + i, 5).Value = 1
        
        With Cells(1 + i, 5).Validation
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:=0
            .ErrorMessage = "Please enter a positive value"
        End With
        
        Cells(1 + i, 6).NumberFormat = "0.00"
        Cells(1 + i, 6).Formula = "=" & Cells(1 + i, 3).Address(False, False) & "*" & Cells(1 + i, 4).Address(False, False) & "*" & Cells(1 + i, 5).Address(False, False)
        Cells(1 + i, 6).Interior.ColorIndex = 36
        Cells(1 + i, 6).Font.ColorIndex = 1
        
        Cells(1 + i, 7).NumberFormat = "0.00"
        Cells(1 + i, 7).Value = 1
        
        With Cells(1 + i, 7).Validation
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:=0
            .ErrorMessage = "Please enter a positive value"
        End With
        
        With Cells(1 + i, 8).Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="May be used, Do not use" 'removed Must be used
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        Cells(1 + i, 8).Value = "May be used"
        
        Cells(1 + i, 8).FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="Must be used"
        Cells(1 + i, 8).FormatConditions(1).Interior.ColorIndex = 4
        
        Cells(1 + i, 8).FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="Do not use"
        Cells(1 + i, 8).FormatConditions(2).Interior.ColorIndex = 43
        
        Cells(1 + i, 9).NumberFormat = "0.00"
        Cells(1 + i, 9).Value = 1
        
        With Cells(1 + i, 9).Validation
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:=-epsilon
            .ErrorMessage = "Please enter a nonnegative value"
        End With
                
        Cells(1 + i, 10).Value = 1
        
        With Cells(1 + i, 10).Validation
            .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="0"
            .ErrorMessage = "Please enter a positive integer value"
        End With
    Next i
    
'    Cells(1, 11).Value = "Total volume capacity:"
'    Cells(2, 11).Formula = "=SUMPRODUCT(" & Cells(2, 6).Address(False, False) & ":" & Cells(1 + num_container_types, 6).Address(False, False) & "," & Cells(2, 10).Address(False, False) & ":" & Cells(1 + num_container_types, 10).Address(False, False) & ")"
'
'    Cells(2, 11).Interior.ColorIndex = 36
'    Cells(2, 11).Font.ColorIndex = 1
'    Cells(2, 11).NumberFormat = "0.00"
    
    ' Display all duplicate values in red
    Dim uv As UniqueValues
    Set uv = Range(Cells(2, 2), Cells(1 + num_container_types, 2)).FormatConditions.AddUniqueValues
    uv.DupeUnique = xlDuplicate
    uv.Interior.ColorIndex = 3
    uv.StopIfTrue = False
    
    Columns("A:J").AutoFit
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Cells(1, 1).Select
End Sub
Sub SetupItemItemCompatibilityWorksheet()
   
    Dim WorksheetExists As Boolean
    
    WorksheetExists = CheckWorksheetExistence("1.Items")
    If WorksheetExists = False Then
        MsgBox "Worksheet 1.Items must exist for the item-item compatibility worksheet to be setup."
        Exit Sub
    End If
    
    Dim num_item_types As Long
    num_item_types = ThisWorkbook.Worksheets("CLP Solver Console").Cells(2, 3).Value
    
    If num_item_types = 1 Then
        MsgBox "There should be more than one type of item for the item-item compatibility worksheet to be setup."
        Exit Sub
    End If
    
    WorksheetExists = CheckWorksheetExistence("1.3.Item-Item Compatibility")
    
    Dim reply As Integer
    If WorksheetExists = True Then
        reply = MsgBox("This will overwrite existing item-item compatibility data. Do you want to continue?", vbYesNo, "CLP Spreadsheet Solver")
        If reply = vbNo Then
            Exit Sub
        Else
            Application.DisplayAlerts = False
            
            ThisWorkbook.Worksheets("1.3.Item-Item Compatibility").Delete
            ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = "1.3.Item-Item Compatibility"
                        
            Application.DisplayAlerts = True
        End If
    Else
        ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "1.3.Item-Item Compatibility"
    End If
    
    ThisWorkbook.Worksheets("1.3.Item-Item Compatibility").Activate
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Cells(2, 1).Value = "Item Type 1"
    Cells(2, 2).Value = "Item type 2"
    Cells(2, 3).Value = "Compatible?"
    
    Range(Cells(1, 1), Cells(2, 3)).Select
    
    Selection.Font.Bold = True
    
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    k = 3
    For i = 1 To num_item_types
    
        For j = i + 1 To num_item_types
        
            ThisWorkbook.Worksheets("1.3.Item-Item Compatibility").Range(Cells(k, 1), Cells(k, 3)).Select
            Selection.Font.ColorIndex = 2
            
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 16
                .TintAndShade = 0
                .weight = xlThin
            End With
            
            ThisWorkbook.Worksheets("1.3.Item-Item Compatibility").Cells(k, 1).Value = ThisWorkbook.Sheets("1.Items").Cells(2 + i, 2).Value
            ThisWorkbook.Worksheets("1.3.Item-Item Compatibility").Cells(k, 1).Font.ColorIndex = 2
            ThisWorkbook.Worksheets("1.3.Item-Item Compatibility").Cells(k, 1).Interior.ColorIndex = 1
            
            ThisWorkbook.Worksheets("1.3.Item-Item Compatibility").Cells(k, 2).Value = ThisWorkbook.Sheets("1.Items").Cells(2 + j, 2).Value
            ThisWorkbook.Worksheets("1.3.Item-Item Compatibility").Cells(k, 2).Font.ColorIndex = 2
            ThisWorkbook.Worksheets("1.3.Item-Item Compatibility").Cells(k, 2).Interior.ColorIndex = 1
            
            With ThisWorkbook.Worksheets("1.3.Item-Item Compatibility").Cells(k, 3).Validation
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Yes, No"
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
            Cells(k, 3).Value = "Yes"
            
            ThisWorkbook.Worksheets("1.3.Item-Item Compatibility").Cells(k, 3).Font.ColorIndex = 2
            ThisWorkbook.Worksheets("1.3.Item-Item Compatibility").Cells(k, 3).Interior.ColorIndex = 50
            
            k = k + 1
        
        Next j
    Next i
    
    ThisWorkbook.Worksheets("1.3.Item-Item Compatibility").Activate
    
    Range(Cells(1, 3), Cells(k - 1, 3)).Select
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="No"
    Selection.FormatConditions(1).Interior.ColorIndex = 43
        
    Columns("A").AutoFit
    Columns("B").AutoFit
    Columns("C").AutoFit
    
    Cells(1, 1).Value = "Two items are compatible if they can be placed into the same container."
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Cells(1, 1).Select
End Sub

Sub SetupContainerItemCompatibilityWorksheet()
   
    Dim WorksheetExists As Boolean
    
    WorksheetExists = CheckWorksheetExistence("1.Items") And CheckWorksheetExistence("2.Containers")
    If WorksheetExists = False Then
        MsgBox "Worksheets 1.Items and 2.Containers must exist for the container-item compatibility worksheet to be setup."
        Exit Sub
    End If
    
    WorksheetExists = CheckWorksheetExistence("2.3.Container-ItemCompatibility")
    Dim reply As Integer
    If WorksheetExists = True Then
        reply = MsgBox("This will overwrite existing container-item compatibility data. Do you want to continue?", vbYesNo, "CLP Spreadsheet Solver")
        If reply = vbNo Then
            Exit Sub
        Else
            Application.DisplayAlerts = False
            
            ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Delete
            ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = "2.3.Container-ItemCompatibility"
                        
            Application.DisplayAlerts = True
        End If
    Else
        ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "2.3.Container-ItemCompatibility"
    End If
    
    ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Activate
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim num_item_types As Long
    num_item_types = ThisWorkbook.Worksheets("CLP Solver Console").Cells(2, 3).Value
    
    Dim num_container_types As Long
    num_container_types = ThisWorkbook.Worksheets("CLP Solver Console").Cells(4, 3).Value
    
    Cells(2, 1).Value = "Container Type"
    Cells(2, 2).Value = "Item type"
    Cells(2, 3).Value = "Compatible?"
    
    Range(Cells(1, 1), Cells(2, 3)).Select
    
    Selection.Font.Bold = True
    
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 16
        .TintAndShade = 0
        .weight = xlThin
    End With
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    k = 3
    For i = 1 To num_container_types
    
        For j = 1 To num_item_types
        
            ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Range(Cells(k, 1), Cells(k, 3)).Select
            Selection.Font.ColorIndex = 2
            
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 16
                .TintAndShade = 0
                .weight = xlThin
            End With
            
            ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Cells(k, 1).Value = ThisWorkbook.Sheets("2.Containers").Cells(1 + i, 2).Value
            ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Cells(k, 1).Font.ColorIndex = 2
            ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Cells(k, 1).Interior.ColorIndex = 1
            
            ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Cells(k, 2).Value = ThisWorkbook.Sheets("1.Items").Cells(2 + j, 2).Value
            ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Cells(k, 2).Font.ColorIndex = 2
            ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Cells(k, 2).Interior.ColorIndex = 1
            
            With ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Cells(k, 3).Validation
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Yes, No"
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
            
            Cells(k, 3).Value = "Yes"
            
            If ((ThisWorkbook.Sheets("1.Items").Cells(2 + j, 4).Value > ThisWorkbook.Sheets("2.Containers").Cells(1 + i, 3).Value) And (ThisWorkbook.Sheets("1.Items").Cells(2 + j, 4).Value > ThisWorkbook.Sheets("2.Containers").Cells(1 + i, 4).Value)) _
             Or ((ThisWorkbook.Sheets("1.Items").Cells(2 + j, 5).Value > ThisWorkbook.Sheets("2.Containers").Cells(1 + i, 3).Value) And (ThisWorkbook.Sheets("1.Items").Cells(2 + j, 5).Value > ThisWorkbook.Sheets("2.Containers").Cells(1 + i, 4).Value)) Then
                Cells(k, 3).Value = "No"
            End If
            
            If ((ThisWorkbook.Sheets("1.Items").Cells(2 + j, 4).Value > ThisWorkbook.Sheets("2.Containers").Cells(1 + i, 3).Value) Or _
               (ThisWorkbook.Sheets("1.Items").Cells(2 + j, 4).Value > ThisWorkbook.Sheets("2.Containers").Cells(1 + i, 4).Value)) And _
               (ThisWorkbook.Sheets("1.Items").Cells(2 + j, 7).Value = "No") Then
                Cells(k, 3).Value = "No"
            End If
            
            ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Cells(k, 3).Font.ColorIndex = 2
            ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Cells(k, 3).Interior.ColorIndex = 50
            
            k = k + 1
        
        Next j
    Next i
    
    ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Activate
    
    Range(Cells(1, 3), Cells(k - 1, 3)).Select
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="No"
    Selection.FormatConditions(1).Interior.ColorIndex = 43
        
    Columns("A").AutoFit
    Columns("B").AutoFit
    Columns("C").AutoFit
    
    Cells(1, 1).Value = "A container is compatible with an item if the item can be placed into the container."
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Cells(1, 1).Select
End Sub
Sub SetupSolutionWorksheet()
    
    Dim WorksheetExists As Boolean
    Dim reply As Integer
    
    WorksheetExists = CheckWorksheetExistence("1.Items") And CheckWorksheetExistence("2.Containers")
    If WorksheetExists = False Then
        MsgBox "Worksheets 1.Items and 2.Containers must exist for the solution worksheet to be setup."
        Exit Sub
    End If
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim num_item_types As Long
    num_item_types = ThisWorkbook.Worksheets("CLP Solver Console").Cells(2, 3).Value
    
    Dim num_items As Long
    
    num_items = 0
    For i = 1 To num_item_types
        num_items = num_items + ThisWorkbook.Worksheets("1.Items").Cells(2 + i, 16).Value
    Next i
    
    If num_items = 0 Then
        MsgBox "There must be at least one item for the solution worksheet to be setup."
        Exit Sub
    End If
    
    Dim num_container_types As Long
    num_container_types = ThisWorkbook.Worksheets("CLP Solver Console").Cells(4, 3).Value
    
    Dim num_containers As Long
    
    num_containers = 0
    For i = 1 To num_container_types
        num_containers = num_containers + ThisWorkbook.Worksheets("2.Containers").Cells(1 + i, 10).Value
    Next i
    
    If num_containers = 0 Then
        MsgBox "There must be at least one container for the solution worksheet to be setup."
        Exit Sub
    End If
    
    Dim estimatedTime As Double
    estimatedTime = 0
    
    estimatedTime = num_items * num_containers * 0.03
    
    WorksheetExists = CheckWorksheetExistence("3.Solution")
    If WorksheetExists = True Then
        WorksheetExists = CheckWorksheetExistence("4.Visualization")
        If WorksheetExists = True Then
            reply = MsgBox("This will take some time (estimated " & estimatedTime & " seconds), overwrite existing solution data, and erase the visualization data. Do you want to continue?", vbYesNo, "CLP Spreadsheet Solver")
            If reply = vbNo Then
                ThisWorkbook.Worksheets("3.Solution").Activate
                Exit Sub
            Else
                Application.DisplayAlerts = False
            
                ThisWorkbook.Worksheets("4.Visualization").Delete
                
                ThisWorkbook.Worksheets("3.Solution").Delete
                ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "3.Solution"
                
                Application.DisplayAlerts = True
            End If
        Else
            reply = MsgBox("This will take some time (estimated " & estimatedTime & " seconds), and overwrite existing solution data. Do you want to continue?", vbYesNo, "CLP Spreadsheet Solver")
            If reply = vbNo Then
                ThisWorkbook.Worksheets("3.Solution").Activate
                Exit Sub
            Else
                Application.DisplayAlerts = False
            
                ThisWorkbook.Worksheets("3.Solution").Delete
                ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = "3.Solution"
                
                Application.DisplayAlerts = True
            End If
        End If
    Else
        reply = MsgBox("This will take some time (estimated " & estimatedTime & " seconds). Do you want to continue?", vbYesNo, "CLP Spreadsheet Solver")
        If reply = vbNo Then
            Exit Sub
        End If
        ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = "3.Solution"
    End If
    
    ThisWorkbook.Worksheets("3.Solution").Activate
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    'ActiveSheet.EnableFormatConditionsCalculation = False
    
    Dim container_count As Long
    Dim formulaText As String
    Dim objectiveFormula As String
    Dim temp_upper_bound As Double
    Dim combinedRange As Range
    
    Dim formula_delimiter As String
    If InStr(CStr(WorksheetFunction.Pi), ".") > 0 Then
        formula_delimiter = ","
    Else
        formula_delimiter = ";"
    End If
    
    objectiveFormula = "="
    Cells(1, 1).Value = "Total net profit:"
    Cells(1, 1).Font.Bold = True
    Cells(1, 2).Font.ColorIndex = 1
    Cells(1, 2).Interior.ColorIndex = 36
    
    Dim offset As Long
    offset = 0
    
    For i = 1 To num_container_types

        container_count = ThisWorkbook.Worksheets("2.Containers").Cells(1 + i, 10).Value
        
        For j = 1 To container_count
            
            Application.StatusBar = "Setting up container type " & i & ", container " & j
            Cells(5, 1 + offset).Value = "Item count"
            Cells(5, 2 + offset).Value = "Item type name"
            Cells(4, 3 + offset).Value = "Origin corner"
            Cells(5, 3 + offset).Value = "x coordinate"
            Cells(5, 4 + offset).Value = "y coordinate"
            Cells(5, 5 + offset).Value = "z coordinate"
            Cells(5, 6 + offset).Value = "Orientation"
            Cells(5, 7 + offset).Value = "Item type ID"
            Cells(5, 8 + offset).Value = "Volume"
            Cells(5, 9 + offset).Value = "Weight"
            Cells(5, 10 + offset).Value = "Profit"
            
            Range(Cells(4, 3 + offset), Cells(4, 5 + offset)).Select
            With Selection
                .HorizontalAlignment = xlCenter
                .Merge
            End With
            
            Range(Cells(6, 2 + offset), Cells(5 + num_items, 2 + offset)).Select

            With Selection
                .Font.ColorIndex = 1
                .Interior.ColorIndex = 33
            End With
            
            With Selection.Validation
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="='1.Items'!" & Cells(3, 2).Address & ":" & Cells(2 + num_item_types, 2).Address
                .IgnoreBlank = True
                .InCellDropdown = True
                .ErrorTitle = "Warning"
                .ErrorMessage = "Please select a value from the list available in the selected cell."
                .ShowError = True
            End With
            
            Range(Cells(6, 3 + offset), Cells(5 + num_items, 6 + offset)).Select

            With Selection
                .Font.ColorIndex = 1
                .Interior.ColorIndex = 33
                .FormatConditions.Add Type:=xlExpression, Formula1:="=ISBLANK(" & Cells(6, 2 + offset).Address(False, True) & ")"
                .FormatConditions(1).Font.ColorIndex = 33
            End With
            
            'temp_upper_bound = ThisWorkbook.Worksheets("2.Containers").Cells(1 + i, 3).Value
            Range(Cells(6, 3 + offset), Cells(5 + num_items, 3 + offset)).Select

            With Selection
            
                .Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:=-epsilon
                .Validation.ErrorMessage = "Please enter a nonnegative value"
                
'                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:=temp_upper_bound
'                .FormatConditions(1).Interior.Pattern = xlNone
'                .FormatConditions(1).Interior.ColorIndex = 3
'                .FormatConditions(1).StopIfTrue = True
                
            End With

            'temp_upper_bound = ThisWorkbook.Worksheets("2.Containers").Cells(1 + i, 4).Value
            Range(Cells(6, 4 + offset), Cells(5 + num_items, 4 + offset)).Select

            With Selection
            
                .Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:=-epsilon
                .Validation.ErrorMessage = "Please enter a nonnegative value"
                
'                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:=temp_upper_bound
'                .FormatConditions(1).Interior.Pattern = xlNone
'                .FormatConditions(1).Interior.ColorIndex = 3
'                .FormatConditions(1).StopIfTrue = True
                
            End With
            
            Range(Cells(6, 6 + offset), Cells(5 + num_items, 6 + offset)).Select

            With Selection
                .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="xyz, zyx, xzy, yzx, yxz, zxy"
                .Validation.IgnoreBlank = True
                .Validation.InCellDropdown = True
                
                .Value = "xyz"
            End With
            
            'temp_upper_bound = ThisWorkbook.Worksheets("2.Containers").Cells(1 + i, 6).Value
            Range(Cells(6, 8 + offset), Cells(5 + num_items, 8 + offset)).Select
            
            With Selection
                
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="="" """
                .FormatConditions(1).Interior.Pattern = xlNone
                .FormatConditions(1).Interior.ColorIndex = 36
                .FormatConditions(1).StopIfTrue = True
    
'                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:=temp_upper_bound
'                .FormatConditions(2).Interior.Pattern = xlNone
'                .FormatConditions(2).Interior.ColorIndex = 3
'                .FormatConditions(2).StopIfTrue = True
            
            End With
            
            'temp_upper_bound = ThisWorkbook.Worksheets("2.Containers").Cells(1 + i, 7).Value
            Range(Cells(6, 9 + offset), Cells(5 + num_items, 9 + offset)).Select
            
            With Selection
                            
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="="" """
                .FormatConditions(1).Interior.Pattern = xlNone
                .FormatConditions(1).Interior.ColorIndex = 36
                .FormatConditions(1).StopIfTrue = True
    
'                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:=temp_upper_bound
'                .FormatConditions(2).Interior.Pattern = xlNone
'                .FormatConditions(2).Interior.ColorIndex = 3
'                .FormatConditions(2).StopIfTrue = True
                
            End With
            
            For k = 1 To num_items
                
                Cells(5 + k, 1 + offset).Value = k
                Cells(5 + k, 1 + offset).Font.ColorIndex = 2
                Cells(5 + k, 1 + offset).Interior.ColorIndex = 1
                
                Range(Cells(5 + k, 7 + offset), Cells(5 + k, 10 + offset)).Select
                Selection.Font.ColorIndex = 1
                Selection.Interior.ColorIndex = 36
                
                Cells(5 + k, 3 + offset).Value = 0
                Cells(5 + k, 3 + offset).NumberFormat = "0.00"
                
                Cells(5 + k, 4 + offset).Value = 0
                Cells(5 + k, 4 + offset).NumberFormat = "0.00"
                
                Cells(5 + k, 5 + offset).Value = 0
                Cells(5 + k, 5 + offset).NumberFormat = "0.00"
                          
                formulaText = "=IFERROR(MATCH(" & "INDIRECT(" & Chr(34) & Cells(5 + k, 2 + offset).Address(False, False) & Chr(34) & ")"
                formulaText = formulaText & ",'1.Items'!" & Cells(3, 2).Address & ":" & Cells(2 + num_item_types, 2).Address & ", 0),"
                formulaText = formulaText & Chr(34) & " " & Chr(34) & ")"
                
                Cells(5 + k, 7 + offset).Formula = formulaText

                formulaText = "=IFERROR(IF(ISNUMBER(" & Cells(5 + k, 7 + offset).Address(False, False) & "),"
                formulaText = formulaText & "INDEX('1.Items'!" & Cells(3, 7).Address(False, False) & ":" & Cells(2 + num_items, 7).Address(False, False)
                formulaText = formulaText & "," & Cells(5 + k, 7 + offset).Address(False, False) & " ),"
                formulaText = formulaText & Chr(34) & " " & Chr(34) & ")," & Chr(34) & " " & Chr(34) & ")"

                Cells(5 + k, 8 + offset).Formula = formulaText
                Cells(5 + k, 8 + offset).NumberFormat = "0.00"

                formulaText = "=IFERROR(IF(ISNUMBER(" & Cells(5 + k, 7 + offset).Address(False, False) & "),"
                formulaText = formulaText & "INDEX('1.Items'!" & Cells(3, 11).Address(False, False) & ":" & Cells(2 + num_items, 11).Address(False, False)
                formulaText = formulaText & "," & Cells(5 + k, 7 + offset).Address(False, False) & " ),"
                formulaText = formulaText & Chr(34) & " " & Chr(34) & ")," & Chr(34) & " " & Chr(34) & ")"

                Cells(5 + k, 9 + offset).Formula = formulaText
                Cells(5 + k, 9 + offset).NumberFormat = "0.00"
                
                formulaText = "=IFERROR(IF(ISNUMBER(" & Cells(5 + k, 7 + offset).Address(False, False) & "),"
                formulaText = formulaText & "INDEX('1.Items'!" & Cells(3, 15).Address(False, False) & ":" & Cells(2 + num_items, 15).Address(False, False)
                formulaText = formulaText & "," & Cells(5 + k, 7 + offset).Address(False, False) & " ),"
                formulaText = formulaText & Chr(34) & " " & Chr(34) & ")," & Chr(34) & " " & Chr(34) & ")"

                Cells(5 + k, 10 + offset).Formula = formulaText
                Cells(5 + k, 10 + offset).NumberFormat = "0.00"

                Range(Cells(5 + k, 1 + offset), Cells(5 + k, 10 + offset)).Select
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 16
                    .TintAndShade = 0
                    .weight = xlThin
                End With
                
            Next k
            
            Cells(3, 1 + offset).Value = "Container " & j & " of " & ThisWorkbook.Worksheets("2.Containers").Cells(1 + i, 2).Value
            
            Cells(3, 7 + offset).Value = "Item count"

            formulaText = "=COUNT(" & Cells(6, 7 + offset).Address & ":" & Cells(5 + num_items, 7 + offset).Address & ")"
        
            Cells(4, 7 + offset).Formula = formulaText
            Cells(4, 7 + offset).Interior.ColorIndex = 36
            
            Cells(3, 8 + offset).Value = "Total volume"

            formulaText = "=SUM(" & Cells(6, 8 + offset).Address & ":" & Cells(5 + num_items, 8 + offset).Address & ")"
        
            Cells(4, 8 + offset).Formula = formulaText
            Cells(4, 8 + offset).Interior.ColorIndex = 36
            Cells(4, 8 + offset).NumberFormat = "0.00"
            
            temp_upper_bound = ThisWorkbook.Worksheets("2.Containers").Cells(1 + i, 6).Value

            Cells(4, 8 + offset).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:=temp_upper_bound
            Cells(4, 8 + offset).FormatConditions(1).Interior.Pattern = xlNone
            Cells(4, 8 + offset).FormatConditions(1).Interior.ColorIndex = 3
            Cells(4, 8 + offset).FormatConditions(1).StopIfTrue = True
            
            Cells(3, 9 + offset).Value = "Total weight"

            formulaText = "=SUM(" & Cells(6, 9 + offset).Address & ":" & Cells(5 + num_items, 9 + offset).Address & ")"
        
            Cells(4, 9 + offset).Formula = formulaText
            Cells(4, 9 + offset).Interior.ColorIndex = 36
            Cells(4, 9 + offset).NumberFormat = "0.00"
            
            temp_upper_bound = ThisWorkbook.Worksheets("2.Containers").Cells(1 + i, 7).Value

            Cells(4, 9 + offset).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:=temp_upper_bound
            Cells(4, 9 + offset).FormatConditions(1).Interior.Pattern = xlNone
            Cells(4, 9 + offset).FormatConditions(1).Interior.ColorIndex = 3
            Cells(4, 9 + offset).FormatConditions(1).StopIfTrue = True

            Cells(3, 10 + offset).Value = "Net profit"

            formulaText = "=SUM(" & Cells(6, 10 + offset).Address & ":" & Cells(5 + num_items, 10 + offset).Address & ")"
            formulaText = formulaText & "+ IF(COUNTA(" & "'3.Solution'!" & Cells(6, 2 + offset).Address & ":" & Cells(5 + num_items, 2 + offset).Address & "),"
            formulaText = formulaText & "-INDEX('2.Containers'!" & Cells(2, 9).Address & ":" & Cells(1 + num_container_types, 9).Address & "," & i & "), 0)"
            
            Cells(4, 10 + offset).Formula = formulaText
            Cells(4, 10 + offset).Interior.ColorIndex = 36
            Cells(4, 10 + offset).NumberFormat = "0.00"
                
            objectiveFormula = objectiveFormula & "+" & Cells(4, 10 + offset).Address
            
            offset = offset + column_offset
        Next j
    Next i
           
    Application.StatusBar = "Setting up unpacked item list"
    
    Cells(5, 1 + offset).Value = "Item count"
    Cells(5, 2 + offset).Value = "Item type name"
    Cells(5, 3 + offset).Value = "Item type ID"
    Cells(5, 4 + offset).Value = "Volume"
    Cells(5, 5 + offset).Value = "Weight"
    Cells(5, 6 + offset).Value = "Profit"
    
    Range(Cells(6, 2 + offset), Cells(5 + num_items, 2 + offset)).Select

    With Selection
        .Font.ColorIndex = 1
        .Interior.ColorIndex = 33
    End With
    
    With Selection.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="='1.Items'!" & Cells(3, 2).Address & ":" & Cells(2 + num_item_types, 2).Address
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Warning"
        .ErrorMessage = "Please select a value from the list available in the selected cell."
        .ShowError = True
    End With
    
    For k = 1 To num_items
        
        Cells(5 + k, 1 + offset).Value = k
        Cells(5 + k, 1 + offset).Font.ColorIndex = 2
        Cells(5 + k, 1 + offset).Interior.ColorIndex = 15
        
        Range(Cells(5 + k, 3 + offset), Cells(5 + k, 6 + offset)).Select
        Selection.Font.ColorIndex = 1
        Selection.Interior.ColorIndex = 36
                  
        formulaText = "=IFERROR(MATCH(" & "INDIRECT(" & Chr(34) & Cells(5 + k, 2 + offset).Address(False, False) & Chr(34) & ")"
        formulaText = formulaText & ",'1.Items'!" & Cells(3, 2).Address & ":" & Cells(2 + num_item_types, 2).Address & ", 0),"
        formulaText = formulaText & Chr(34) & " " & Chr(34) & ")"
        
        Cells(5 + k, 3 + offset).Formula = formulaText

        formulaText = "=IFERROR(IF(ISNUMBER(" & Cells(5 + k, 3 + offset).Address(False, False) & "),"
        formulaText = formulaText & "INDEX('1.Items'!" & Cells(3, 7).Address(False, False) & ":" & Cells(2 + num_items, 7).Address(False, False)
        formulaText = formulaText & "," & Cells(5 + k, 3 + offset).Address(False, False) & " ),"
        formulaText = formulaText & Chr(34) & " " & Chr(34) & ")," & Chr(34) & " " & Chr(34) & ")"

        Cells(5 + k, 4 + offset).Formula = formulaText
        Cells(5 + k, 4 + offset).NumberFormat = "0.00"

        formulaText = "=IFERROR(IF(ISNUMBER(" & Cells(5 + k, 3 + offset).Address(False, False) & "),"
        formulaText = formulaText & "INDEX('1.Items'!" & Cells(3, 11).Address(False, False) & ":" & Cells(2 + num_items, 11).Address(False, False)
        formulaText = formulaText & "," & Cells(5 + k, 3 + offset).Address(False, False) & " ),"
        formulaText = formulaText & Chr(34) & " " & Chr(34) & ")," & Chr(34) & " " & Chr(34) & ")"

        Cells(5 + k, 5 + offset).Formula = formulaText
        Cells(5 + k, 5 + offset).NumberFormat = "0.00"
        
        formulaText = "=IFERROR(IF(ISNUMBER(" & Cells(5 + k, 3 + offset).Address(False, False) & "),"
        formulaText = formulaText & "INDEX('1.Items'!" & Cells(3, 15).Address(False, False) & ":" & Cells(2 + num_items, 15).Address(False, False)
        formulaText = formulaText & "," & Cells(5 + k, 3 + offset).Address(False, False) & " ),"
        formulaText = formulaText & Chr(34) & " " & Chr(34) & ")," & Chr(34) & " " & Chr(34) & ")"

        Cells(5 + k, 6 + offset).Formula = formulaText
        Cells(5 + k, 6 + offset).NumberFormat = "0.00"

        Range(Cells(5 + k, 1 + offset), Cells(5 + k, 6 + offset)).Select
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 16
            .TintAndShade = 0
            .weight = xlThin
        End With
        
    Next k
    
    Cells(3, 1 + offset).Value = "Unpacked items"
    
    Cells(3, 2 + offset).Value = "Item count"

    formulaText = "=COUNT(" & Cells(6, 3 + offset).Address & ":" & Cells(5 + num_items, 3 + offset).Address & ")"

    Cells(4, 2 + offset).Formula = formulaText
    Cells(4, 2 + offset).Interior.ColorIndex = 36
    
    Cells(3, 4 + offset).Value = "Total volume"

    formulaText = "=SUM(" & Cells(6, 4 + offset).Address & ":" & Cells(5 + num_items, 4 + offset).Address & ")"

    Cells(4, 4 + offset).Formula = formulaText
    Cells(4, 4 + offset).Interior.ColorIndex = 36
    Cells(4, 4 + offset).NumberFormat = "0.00"
    
    Cells(3, 5 + offset).Value = "Total weight"

    formulaText = "=SUM(" & Cells(6, 5 + offset).Address & ":" & Cells(5 + num_items, 5 + offset).Address & ")"

    Cells(4, 5 + offset).Formula = formulaText
    Cells(4, 5 + offset).Interior.ColorIndex = 36
    Cells(4, 5 + offset).NumberFormat = "0.00"

    Cells(3, 6 + offset).Value = "Missed profit"

    formulaText = "=SUM(" & Cells(6, 6 + offset).Address & ":" & Cells(5 + num_items, 6 + offset).Address & ")"
    
    Cells(4, 6 + offset).Formula = formulaText
    Cells(4, 6 + offset).Interior.ColorIndex = 36
    Cells(4, 6 + offset).NumberFormat = "0.00"
           
    Cells(1, 2).Value = objectiveFormula
    
    Cells(1, 2).FormatConditions.Add Type:=xlExpression, Formula1:="=ISERROR(" & Cells(1, 2).Address & ")"
    Cells(1, 2).FormatConditions(1).Interior.ColorIndex = 3
    
    Rows("1:5").Select
    Selection.Font.Bold = True
    
    Columns.AutoFit
    
    Cells(num_items + 7, 1).Value = "List of detected infeasibilities"
    Cells(num_items + 7, 1).Font.Bold = True
    
    offset = 0

    For i = 1 To num_container_types

        container_count = ThisWorkbook.Worksheets("2.Containers").Cells(1 + i, 10).Value
        
        For j = 1 To container_count

            Columns(7 + offset).EntireColumn.Hidden = True

            offset = offset + column_offset
        Next j
    Next i
    
    Columns(3 + offset).EntireColumn.Hidden = True
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    'ActiveSheet.EnableFormatConditionsCalculation = True
    
    Cells(1, 1).Select
    
    MsgBox ("Solution worksheet setup complete.")
    
End Sub
Sub SetupVisualizationWorksheet()

    Dim WorksheetExists As Boolean
    Dim MissingCoordinates As Boolean
    Dim reply As Integer
    
    WorksheetExists = CheckWorksheetExistence("1.Items") And CheckWorksheetExistence("2.Containers") And CheckWorksheetExistence("3.Solution")
    If WorksheetExists = False Then
        MsgBox "Worksheets 1.Items, 2.Containers, and 3.Solution must exist for the visualization worksheet to be setup."
        Exit Sub
    End If
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim num_item_types As Long
    num_item_types = ThisWorkbook.Worksheets("CLP Solver Console").Cells(2, 3).Value
    
    Dim num_items As Long
    
    num_items = 0
    For i = 1 To num_item_types
        num_items = num_items + ThisWorkbook.Worksheets("1.Items").Cells(2 + i, 16).Value
    Next i
    
    If num_items = 0 Then
        MsgBox "There must be at least one item for the solution worksheet to be setup."
        Exit Sub
    End If
    
    Dim num_container_types As Long
    num_container_types = ThisWorkbook.Worksheets("CLP Solver Console").Cells(4, 3).Value
    
    Dim num_containers As Long
    
    num_containers = 0
    For i = 1 To num_container_types
        num_containers = num_containers + ThisWorkbook.Worksheets("2.Containers").Cells(1 + i, 10).Value
    Next i
    
    If num_containers = 0 Then
        MsgBox "There must be at least one container for the solution worksheet to be setup."
        Exit Sub
    End If
    
    WorksheetExists = CheckWorksheetExistence("4.Visualization")
    If WorksheetExists = True Then
        
        reply = MsgBox("This will overwrite existing visualization data. Do you want to continue?", vbYesNo, "CLP Spreadsheet Solver")

        If reply = vbNo Then
            ThisWorkbook.Worksheets("4.Visualization").Activate
            Exit Sub
        Else
            Application.DisplayAlerts = False
            ThisWorkbook.Worksheets("4.Visualization").Delete
            Application.DisplayAlerts = True
            ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = "4.Visualization"
        End If
    Else
        ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = "4.Visualization"
    End If
    
    Call RefreshVisualizationWorksheet
End Sub
Sub RefreshVisualizationWorksheet()
    
    Dim WorksheetExists As Boolean
    
    WorksheetExists = CheckWorksheetExistence("4.Visualization")
    If WorksheetExists = False Then
        Exit Sub
    End If
    
    ThisWorkbook.Worksheets("4.Visualization").Activate
    
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
    
    Rows("1").Select
    Selection.Clear
    Cells(1, 1).Select
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim num_item_types As Long
    num_item_types = ThisWorkbook.Worksheets("CLP Solver Console").Cells(2, 3).Value
    
    Dim num_items As Long
    
    num_items = 0
    For i = 1 To num_item_types
        num_items = num_items + ThisWorkbook.Worksheets("1.Items").Cells(2 + i, 16).Value
    Next i
   
    Dim num_container_types As Long
    num_container_types = ThisWorkbook.Worksheets("CLP Solver Console").Cells(4, 3).Value
    
    Dim num_containers As Long
    Dim container_count As Long
    
    num_containers = 0
    For i = 1 To num_container_types
        num_containers = num_containers + ThisWorkbook.Worksheets("2.Containers").Cells(1 + i, 10).Value
    Next i
    
    Dim item_label_option As String
    item_label_option = ThisWorkbook.Worksheets("CLP Solver Console").Cells(8, 3).Value
    
    Dim container_label_option As String
    container_label_option = ThisWorkbook.Worksheets("CLP Solver Console").Cells(9, 3).Value

    Application.ScreenUpdating = False
            
    'scaling - determine the maximum width and height of the containers
    
    Dim focus_shape As Long
    focus_shape = 0
    
    Dim row_index As Long
    Dim file_name As String
    
    Dim v_scale As Double
    Dim offset As Long
    Dim x_offset As Long
    Dim y_offset As Double
    Dim item_x_offset As Double
    Dim item_y_offset As Double
    
    Dim nw_x As Double
    Dim nw_y As Double
    
    Dim container_width As Double
    Dim container_height As Double
    Dim container_length As Double
    Dim max_x As Double
    Dim max_y As Double
    
    Dim rotation As String
    
    Dim original_item_width As Double
    Dim original_item_height As Double
    Dim original_item_length As Double
    
    Dim rotated_item_width As Double
    Dim rotated_item_height As Double
    Dim rotated_item_length As Double
    
    Dim max_item_dimension As Double
    
    Dim swap_double As Double
    
    Dim row_height As Double
    row_height = Rows("1").height
        
    Dim column_width As Double
    column_width = Columns("Z").width
    
    max_x = ThisWorkbook.Sheets("2.Containers").Cells(2, 3) + (displacement_multiplier * ThisWorkbook.Sheets("2.Containers").Cells(2, 5))
    max_y = ThisWorkbook.Sheets("2.Containers").Cells(2, 4) + (displacement_multiplier * ThisWorkbook.Sheets("2.Containers").Cells(2, 5))
    
    For i = 1 To num_container_types
        
        If max_x < ThisWorkbook.Sheets("2.Containers").Cells(1 + i, 3) + (displacement_multiplier * ThisWorkbook.Sheets("2.Containers").Cells(1 + i, 5)) Then
            max_x = ThisWorkbook.Sheets("2.Containers").Cells(1 + i, 3) + (displacement_multiplier * ThisWorkbook.Sheets("2.Containers").Cells(1 + i, 5))
        End If
               
        If max_y < ThisWorkbook.Sheets("2.Containers").Cells(1 + i, 4) + (displacement_multiplier * ThisWorkbook.Sheets("2.Containers").Cells(1 + i, 5)) Then
            max_y = ThisWorkbook.Sheets("2.Containers").Cells(1 + i, 4) + (displacement_multiplier * ThisWorkbook.Sheets("2.Containers").Cells(1 + i, 5))
        End If
    Next i
    
    If max_x >= max_y Then
        v_scale = width_limit / max_x
    Else
        v_scale = width_limit / max_y
    End If
    
    'v_scale = v_scale * 2
    'ActiveWindow.Zoom = 50
    'MsgBox v_scale
    
    offset = 0
    
    x_offset = 1
    y_offset = 2 * row_height
    
    For i = 1 To num_container_types
    
        container_count = ThisWorkbook.Worksheets("2.Containers").Cells(1 + i, 10).Value
        
        For j = 1 To container_count
            
            container_width = v_scale * ThisWorkbook.Worksheets("2.Containers").Cells(1 + i, 3).Value
            container_height = v_scale * ThisWorkbook.Worksheets("2.Containers").Cells(1 + i, 4).Value
            container_length = v_scale * ThisWorkbook.Worksheets("2.Containers").Cells(1 + i, 5).Value
            
            nw_x = column_width * x_offset
            nw_y = y_offset + v_scale * max_y - container_height
            
            Cells(1, 1 + x_offset).Value = ThisWorkbook.Worksheets("3.Solution").Cells(3, 1 + offset).Value
            Cells(1, 1 + x_offset).Font.Bold = msoTrue
            
            ActiveSheet.Shapes.AddShape(msoShapeRectangle, nw_x, nw_y, container_width, container_height).Select 'type, left, top, width, height
            ActiveSheet.Shapes(ActiveSheet.Shapes.Count).Fill.Transparency = 1
            ActiveSheet.Shapes(ActiveSheet.Shapes.Count).Name = "Container"
            Selection.Interior.ColorIndex = 50
            
            
            With Selection.ShapeRange.ThreeD
                .Visible = True
                .depth = container_length
                .ContourWidth = 0.2
                .ExtrusionColor.RGB = RGB(255, 255, 255) 'for containers
                .SetExtrusionDirection msoExtrusionTopRight
                .PresetLightingDirection = msoLightingTopRight
                .PresetMaterial = msoMaterialTranslucentPowder 'for containers
            End With
            
            If container_label_option = "Yes" Then
                Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = ThisWorkbook.Worksheets("3.Solution").Cells(3, 1 + offset).Value
            End If
            Selection.ShapeRange.TextFrame2.TextRange.Font.Bold = msoTrue
            Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
            Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            
            l = ThisWorkbook.Worksheets("3.Solution").Cells(4, 7 + offset).Value

            If (l > 0) And (focus_shape = 0) Then
                focus_shape = ActiveSheet.Shapes.Count
            End If
            
            
            For k = 1 To l

                If ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 2).Value <> "" Then

                    rotation = ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 6).Text
                    row_index = ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 7).Value + 2

                    original_item_width = v_scale * ThisWorkbook.Worksheets("1.Items").Cells(row_index, 4)
                    original_item_height = v_scale * ThisWorkbook.Worksheets("1.Items").Cells(row_index, 5)
                    original_item_length = v_scale * ThisWorkbook.Worksheets("1.Items").Cells(row_index, 6)

                    ' rotate it
                    
                    If rotation = "xyz" Then
                        rotated_item_width = original_item_width
                        rotated_item_height = original_item_height
                        rotated_item_length = original_item_length
                    End If
                    
                    If rotation = "zyx" Then
                        rotated_item_width = original_item_length
                        rotated_item_height = original_item_height
                        rotated_item_length = original_item_width
                    End If
                    
                    If rotation = "xzy" Then
                        rotated_item_width = original_item_width
                        rotated_item_height = original_item_length
                        rotated_item_length = original_item_height
                    End If
                    
                    If rotation = "yzx" Then
                        rotated_item_width = original_item_height
                        rotated_item_height = original_item_length
                        rotated_item_length = original_item_width
                    End If
                    
                    If rotation = "yxz" Then
                        rotated_item_width = original_item_height
                        rotated_item_height = original_item_width
                        rotated_item_length = original_item_length
                    End If
                    
                    If rotation = "zxy" Then
                        rotated_item_width = original_item_length
                        rotated_item_height = original_item_width
                        rotated_item_length = original_item_height
                    End If
                    
                    max_item_dimension = rotated_item_width + (displacement_multiplier * rotated_item_length)
                    If max_item_dimension < original_item_height + (displacement_multiplier * rotated_item_length) Then
                        max_item_dimension = original_item_height + (displacement_multiplier * rotated_item_length)
                    End If

                    'create it at the max_item_dimension, max_item_dimension

                    ActiveSheet.Shapes.AddShape(msoShapeRectangle, max_item_dimension, max_item_dimension, rotated_item_width, rotated_item_height).Select
                    ActiveSheet.Shapes(ActiveSheet.Shapes.Count).Name = "Item"

                    If item_label_option = "Yes" Then
                        Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 2).Value
                        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                        'Selection.ShapeRange.TextFrame2.WordWrap = msoFalse
                    End If

                    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
                    Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

                    file_name = ThisWorkbook.Worksheets("1.Items").Cells(row_index, 3)

                    If file_name <> "" Then

                        #If Mac Then
                            file_name = Application.ThisWorkbook.Path & ":" & ThisWorkbook.Worksheets("1.Items").Cells(row_index, 3)
                        #Else
                            file_name = Application.ThisWorkbook.Path & "\" & ThisWorkbook.Worksheets("1.Items").Cells(row_index, 3)
                        #End If

                        If Dir(file_name) <> "" Then
                            With Selection.ShapeRange.Fill
                                .UserPicture file_name
                            End With
                        End If

                    Else
                        Selection.Interior.Color = ThisWorkbook.Worksheets("1.Items").Cells(row_index, 3).Interior.Color
                    End If

                    With Selection.ShapeRange.ThreeD
                        .Visible = True
                        .depth = rotated_item_length
                        .ContourWidth = 0
                        .ExtrusionColor.RGB = RGB(255, 255, 255) 'for items
                        .SetExtrusionDirection msoExtrusionTopRight
                        .PresetLightingDirection = msoLightingTopRight
                        .PresetMaterial = msoMaterialClear 'for items
                    End With
                    
                    'carry it to the sw corner

                    Selection.ShapeRange.IncrementLeft nw_x - max_item_dimension
                    Selection.ShapeRange.IncrementTop nw_y + container_height - rotated_item_height - max_item_dimension
                    
                    'carry it to the origin
                    
                    Selection.ShapeRange.IncrementLeft displacement_multiplier * (container_length - rotated_item_length)
                    Selection.ShapeRange.IncrementTop -displacement_multiplier * (container_length - rotated_item_length)

                    'carry it to its place

                    Selection.ShapeRange.IncrementLeft (v_scale * ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 3).Value) - v_scale * displacement_multiplier * ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 5).Value
                    Selection.ShapeRange.IncrementTop -(v_scale * ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 4).Value) + v_scale * displacement_multiplier * ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 5).Value

                Else
                    Exit For
                End If

            Next k
            
            offset = offset + column_offset
            
            x_offset = x_offset + Int(v_scale * max_x / column_width) + 2
            
        Next j
        
    Next i
    
    'visualize unpacked items
    
    Cells(1, 1 + x_offset).Value = "Unpacked items"
    Cells(1, 1 + x_offset).Font.Bold = True
    
    item_x_offset = x_offset * column_width

    l = ThisWorkbook.Worksheets("3.Solution").Cells(4, 2 + offset).Value

    For k = 1 To l

        If ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 2).Value <> "" Then

            row_index = ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 3).Value + 2

            original_item_width = v_scale * ThisWorkbook.Worksheets("1.Items").Cells(row_index, 4)
            original_item_height = v_scale * ThisWorkbook.Worksheets("1.Items").Cells(row_index, 5)
            original_item_length = v_scale * ThisWorkbook.Worksheets("1.Items").Cells(row_index, 6)

            rotated_item_width = original_item_width
            rotated_item_height = original_item_height
            rotated_item_length = original_item_length

            max_item_dimension = rotated_item_width + (displacement_multiplier * rotated_item_length)
            If max_item_dimension < original_item_height + (displacement_multiplier * rotated_item_length) Then
                max_item_dimension = original_item_height + (displacement_multiplier * rotated_item_length)
            End If

            'create it at the max_item_dimension, max_item_dimension

            ActiveSheet.Shapes.AddShape(msoShapeRectangle, max_item_dimension, max_item_dimension, rotated_item_width, rotated_item_height).Select

            If item_label_option = "Yes" Then
                Selection.ShapeRange.TextFrame2.TextRange.Characters.Text = ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 2).Value
                Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            End If

            Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
            Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

            file_name = ThisWorkbook.Worksheets("1.Items").Cells(row_index, 3)

            If file_name <> "" Then

                #If Mac Then
                    file_name = Application.ThisWorkbook.Path & ":" & ThisWorkbook.Worksheets("1.Items").Cells(row_index, 3)
                #Else
                    file_name = Application.ThisWorkbook.Path & "\" & ThisWorkbook.Worksheets("1.Items").Cells(row_index, 3)
                #End If

                If Dir(file_name) <> "" Then
                    With Selection.ShapeRange.Fill
                        .UserPicture file_name
                    End With
                End If

            Else
                Selection.Interior.Color = ThisWorkbook.Worksheets("1.Items").Cells(row_index, 3).Interior.Color
            End If

            With Selection.ShapeRange.ThreeD
                .Visible = True
                .depth = rotated_item_length
                .ContourWidth = 0
                .ExtrusionColor.RGB = RGB(255, 255, 255) 'for items
                .SetExtrusionDirection msoExtrusionTopRight
                .PresetLightingDirection = msoLightingTopRight
                .PresetMaterial = msoMaterialClear 'for items
            End With

            'carry it to the sw corner

            Selection.ShapeRange.IncrementLeft item_x_offset - max_item_dimension
            Selection.ShapeRange.IncrementTop height_limit - max_item_dimension

            item_x_offset = item_x_offset + rotated_item_width + 5
        Else
            Exit For
        End If

    Next k
            
    Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    
End Sub


Sub AnimateVisualizationWorksheet()

    Application.EnableCancelKey = xlErrorHandler
    On Error Resume Next
    
    Dim WorksheetExists As Boolean
    Dim reply As Integer
    
    WorksheetExists = CheckWorksheetExistence("1.Items") And CheckWorksheetExistence("2.Containers") And CheckWorksheetExistence("3.Solution") And CheckWorksheetExistence("4.Visualization")
    If WorksheetExists = False Then
        MsgBox "Worksheets 1.Items, 2.Containers, 3.Solution, and 4.Visualization must exist for the animation to start."
        Exit Sub
    End If
    
    Dim i As Long
    
    Dim max_row As Long
    max_row = (height_limit / Rows(1).RowHeight) + 2
    
    Dim offset As Long
    offset = 1
    Dim offset_increment As Long
    offset_increment = (width_limit / Columns(1).ColumnWidth) + 1
    
    Dim original_zoom As Long
    original_zoom = ActiveWindow.Zoom
    
    Dim selection_column_index As Long
    selection_column_index = ActiveCell.Column
    
    Dim num_item_types As Long
    num_item_types = ThisWorkbook.Worksheets("CLP Solver Console").Cells(2, 3).Value
    
    Dim num_items As Long
    
    num_items = 0
    For i = 1 To num_item_types
        num_items = num_items + ThisWorkbook.Worksheets("1.Items").Cells(2 + i, 16).Value
    Next i
    
    If num_items = 0 Then
        MsgBox "There must be at least one item for the animation to start."
        Exit Sub
    End If
    
    Dim num_container_types As Long
    num_container_types = ThisWorkbook.Worksheets("CLP Solver Console").Cells(4, 3).Value
    
    Dim num_containers As Long
    
    num_containers = 0
    For i = 1 To num_container_types
        num_containers = num_containers + ThisWorkbook.Worksheets("2.Containers").Cells(1 + i, 10).Value
    Next i
    
    If num_containers = 0 Then
        MsgBox "There must be at least one container for the animation to start."
        Exit Sub
    End If
    
    reply = MsgBox("This will take up to " & num_items + num_containers & " seconds. You can speed up the animation by pressing (or holding down) the <Esc> key. Do you want to continue?", vbYesNo, "CLP Spreadsheet Solver")
    If reply = vbNo Then
        Exit Sub
    End If
    
    Dim animation_advance_type As Long
    
    If ThisWorkbook.Worksheets("CLP Solver Console").Cells(10, 3).Value = "One second intervals" Then
        animation_advance_type = 0
    Else
        animation_advance_type = 1
    End If
    
    ThisWorkbook.Worksheets("4.Visualization").Activate
    
    Cells(1, 1).Select
    
    ActiveSheet.Shapes.SelectAll

    Selection.Visible = False

    DoEvents

    Application.Wait (Now() + TimeValue("0:00:01"))

    For i = 1 To ActiveSheet.Shapes.Count
    
        If ActiveSheet.Shapes(i).Name = "Container" Then
            Range(Cells(1, offset), Cells(max_row, offset + column_offset)).Select
            ActiveWindow.Zoom = True
            offset = offset + offset_increment
        End If
        
        ActiveSheet.Shapes(i).Visible = True
        DoEvents
        Cells(1, ActiveSheet.Shapes(i).BottomRightCell.Column).Select
        DoEvents
        If animation_advance_type = 0 Then
            Application.Wait (Now() + TimeValue("0:00:01"))
            DoEvents
        ElseIf i < ActiveSheet.Shapes.Count Then
            reply = MsgBox("Next item?", vbOKOnly, "CLP Spreadsheet Solver")
        End If
    Next i
    
    ThisWorkbook.Worksheets("4.Visualization").Activate
    
    Cells(1, selection_column_index).Select
    
    ActiveWindow.Zoom = original_zoom
    
    reply = MsgBox("Animation finished.         ", vbOKOnly, "CLP Spreadsheet Solver")
    
End Sub


Private Sub About()
    Dim reply As Integer
    
    reply = MsgBox("CLP Spreadsheet Solver v1.2 (Release 2)" & Chr(13) & "Open source, developed by Dr Gunes Erdogan, 2019 (G.Erdogan@bath.ac.uk), School of Management, University of Bath." & Chr(13) _
& Chr(13) & "The latest version of the solver can be downloaded at: http://people.bath.ac.uk/ge277/index.php/clp-spreadsheet-solver/" & Chr(13) & Chr(13) & "DISCLAIMER: THIS SOFTWARE IS PROVIDED 'AS IS' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE ORIGINAL DEVELOPER BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,  WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.", vbOKOnly, "CLP Spreadsheet Solver")
End Sub
Private Sub SendFeedback()
    ThisWorkbook.FollowHyperlink "mailto:G.Erdogan@bath.ac.uk&subject=CLP Spreadsheet Solver"
End Sub


Private Sub ResetWorkbook()

    Dim WorksheetExists As Boolean
    Dim reply As Integer
    
    reply = MsgBox("This will delete all existing problem data. Do you want to continue?", vbYesNo, "CLP Spreadsheet Solver")
    If reply = vbNo Then
        Exit Sub
    Else
    
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        
        WorksheetExists = CheckWorksheetExistence("4.Visualization")
        If WorksheetExists = True Then
            ThisWorkbook.Worksheets("4.Visualization").Delete
            DoEvents
        End If

        WorksheetExists = CheckWorksheetExistence("3.Solution")
        If WorksheetExists = True Then
            ThisWorkbook.Worksheets("3.Solution").Delete
            DoEvents
        End If
        
        WorksheetExists = CheckWorksheetExistence("2.3.Container-ItemCompatibility")
        If WorksheetExists = True Then
            ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Delete
            DoEvents
        End If
        
        WorksheetExists = CheckWorksheetExistence("2.Containers")
        If WorksheetExists = True Then
            ThisWorkbook.Worksheets("2.Containers").Delete
            DoEvents
        End If
        
        WorksheetExists = CheckWorksheetExistence("1.3.Item-Item Compatibility")
        If WorksheetExists = True Then
            ThisWorkbook.Worksheets("1.3.Item-Item Compatibility").Delete
            DoEvents
        End If
        
        WorksheetExists = CheckWorksheetExistence("1.Items")
        If WorksheetExists = True Then
            ThisWorkbook.Worksheets("1.Items").Delete
            DoEvents
        End If
    
        WorksheetExists = CheckWorksheetExistence("CLP Solver Console")
        If WorksheetExists = False Then
            Call SetupConsoleWorksheet
        End If
        
        Application.DisplayAlerts = True
    End If

    WorksheetExists = CheckWorksheetExistence("CLP Solver Console")
    If WorksheetExists = False Then
        Call SetupConsoleWorksheet
    End If
    
    Application.EnableEvents = True
    ThisWorkbook.Worksheets("CLP Solver Console").Activate
End Sub

Sub SortItemTypes()
    
    Dim WorksheetExists As Boolean
    Dim reply As Integer
    
    Dim num_item_types As Long
    num_item_types = ThisWorkbook.Worksheets("CLP Solver Console").Cells(2, 3).Value
        
    WorksheetExists = CheckWorksheetExistence("1.Items")
    If WorksheetExists = False Then
        MsgBox "Worksheet 1.Items must exist for the items to be sorted."
        Exit Sub
    End If
    
    WorksheetExists = CheckWorksheetExistence("2.Containers") Or CheckWorksheetExistence("3.Solution") Or CheckWorksheetExistence("4.Visualization")
    If WorksheetExists = True Then
        reply = MsgBox("This will delete existing container, solution, and visualization data. Do you want to continue?", vbYesNo, "CLP Spreadsheet Solver")
        If reply = vbNo Then
            Exit Sub
        End If
    End If
    
    WorksheetExists = CheckWorksheetExistence("4.Visualization")
    If WorksheetExists = True Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("4.Visualization").Delete
        Application.DisplayAlerts = True
    End If
    ThisWorkbook.Worksheets("1.Items").Activate
    
    WorksheetExists = CheckWorksheetExistence("3.Solution")
    If WorksheetExists = True Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("3.Solution").Delete
        Application.DisplayAlerts = True
    End If
    
    WorksheetExists = CheckWorksheetExistence("2.3.Container-ItemCompatibility")
    If WorksheetExists = True Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Delete
        Application.DisplayAlerts = True
    End If
    
    WorksheetExists = CheckWorksheetExistence("2.Containers")
    If WorksheetExists = True Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("2.Containers").Delete
        Application.DisplayAlerts = True
    End If
    
    WorksheetExists = CheckWorksheetExistence("1.3.Item-Item Compatibility")
    If WorksheetExists = True Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("2.3.Item-Item Compatibility").Delete
        Application.DisplayAlerts = True
    End If
    
    Range(Cells(3, 2), Cells(num_item_types + 2, 16)).Select
    
    Selection.Sort Key1:=Columns("B"), Order1:=xlAscending, Header:=xlNo
    
End Sub

Sub SortContainerTypes()
    
    Dim WorksheetExists As Boolean
    Dim reply As Integer
    
    Dim num_container_types As Long
    num_container_types = ThisWorkbook.Worksheets("CLP Solver Console").Cells(4, 3).Value
        
    WorksheetExists = CheckWorksheetExistence("2.Containers")
    If WorksheetExists = False Then
        MsgBox "Worksheet 2.Containers must exist for the containers to be sorted."
        Exit Sub
    End If
    
    WorksheetExists = CheckWorksheetExistence("3.Solution") Or CheckWorksheetExistence("4.Visualization")
    If WorksheetExists = True Then
        reply = MsgBox("This will delete existing solution and visualization data. Do you want to continue?", vbYesNo, "CLP Spreadsheet Solver")
        If reply = vbNo Then
            Exit Sub
        End If
    End If
    
    WorksheetExists = CheckWorksheetExistence("4.Visualization")
    If WorksheetExists = True Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("4.Visualization").Delete
        Application.DisplayAlerts = True
    End If
    
    WorksheetExists = CheckWorksheetExistence("3.Solution")
    If WorksheetExists = True Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("3.Solution").Delete
        Application.DisplayAlerts = True
    End If
    
    WorksheetExists = CheckWorksheetExistence("2.3.Container-ItemCompatibility")
    If WorksheetExists = True Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Delete
        Application.DisplayAlerts = True
    End If
    
    ThisWorkbook.Worksheets("2.Containers").Activate
    
    Range(Cells(2, 2), Cells(num_container_types + 1, 10)).Select
    
    Selection.Sort Key1:=Columns("B"), Order1:=xlAscending, Header:=xlNo
    
End Sub
Sub SetupMenuItems()
    Call ThisWorkbook.Workbook_Activate
End Sub

Private Sub WatchTutorial()
    ThisWorkbook.FollowHyperlink "https://www.youtube.com/watch?v=-0MCkd-CjD0"
End Sub


' ribbon calls and tab activation

#If Win32 Or Win64 Or (MAC_OFFICE_VERSION >= 15) Then

Sub ResetWorkbookRibbonCall(control As IRibbonControl)
    Call ResetWorkbook
End Sub
Sub SetupItemsWorksheetRibbonCall(control As IRibbonControl)
    Call SetupItemsWorksheet
End Sub
Sub SortItemsRibbonCall(control As IRibbonControl)
    Call SortItemTypes
End Sub
Sub SetupItemItemCompatibilityWorksheetRibbonCall(control As IRibbonControl)
    Call SetupItemItemCompatibilityWorksheet
End Sub
Sub SetupContainersWorksheetRibbonCall(control As IRibbonControl)
    Call SetupContainersWorksheet
End Sub
Sub SortContainerTypesRibbonCall(control As IRibbonControl)
    Call SortContainerTypes
End Sub
Sub SetupContainerItemCompatibilityWorksheetRibbonCall(control As IRibbonControl)
    Call SetupContainerItemCompatibilityWorksheet
End Sub
Sub SetupSolutionWorksheetRibbonCall(control As IRibbonControl)
    Call SetupSolutionWorksheet
End Sub
Sub SetupVisualizationWorksheetRibbonCall(control As IRibbonControl)
    Call SetupVisualizationWorksheet
End Sub
Sub AnimateVisualizationWorksheetRibbonCall(control As IRibbonControl)
    Call AnimateVisualizationWorksheet
End Sub
Private Sub SendFeedbackRibbonCall(control As IRibbonControl)
    Call SendFeedback
End Sub
Private Sub WatchTutorialRibbonCall(control As IRibbonControl)
    Call WatchTutorial
End Sub
Private Sub AboutRibbonCall(control As IRibbonControl)
    Call About
End Sub
Sub tabActivate(ribbon As IRibbonUI)
    ribbon.ActivateTab ("CLPSpreadsheetSolver")
End Sub

#End If
