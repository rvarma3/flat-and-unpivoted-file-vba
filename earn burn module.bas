Attribute VB_Name = "Module1"
Sub earn_burn()
'
' Macro1 Macro
'

Dim ValueRange As Range
Dim OriginRange As Range
Dim DestRange As Range
'Dim macro_book As Workbook
Dim current_book As Workbook
Dim new_book As Workbook
Dim LastRow1 As Long
Dim LastRow2 As Long

Dim myFile As String
Dim myPath As String

'Dim Table1 As Range
'Dim Table2 As Range
'

Application.ScreenUpdating = False

'MsgBox Environ("username")

myPath = "C:\Users\" & Environ("username") & "\OneDrive - Emirates Group\earn burn master file"
'MsgBox myPath
myFile = Dir(myPath & "\" & "*EK Skywards Earn Burn MasterSheet*.xlsx")

'targetfile = myPath & myFile

Set current_book = Workbooks.Open(myPath & "\" & myFile)
'Set current_book = macro_book
Set new_book = Workbooks.Add



' loop through each sheet to stack values in another sheet



For i = 2 To 19

        current_book.Activate

' define the ranges

        Set OriginRange = current_book.Sheets(i).Range("a1").Offset(4, 2)
        Set DestRange = OriginRange.Offset(rowoffset:=2)
        Set ValueRange = OriginRange.Offset(2, 2)


' copy and paste the values into another sheet


        ' copy the origin region
        OriginRange.Copy

        new_book.Activate

        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        ActiveCell.Copy

        ActiveSheet.Range(ActiveCell.Address, ActiveCell.Offset(rowoffset:=17)).Select

        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        ' swtich between workbooks to select cells of interest
        
        'copy the detination region

        current_book.Sheets(i).Activate

        ActiveSheet.Range(DestRange.Address, DestRange.Offset(rowoffset:=17)).Copy

        new_book.Activate

        ActiveCell.Offset(columnoffset:=1).Select

        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        'copy the  value range for the miles

        current_book.Sheets(i).Activate

        ActiveSheet.Range(ValueRange.Address, ActiveSheet.Range(ValueRange.Address).End(xlToRight).End(xlDown)).Select


        Selection.Resize(Selection.Rows.Count + -1, Selection.Columns.Count).Copy

        new_book.Activate
        
        ActiveSheet.Range("c" & ActiveCell.CurrentRegion.Rows.Count + -17, _
        "c" & ActiveCell.CurrentRegion.Rows.Count).Select  ' Value = "6th Freedom"
        
        ActiveSheet.Range("c" & ActiveCell.CurrentRegion.Rows.Count + -17, _
        "c" & ActiveCell.CurrentRegion.Rows.Count).Value = "6th Freedom"

        ActiveCell.Offset(columnoffset:=1).Select

        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        ActiveCell.End(xlToLeft).End(xlDown).Offset(1).Select
        
        
               
        
        
        'fifth freedom
        
        OriginRange.Copy
        
        new_book.Activate

        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        ActiveCell.Copy

        ActiveSheet.Range(ActiveCell.Address, ActiveCell.Offset(rowoffset:=17)).Select

        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        current_book.Sheets(i).Activate

        ActiveSheet.Range(DestRange.Address, DestRange.Offset(rowoffset:=17)).Copy

        new_book.Activate

        ActiveCell.Offset(columnoffset:=1).Select

        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        

        new_book.Activate
        
        ActiveSheet.Range(ActiveCell.End(xlDown).Offset(columnoffset:=1), ActiveCell.Offset(columnoffset:=1).End(xlUp).Offset(rowoffset:=1)).Value = "5th Freedom"
        
        current_book.Sheets(i).Activate

        ActiveSheet.Range(ValueRange.Offset(rowoffset:=51).Address, ActiveSheet.Range(ValueRange.Offset(rowoffset:=52).Address).End(xlDown).End(xlToRight)).Select


        Selection.Resize(Selection.Rows.Count - 1, Selection.Columns.Count).Copy
        
        new_book.Activate
        
        ActiveCell.Offset(columnoffset:=2).Select

        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        ActiveCell.End(xlToLeft).End(xlDown).Offset(1).Select
        
        
        
        
        

' deactivate copy paste
        Application.CutCopyMode = False

Next i

' create blank columns to perform look ups

new_book.Activate

ActiveSheet.Range("a1").EntireRow.Insert

For i = 1 To 3

        ActiveSheet.Range("a1").Offset(columnoffset:=3).EntireColumn.Insert

Next i

' fill in the headers for the file

'ActiveSheet.Range("a2", ActiveSheet.Range("a2").End(xlDown)).Offset(columnoffset:=2).Value = "6th Freedom"

ActiveSheet.Range("a1").Offset(columnoffset:=0).Value = "OriginRegion"
ActiveSheet.Range("a1").Offset(columnoffset:=1).Value = "DestinationRegion"
ActiveSheet.Range("a1").Offset(columnoffset:=2).Value = "Freedom"
ActiveSheet.Range("a1").Offset(columnoffset:=3).Value = "OriginZone"
ActiveSheet.Range("a1").Offset(columnoffset:=4).Value = "DestinationZone"
ActiveSheet.Range("a1").Offset(columnoffset:=5).Value = "Zonepair"
ActiveSheet.Range("a1").Offset(columnoffset:=6).Value = "Y_Special"
ActiveSheet.Range("a1").Offset(columnoffset:=7).Value = "Y_Saver"
ActiveSheet.Range("a1").Offset(columnoffset:=8).Value = "Y_Flex"
ActiveSheet.Range("a1").Offset(columnoffset:=9).Value = "Y_Flex Plus"
ActiveSheet.Range("a1").Offset(columnoffset:=10).Value = "PY_Flex Plus"
ActiveSheet.Range("a1").Offset(columnoffset:=11).Value = "J_Special"
ActiveSheet.Range("a1").Offset(columnoffset:=12).Value = "J_Saver"
ActiveSheet.Range("a1").Offset(columnoffset:=13).Value = "J_Flex"
ActiveSheet.Range("a1").Offset(columnoffset:=14).Value = "J_Flex Plus"
ActiveSheet.Range("a1").Offset(columnoffset:=15).Value = "F_Flex Plus"
ActiveSheet.Range("a1").Offset(columnoffset:=16).Value = "F_Flex"


' format cells and fonts

ActiveSheet.Range("g1").Select
Range(Selection, Selection.End(xlToRight)).Select

    With Selection
        .Font.Bold = True
        .Font.ThemeColor = xlThemeColorDark1
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 5287936
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
    End With
    
    
ActiveSheet.Range("g1").Offset(columnoffset:=-1).Select
Range(Selection, Selection.End(xlToLeft)).Select

    With Selection
        .Font.Bold = True
        .Font.ThemeColor = xlThemeColorDark1
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 12611584
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
    End With
    


' formating zone summary sheet to perfrom look ups around the zone

current_book.Sheets(1).Activate

ActiveSheet.Range("h1").EntireColumn.Delete

ActiveSheet.Range("a2").CurrentRegion.Select


' unmerge the cells for vlookup


Selection.UnMerge
'

'create an object that seeks the last row in the sheet

LastRow1 = current_book.Sheets(1).Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

ActiveSheet.Range("a3:g" & LastRow1).Select



' similar to go to special cells in excel to auto fill the values in the excel sheet

On Error GoTo errHandler

Selection.SpecialCells(xlCellTypeBlanks).Select
Application.CutCopyMode = False
Selection.FormulaR1C1 = "=R[-1]C"


ActiveSheet.Range("a3:g" & LastRow1).Copy
ActiveSheet.Range("a3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

ActiveSheet.Range("b3:b" & LastRow1).Select

Cells.Replace What:=Chr(10), Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False


ActiveSheet.Range("a1").EntireColumn.Copy

ActiveSheet.Range("a1").Offset(columnoffset:=2).Insert Shift:=xlToRight

errHandler:

Application.CutCopyMode = False

new_book.ActiveSheet.Activate

'ActiveCell.Offset(columnoffset:=3).End(xlUp).Offset(rowoffset:=1).Select


' last row for the new workbook
LastRow2 = new_book.ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
'ActiveSheet.Range("a2:b" & LastRow2).Select



Cells.Replace What:="Dubai", Replacement:="Dubai / UAE", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
Cells.Replace What:=" - UAE", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False

Cells.Replace What:="Asian Sub-Continent (North)", Replacement:="Asian Sub Continent (North)", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
Cells.Replace What:="Asian Sub-Cont. (North)", Replacement:="Asian Sub Continent (North)", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False

Cells.Replace What:="Asian Sub-Continent (South)", Replacement:="Asian Sub Continent (South)", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
Cells.Replace What:="Asian Sub-Cont. (South)", Replacement:="Asian Sub Continent (South)", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False

Cells.Replace What:="Asian Sub-Continent (East)", Replacement:="Asian Sub Continent (East)", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
Cells.Replace What:="Asian Sub-Cont. (East)", Replacement:="Asian Sub Continent (East)", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False


'--------------------------------------


' vlookup function (fails if an NA value get populated

'On Error Resume Next
Dim Dept_Row As Double
Dim Dept_Clm As Double
Table1 = current_book.Sheets(1).Range("b3:g" & LastRow1) ' look up table
Table2 = new_book.Sheets(1).Range("a2:a" & LastRow2)  ' range of values to look up
'''Table1 = Sheet1.Range("A3:A13") ' Employee_ID Column from Employee table
'''Table2 = Sheet1.Range("H3:I13") ' Range of Employee Table 1
Dept_Row = new_book.ActiveSheet.Range("d2").Row ' Change E3 with the cell from where you need to start populating the Department
Dept_Clm = new_book.ActiveSheet.Range("d2").Column
'
For Each cl In Table2
    new_book.ActiveSheet.Cells(Dept_Row, Dept_Clm) = Application.WorksheetFunction.VLookup(cl, Table1, 2, 0)

    Dept_Row = Dept_Row + 1
Next cl
'MsgBox Dept_Row & Dept_Clm
'current_book.Sheets(1).Range("b3:g299")

new_book.ActiveSheet.Activate



Dept_Row = new_book.ActiveSheet.Range("e2").Row ' Change E3 with the cell from where you need to start populating the Department
Dept_Clm = new_book.ActiveSheet.Range("e2").Column
Table2 = new_book.Sheets(1).Range("b2:b" & LastRow2)
'

For Each cl2 In Table2
    new_book.ActiveSheet.Cells(Dept_Row, Dept_Clm) = Application.WorksheetFunction.VLookup(cl2, Table1, 2, 0)

    Dept_Row = Dept_Row + 1
Next cl2




Dept_Row = new_book.ActiveSheet.Range("f2").Row ' Change E3 with the cell from where you need to start populating the Department
Dept_Clm = new_book.ActiveSheet.Range("f2").Column

For Each cl3 In Table2
    new_book.ActiveSheet.Cells(Dept_Row, Dept_Clm) = new_book.ActiveSheet.Cells(Dept_Row, Dept_Clm - 2).Value & "-" _
    & new_book.ActiveSheet.Cells(Dept_Row, Dept_Clm - 1).Value

    Dept_Row = Dept_Row + 1
Next cl3





ActiveCell.End(xlUp).End(xlUp).Offset(1, 5).Select



Application.ScreenUpdating = True


new_book.Activate


ActiveSheet.Range("a1").Select

' resize columns ro autfit contents

Cells.Select
Cells.EntireColumn.AutoFit

ActiveWindow.DisplayGridlines = False

MsgBox Application.UserName & ": Yo! its done"
'
' save the newly created workbook

new_book.ActiveSheet.Name = "Flat File"

new_book.SaveAs current_book.Path & "\earn_burn_flat_file_" & Format(Date, "YYYYMMDD") & ".csv"
'
current_book.Close SaveChanges:=False
'
''new_book.Close

End Sub



Sub testing()


'
' Macro1 Macro
'

Dim ValueRange As Range
Dim OriginRange As Range
Dim DestRange As Range
'Dim macro_book As Workbook
Dim current_book As Workbook
Dim new_book As Workbook
Dim LastRow1 As Long
Dim LastRow2 As Long

'Dim Table1 As Range
'Dim Table2 As Range
'

Application.ScreenUpdating = False


Set current_book = Workbooks.Open(ThisWorkbook.Path & "\EK Skywards Earn Burn MasterSheet_17Mar2022.xlsx")
'Set current_book = macro_book
Set new_book = Workbooks.Add



' loop through each sheet to stack values in another sheet



For i = 2 To 19

        current_book.Activate

' define the ranges

        Set OriginRange = current_book.Sheets(i).Range("a1").Offset(4, 2)
        Set DestRange = OriginRange.Offset(rowoffset:=2)
        Set ValueRange = OriginRange.Offset(2, 2)


' copy and paste the values into another sheet


        ' copy the origin region
        OriginRange.Copy

        new_book.Activate

        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        ActiveCell.Copy

        ActiveSheet.Range(ActiveCell.Address, ActiveCell.Offset(rowoffset:=17)).Select

        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        ' swtich between workbooks to select cells of interest
        
        'copy the detination region

        current_book.Sheets(i).Activate

        ActiveSheet.Range(DestRange.Address, DestRange.Offset(rowoffset:=17)).Copy

        new_book.Activate

        ActiveCell.Offset(columnoffset:=1).Select

        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        'copy the  value range for the miles

        current_book.Sheets(i).Activate

        ActiveSheet.Range(ValueRange.Address, ActiveSheet.Range(ValueRange.Address).End(xlToRight).End(xlDown)).Select


        Selection.Resize(Selection.Rows.Count + -1, Selection.Columns.Count).Copy

        new_book.Activate
        
        ActiveSheet.Range("c" & ActiveCell.CurrentRegion.Rows.Count + -17, _
        "c" & ActiveCell.CurrentRegion.Rows.Count).Select  ' Value = "6th Freedom"
        
        ActiveSheet.Range("c" & ActiveCell.CurrentRegion.Rows.Count + -17, _
        "c" & ActiveCell.CurrentRegion.Rows.Count).Value = "6th Freedom"

        ActiveCell.Offset(columnoffset:=1).Select

        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        ActiveCell.End(xlToLeft).End(xlDown).Offset(1).Select
        
        
               
        
        
        'fifth freedom
        
        OriginRange.Copy
        
        new_book.Activate

        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

        ActiveCell.Copy

        ActiveSheet.Range(ActiveCell.Address, ActiveCell.Offset(rowoffset:=17)).Select

        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        current_book.Sheets(i).Activate

        ActiveSheet.Range(DestRange.Address, DestRange.Offset(rowoffset:=17)).Copy

        new_book.Activate

        ActiveCell.Offset(columnoffset:=1).Select

        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        

        new_book.Activate
        
        ActiveSheet.Range(ActiveCell.End(xlDown).Offset(columnoffset:=1), ActiveCell.Offset(columnoffset:=1).End(xlUp).Offset(rowoffset:=1)).Value = "5th Freedom"
        
        current_book.Sheets(i).Activate

        ActiveSheet.Range(ValueRange.Offset(rowoffset:=51).Address, ActiveSheet.Range(ValueRange.Offset(rowoffset:=52).Address).End(xlDown).End(xlToRight)).Select


        Selection.Resize(Selection.Rows.Count - 1, Selection.Columns.Count).Copy
        
        new_book.Activate
        
        ActiveCell.Offset(columnoffset:=2).Select

        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        ActiveCell.End(xlToLeft).End(xlDown).Offset(1).Select
        
        
        
        
        

' deactivate copy paste
        Application.CutCopyMode = False

Next i


End Sub



Public Sub unpivotData()
   
    ' Pass source/destination in as parameters or just define here
    ' Public Function tranposeData(inputSheetName As String, outputSheetName As String)
    
    Dim current_book As Workbook
    
    Set current_book = Workbooks.Open(ThisWorkbook.Path & "\earn_burn_flat_file_" & _
    Format(Date, "YYYYMMDD") & ".csv")
    
    
    Dim inputSheetName As String
    inputSheetName = ActiveSheet.Name
    
    Dim outputSheetName As String
    outputSheetName = "Unpivoted"
   
   
    'How many static columns are there?
    Dim staticColumnCount As Integer
    staticColumnCount = 6
       
       
    ' Variables to reference the worksheets we're using
    Dim inputSheet As Worksheet, outputSheet As Worksheet
    Dim oldRow As Long, theDate As Long, newRow As Long
    
    Set inputSheet = ActiveWorkbook.Worksheets(inputSheetName)
    
    Set outputSheet = Worksheets.Add
        outputSheet.Select
        outputSheet.Name = outputSheetName ' Rename
    
    

  
    ' Fill in static column headers, based on number of columns
    For i = 1 To staticColumnCount
        outputSheet.Cells(1, i).Value = inputSheet.Cells(1, i).Value
    Next i


    'Fill in dynamic column headers for the transposed values
    outputSheet.Cells(1, staticColumnCount + 1).Value = "BrandClass"
    outputSheet.Cells(1, staticColumnCount + 2).Value = "MilesValue"
    
    

    'Start the actual transposing
    inputRow = 2
    outputRow = 2
    
    ' Starting at Row 1, go DOWN until it its a blank value
    Do While inputSheet.Cells(inputRow, 1).Value <> ""
        
        ' the "transposed" (dynamic) values will start one column after the static values
        inputColumn = staticColumnCount + 1
        
        ' Go to the RIGHT until you hit a blank value
        Do While inputSheet.Cells(1, inputColumn).Value <> ""
        
            'Put the value in, for however many static columns there are
            For j = 1 To staticColumnCount
            
                ' Hard-coded sample if there were 2 static columns:
                'outputSheet.Cells(outputRow, 1).value = inputSheet.Cells(inputRow, 1)
                'outputSheet.Cells(outputRow, 2).value = inputSheet.Cells(inputRow, 2)
                
                ' Dynamic version:
                outputSheet.Cells(outputRow, j).Value = inputSheet.Cells(inputRow, j)
                
            Next j

            
            ' Transpose the column header
            outputSheet.Cells(outputRow, staticColumnCount + 1).Value = inputSheet.Cells(1, inputColumn)
            
            
            ' Transpose the actual value
            outputSheet.Cells(outputRow, staticColumnCount + 2).Value = inputSheet.Cells(inputRow, inputColumn)
            
            inputColumn = inputColumn + 1
            outputRow = outputRow + 1
        Loop
        
        inputRow = inputRow + 1
    Loop
    

    
    ' Format the data as needed

    'Fromat TEXT
    'outputSheet.Columns("A:B").Select
    'Selection.NumberFormat = "@"

    'Format DATE
    'outputSheet.Columns("C:C").Select
    'Selection.NumberFormat = "m/d/yyyy"

    'Format NUMBERS
    'outputSheet.Columns("D:D").Select
    'Selection.Style = "Comma"
    
    
    
    

    'Freeze the first row
'    With ActiveWindow
'        .SplitColumn = 0
'        .SplitRow = 1
'    End With
'    ActiveWindow.FreezePanes = True

    'Select cell A1 when done
    outputSheet.Cells(1, 1).Select
    
    Range(Selection, Selection.Offset(columnoffset:=5)).Select

    With Selection
        .Font.Bold = True
        .Font.ThemeColor = xlThemeColorDark1
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 5287936
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
    End With
    
ActiveCell.Select
Range(Selection.Offset(columnoffset:=6), Selection.Offset(columnoffset:=6).End(xlToRight)).Select

    With Selection
        .Font.Bold = True
        .Font.ThemeColor = xlThemeColorDark1
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 12611584
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
    End With
    
Cells.Select
Cells.EntireColumn.AutoFit

ActiveWindow.DisplayGridlines = False
    
current_book.SaveAs current_book.Path & "\earn_burn_unpivoted_file_" & Format(Date, "YYYYMMDD") & ".xlsx"
    

MsgBox Application.UserName & ": Yo! its done again"





End Sub







