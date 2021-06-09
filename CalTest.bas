Attribute VB_Name = "calcModule"

'=====================CREDITS======================'
'AUTHOR: Andrés Alberto Andreo Acosta'
'GitHub: https://github.com/andriandreo'
'DATE (DD/MM/YY): 08/06/21'
'Version: v1.0'
'LICENSE: MIT License'
'=================================================='

Sub rangeNames()

        Dim wb As Workbook
        Dim ws As Worksheet
        Dim rname As String 'Declare the name variable as a String
        Set wb = ActiveWorkbook
        Set ws = ActiveSheet

        rStart = Worksheets(ActiveSheet.Name).Cells(1, 12)
        rname = Mid(ws.Name, 1, InStr(1, ws.Name, "(") - 1) & Mid(ws.Name, InStr(1, ws.Name, "(") + 1, 1)

        'Debugging names
        'Worksheets(ActiveSheet.Name).Cells(1, 9) = rname

        'Workbook scope:
        'wb.Names.Add Name:=rname, RefersTo:=Worksheets(ActiveSheet.Name).Range(Worksheets(ActiveSheet.Name).Cells(rStart, 2), Worksheets(ActiveSheet.Name).Cells(16000, 2))

        'Worksheet scope:
        wb.Worksheets(ws.Name).Names.Add Name:=rname, RefersTo:=Worksheets(ActiveSheet.Name).Range(Worksheets(ActiveSheet.Name).Cells(rStart, 2), Worksheets(ActiveSheet.Name).Cells(16000, 2))

End Sub

'=========================================='

Sub RCalcId()

j = 30 'Initialise reading variable (at a reasonable value where sweep started)

'Read data from the recording form (Id-Vd)
Dsteps = Worksheets(ActiveSheet.Name).Cells(2, 1) 'The number of Vd steps specified in the form
For i = 1 To Dsteps
 Id1 = Worksheets(ActiveSheet.Name).Cells(j, 2)
 Id2 = Worksheets(ActiveSheet.Name).Cells(j + 1, 2)
 diff = Abs(Id2 - Id1) 'Difference between previous and next current

 While diff < 0.0003 'While no step, continue reading
  Id1 = Worksheets(ActiveSheet.Name).Cells(j, 2)
  Id2 = Worksheets(ActiveSheet.Name).Cells(j + 1, 2)
  diff = Abs(Id2 - Id1)
  j = j + 1
 Wend

 'Set the cells' range for calculating the average of currents:
 Set avRange = Worksheets(ActiveSheet.Name).Range(Worksheets(ActiveSheet.Name).Cells(j - 21, 2), Worksheets(ActiveSheet.Name).Cells(j - 1, 2))
 'Write data to the plot cells:
 Worksheets(ActiveSheet.Name).Cells(17 + i, 5) = Application.WorksheetFunction.Average(avRange)

 j = j + 30 'In order to ensure next step

Next i

End Sub

'=========================================='

Sub CCalcId()

'=================================================================================
'If you HAVE NOT added the analytes at the right time (NOT CONSTANT TIME INTERVAL)
'=================================================================================

'Determine the parameters as a function of the System State
If Worksheets(ActiveSheet.Name).Cells(2, 13) = "QSS" Then 'Quasi Solid-State'
   Scoff = 0.3 'The cutoff correction factor for the steps
   Sstep = 53 'The increment in t for calculating the next step
ElseIf Worksheets(ActiveSheet.Name).Cells(2, 13) = "LS" Then 'Liquid-State'
   Scoff = 0.1 'The cutoff correction factor for the steps
   Sstep = 23 'The increment in t for calculating the next step
End If

'=============CREATE A RANGE TO PLOT EACH TIME TRACE TO COMPARE===================
        Dim wb As Workbook
        Dim ws As Worksheet
        Dim rname As String 'Declare the name variable as a String
        Set wb = ActiveWorkbook
        Set ws = ActiveSheet

        rStart = Worksheets(ActiveSheet.Name).Cells(1, 12)
        rname = Mid(ws.Name, 1, InStr(1, ws.Name, "(") - 1) & Mid(ws.Name, InStr(1, ws.Name, "(") + 1, 1)
        wb.Worksheets(ws.Name).Names.Add Name:=rname, RefersTo:=Worksheets(ActiveSheet.Name).Range(Worksheets(ActiveSheet.Name).Cells(rStart, 2), Worksheets(ActiveSheet.Name).Cells(16000, 2))
'=================================================================================

'In the case you waited (or missed to add) for a longer time in a step:
spStep = Val(InputBox("Any special step? In affirmative case, write the step number   (otherwise, type 0):", "Special step?"))
If spStep <> 0 Then
 spaddt = Val(InputBox("Type the duration time for this step (s):", "Special addition time (s)"))
End If

'Read the time between additions (each "?" cells)
addt = Worksheets(ActiveSheet.Name).Cells(1, 13)

'Calculate the average for the Io (stable current before the first addition)
 j = Worksheets(ActiveSheet.Name).Cells(1, 12) 'Set the variable at the time for the first addition
 'Set the cells' range for calculating the average of currents:
  Set avRange = Worksheets(ActiveSheet.Name).Range(Worksheets(ActiveSheet.Name).Cells(j - 22, 2), Worksheets(ActiveSheet.Name).Cells(j - 2, 2))
 'Write data:
  Worksheets(ActiveSheet.Name).Cells(4, 12) = Application.WorksheetFunction.Average(avRange)

'Because 1st step is not always readible
j = Worksheets(ActiveSheet.Name).Cells(1, 12) + (addt - 20) 'Initialise reading variable (at a reasonable value when 2nd sweep started)

'Loop for determine diffs' cutoff value:
 i = 2
 While Worksheets(ActiveSheet.Name).Cells(i, 2) <> 0
  Worksheets(ActiveSheet.Name).Cells(i, 1) = Abs(Worksheets(ActiveSheet.Name).Cells(i, 2) - Worksheets(ActiveSheet.Name).Cells(i + 1, 2))
  i = i + 1
 Wend

 'Set the cells' range for calculating the cutoff value for the diffs:
 Set coffRange = Worksheets(ActiveSheet.Name).Range(Worksheets(ActiveSheet.Name).Cells(j, 1), Worksheets(ActiveSheet.Name).Cells(j + (addt - 250), 1))
 cutoff = WorksheetFunction.Max(coffRange) - Scoff * (WorksheetFunction.Max(coffRange)) 'In order to reduce a little the max diff
 ActiveSheet.Columns(1).EntireColumn.ClearContents 'Clear the generated diffs column

'=====================DEBUGGING=======================
 'Write cutoff value in the sheet (debugging)
 Worksheets(ActiveSheet.Name).Cells(1, 1) = "Cutoff:"
 Worksheets(ActiveSheet.Name).Cells(2, 1) = cutoff
'=====================================================

'Read data from the recording form (the active sheet)
j = j + 6 'For getting closer to the 1st step
Dsteps = Worksheets(ActiveSheet.Name).Cells(2, 12) 'The number of additions specified in the form
For i = 1 To Dsteps
 Id1 = Worksheets(ActiveSheet.Name).Cells(j, 2)
 Id2 = Worksheets(ActiveSheet.Name).Cells(j + 1, 2)
 diff = Abs(Id2 - Id1) 'Difference between previous and next current

 While diff < cutoff 'While no step, continue reading
  Id1 = Worksheets(ActiveSheet.Name).Cells(j, 2)
  Id2 = Worksheets(ActiveSheet.Name).Cells(j + 1, 2)
  diff = Abs(Id2 - Id1)
  j = j + 1
 Wend

 'Set the cells' range for calculating the average of currents:
 Set avRange = Worksheets(ActiveSheet.Name).Range(Worksheets(ActiveSheet.Name).Cells(j - 21, 2), Worksheets(ActiveSheet.Name).Cells(j - 1, 2))
 'Write data:
 Worksheets(ActiveSheet.Name).Cells(4 + i, 12) = Application.WorksheetFunction.Average(avRange)

 If (i + 1) = spStep Then 'Considering the special step
  j = j + (spaddt - Sstep) 'To ensure next diff is a step (for the specified special step)
 Else
 j = j + (addt - Sstep) 'To ensure next diff is a step (each "?" defined cells)
 End If

Next i

End Sub

'=========================================='

Sub ctCCalcId()

'=========================================================================
'IF YOU HAVE ADDED THE ANALYTES AT THE RIGHT TIME (CONSTANT TIME INTERVAL)
'=========================================================================

'Read the time for the 1st addition (cells)
add0 = Worksheets(ActiveSheet.Name).Cells(1, 12)
'Read the time between additions (each "?" cells)
addt = Worksheets(ActiveSheet.Name).Cells(1, 13)

'Because 1st step is not always readible:
j = add0 + (addt - 3) 'Initialise reading variable (right before the 2nd sweep started)

'Read data from the recording form (the active sheet)
Dsteps = Worksheets(ActiveSheet.Name).Cells(2, 12) 'The number of additions specified in the form
For i = 1 To Dstep

 'Set the cells' range for calculating the average of currents:
 Set avRange = Worksheets(ActiveSheet.Name).Range(Worksheets(ActiveSheet.Name).Cells(j - 21, 2), Worksheets(ActiveSheet.Name).Cells(j - 1, 2))

 'Write data:
 Worksheets(ActiveSheet.Name).Cells(3 + i, 12) = Application.WorksheetFunction.Average(avRange)
 j = add0 + (i + 1) * addt - 3 'Going right before the next step (restarting "j" to the previous step)

Next i

End Sub

'=========================================='

Sub oldCalcIdgm()

j = 30 'Initialise reading variable (at a reasonable value where sweep started)

'Loop for determine diffs' cutoff value:
 i = 3
 While Worksheets(ActiveSheet.Name).Cells(i, 2) <> 0
  Worksheets(ActiveSheet.Name).Cells(i, 1) = Abs(Worksheets(ActiveSheet.Name).Cells(i, 2) - Worksheets(ActiveSheet.Name).Cells(i + 1, 2))
  i = i + 1
 Wend

 'Set the cells' range for calculating the cutoff value for the diffs:
 Set coffRange = Worksheets(ActiveSheet.Name).Range(Worksheets(ActiveSheet.Name).Cells(50, 1), Worksheets(ActiveSheet.Name).Cells(100, 1))
 cutoff = WorksheetFunction.Max(coffRange) - 0.45 * WorksheetFunction.Max(coffRange) 'In order to reduce a little the max diff
 ActiveSheet.Columns(1).Range(Cells(3, 1), Cells(i, 1)).ClearContents 'Clear the generated diffs column

'Read data from the recording form (Id-Vd)
Gsteps = Worksheets(ActiveSheet.Name).Cells(2, 1) 'The number of Vg steps specified in the form
For i = 1 To Gsteps
 Id1 = Worksheets(ActiveSheet.Name).Cells(j, 2)
 Id2 = Worksheets(ActiveSheet.Name).Cells(j + 1, 2)
 diff = Abs(Id2 - Id1) 'Difference between previous and next current

 While diff < cutoff 'While no step, continue reading
  Id1 = Worksheets(ActiveSheet.Name).Cells(j, 2)
  Id2 = Worksheets(ActiveSheet.Name).Cells(j + 1, 2)
  diff = Abs(Id2 - Id1)
  j = j + 1
 Wend

 'Set the cells' range for calculating the average of currents:
 Set avRange = Worksheets(ActiveSheet.Name).Range(Worksheets(ActiveSheet.Name).Cells(j - 21, 2), Worksheets(ActiveSheet.Name).Cells(j - 1, 2))
 'Write data to the plot cells:
 Worksheets(ActiveSheet.Name).Cells(17 + i, 7) = Application.WorksheetFunction.Average(avRange)

 j = j + 80 'In order to ensure next step

Next i

End Sub

'=========================================='

Sub gmCalcId()

j = 30 'Initialise reading variable (at a reasonable value where sweep started)

'Loop for determine diffs' cutoff value:
 i = 5
 While Worksheets(ActiveSheet.Name).Cells(i, 2) <> 0
  Worksheets(ActiveSheet.Name).Cells(i, 1) = Abs(Worksheets(ActiveSheet.Name).Cells(i, 2) - Worksheets(ActiveSheet.Name).Cells(i + 1, 2))
  i = i + 1
 Wend

 'Set the cells' range for calculating the cutoff value for the diffs:
 Set coffRange = Worksheets(ActiveSheet.Name).Range(Worksheets(ActiveSheet.Name).Cells(50, 1), Worksheets(ActiveSheet.Name).Cells(150, 1))
 cutoff = WorksheetFunction.Max(coffRange) - 0.45 * WorksheetFunction.Max(coffRange) 'In order to reduce a little the max diff
 ActiveSheet.Columns(1).Range(Cells(5, 1), Cells(i, 1)).ClearContents 'Clear the generated diffs column


'Read data from the recording form (Id-Vd)
Dsteps = Worksheets(ActiveSheet.Name).Cells(2, 1) 'The number of Vd steps specified in the form
Gsteps = Worksheets(ActiveSheet.Name).Cells(4, 1) 'The number of Vg steps specified in the form
w = 17 'The start and counting writing variable

For k = 1 To Dsteps
 For i = 1 To Gsteps
  Id1 = Worksheets(ActiveSheet.Name).Cells(j, 2)
  Id2 = Worksheets(ActiveSheet.Name).Cells(j + 1, 2)
  diff = Abs(Id2 - Id1) 'Difference between previous and next current

  While diff < cutoff 'While no step, continue reading
   Id1 = Worksheets(ActiveSheet.Name).Cells(j, 2)
   Id2 = Worksheets(ActiveSheet.Name).Cells(j + 1, 2)
   diff = Abs(Id2 - Id1)
   j = j + 1
  Wend

  'Set the cells' range for calculating the average of currents:
  Set avRange = Worksheets(ActiveSheet.Name).Range(Worksheets(ActiveSheet.Name).Cells(j - 21, 2), Worksheets(ActiveSheet.Name).Cells(j - 1, 2))
  'Write data to the plot cells:
  Worksheets(ActiveSheet.Name).Cells(w + i, 7) = Application.WorksheetFunction.Average(avRange)

  'If i = (Gsteps - 1) Then 'Because the last step lasts 200s instead of 100s
   'j = j + 170 'In order to ensure next step
  'Else
   j = j + 80 'In order to ensure next step
  'End If

 Next i
 w = w + i - 1 'To start the next loop right where it ended (before "Next i")

Next k

End Sub
