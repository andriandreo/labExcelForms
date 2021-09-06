Attribute VB_Name = "calcModule"
'=====================CREDITS======================'
'AUTHOR: Andres Alberto Andreo Acosta'
'GitHub: https://github.com/andriandreo'
'DATE (DD/MM/YY): 06/09/21'
'Version: v2.0'
'LICENSE: MIT License'
'================================================='

Sub CreateName()

'=============CREATE A RANGE TO PLOT EACH TIME TRACE TO COMPARE===================
        Dim wb As Workbook
        Dim ws As Worksheet
        Dim rname As String 'Declare the name variable as a String
        Set wb = ActiveWorkbook
        Set ws = ActiveSheet

            'Read the time between additions (each "?" cells)
            addt = Worksheets(ActiveSheet.Name).Cells(1, 13)

        rStart = Worksheets(ActiveSheet.Name).Cells(1, 12) - addt ' "-addt" in order to capture before the first step
        rname = Mid(ws.Name, 1, InStr(1, ws.Name, "(") - 1) & Mid(ws.Name, InStr(1, ws.Name, "(") + 1, 1)
        wb.Worksheets(ws.Name).Names.Add Name:=rname, RefersTo:=Worksheets(ActiveSheet.Name).Range(Worksheets(ActiveSheet.Name).Cells(rStart, 2), Worksheets(ActiveSheet.Name).Cells(16000, 2))
'=====================================================================

End Sub


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

            'Read the time between additions (each "?" cells)
            addt = Worksheets(ActiveSheet.Name).Cells(1, 13)

        rStart = Worksheets(ActiveSheet.Name).Cells(1, 12) - addt ' "-addt" in order to capture before the first step
        rname = Mid(ws.Name, 1, InStr(1, ws.Name, "(") - 1) & Mid(ws.Name, InStr(1, ws.Name, "(") + 1, 1)
        wb.Worksheets(ws.Name).Names.Add Name:=rname, RefersTo:=Worksheets(ActiveSheet.Name).Range(Worksheets(ActiveSheet.Name).Cells(rStart, 2), Worksheets(ActiveSheet.Name).Cells(16000, 2))
'=====================================================================

'As "non-visible" steps are not readable:
iStep = Val(InputBox("What is the first readable step in the time trace? (1, 2, 3...):", "1st step?"))

'In the case you waited (or missed to add) for a longer time in a step:
spStep = Val(InputBox("Any special step? In affirmative case, write the step number (otherwise, type 0):", "Special step?")) '[!!!!!!]
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

'k = iStep
For k = iStep To 3 'In the case the initial steps correspond to the very first, the calculation for the subsequent may not be accurate
'Because 1st step is not always readible
j = Worksheets(ActiveSheet.Name).Cells(1, 12) + (k * addt - 20) 'Initialise reading variable (at a reasonable value when 2nd sweep started)

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
nSteps = Worksheets(ActiveSheet.Name).Cells(2, 12) 'The number of additions specified in the form

errflag = 0 'The flag to control the overflow error
For i = k To nSteps
 Id1 = Worksheets(ActiveSheet.Name).Cells(j, 2)
 Id2 = Worksheets(ActiveSheet.Name).Cells(j + 1, 2)
 diff = Abs(Id2 - Id1) 'Difference between previous and next current

 Do While diff < cutoff 'While no step, continue reading
  Id1 = Worksheets(ActiveSheet.Name).Cells(j, 2)
  Id2 = Worksheets(ActiveSheet.Name).Cells(j + 1, 2)
  diff = Abs(Id2 - Id1)
  j = j + 1
  If j = 1000000 Then 'In the case the overflow error appears for the reading variable ("j") and the considered "cutoff"
   errflag = 1
   Exit Do
  End If
 Loop

 If errflag = 1 Then Exit For 'In the case the overflow error appears for the reading variable ("j") and the considered "cutoff"
 
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
Next k

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

Sub RegStudy()

 'Declare two variables as matrices for registering the regression stats and calculate the residuals (and the y, x ranges), respectively:
  Dim rstats() As Variant
  Dim residuals() As Variant
  Dim Y As Range
  Dim X As Range
  Dim Yold As Range 'Saves the previous Y range
  Dim Xold As Range 'Saves the previous X range
  Dim dY As Range 'Background corrected data
  Dim dYold As Range
  Dim nrY As Range 'Normalised response
  Dim nrYold As Range
  Dim wb As Workbook
  Dim ws As Worksheet
  Dim rs As Worksheet 'To save the current "RegStudy" Sheet
  
  Set wb = ActiveWorkbook
  Set ws = ActiveSheet 'Store the name of the "CCalc" Active Sheet

 'Define the number of observations and the desired ranges:
  nSteps = Worksheets(ws.Name).Cells(2, 12) 'The number of additions specified in the form
  'As "non-visible" steps are not readable:
   iStep = Val(InputBox("What is the first readable step in the time trace? (1, 2, 3...):", "1st step?"))
  n = nSteps - iStep + 1 'Calculate the number of observations (based on the "CCalcId" Module, +1 to count since the 1st visible -not the 2nd-)
  it = 0 'Indicates the number of completed iterations
  Set Y = ws.Range(ActiveSheet.Cells(4 + iStep, 12), ws.Cells(4 + iStep + (n - 1), 12))
  Set nrY = ws.Range(ActiveSheet.Cells(4 + iStep, 13), ws.Cells(4 + iStep + (n - 1), 13))
  Set dY = ws.Range(ActiveSheet.Cells(4 + iStep, 14), ws.Cells(4 + iStep + (n - 1), 14))
  Set X = ws.Range(ws.Cells(4 + iStep, 11), ws.Cells(4 + iStep + (n - 1), 11))

 'Calculate a 2x5 matrix with the linear regression and its stats:
  rstats() = Application.WorksheetFunction.LinEst(Y, X, , True)
  
 'Calculate the Residuals as the difference between the read data and the data predicted by the function "TREND":
  residuals() = Application.WorksheetFunction.Trend(Y, X, X)
  For i = 1 To n
   residuals(i, 1) = ws.Cells(3 + iStep + i, 12) - residuals(i, 1)
  Next i
   
'================CREATE A NEW SHEET AND WRITE THE RESULTS (DEBUGGING)================
  Sheets.Add.Name = "RegStudy(" & Mid(ws.Name, InStr(1, ws.Name, "(") + 1, 1) & ")"
  Set rs = ActiveSheet 'Because after the ActiveSheet moved to "RegStudy" after creating it
  
  'Loop for writing the "rstats":
     rs.Cells(1, 1) = "rstats"
     For l = 1 To 2
      For i = 1 To 5
       rs.Cells(i + 1, l) = rstats(i, l)
      Next i
     Next l
  
   'Loop for writing the residuals:
     rs.Cells(1, 3) = "Residuals"
     For i = 1 To n
       rs.Cells(i + 1, 3) = residuals(i, 1)
     Next i
'=====================================================================
   
'===================CALCULATING LOOP========================================
 up = 1 'Initialise the variable for studying the lower and upper limits within the loop ("=1" to force start reducing by the lower limit)
 it = 1 'Indicates the number of completed iterations
 j = 0 'The variable to control the lower limit when is the lower the one which changes
 k = 1 'The variable to control the upper limit when is the lower the one which changes ("=1" because the upper limit refers to "n0")
 n0 = n 'To save the initial value of the observations
 Do Until n = 4 'In order to have at least 4 points within the linear range
  
  Set Yold = Y 'Saves the previous Y range(s)
  Set nrYold = nrY
  Set dYold = dY
  Set Xold = X 'Saves the previous X range
  R2old = rstats(3, 1) 'Saves the value for the previous R2 coefficient
  pVARold = rstats(5, 2) / n 'Saves the value for the previous pVAR coefficient (VAR considering a mean value = 0)
  
  'Calculate the max for the residuals and determine its position:
   Mres = Abs(residuals(1, 1))
   For i = 2 To n
    If Abs(residuals(i, 1)) > Mres Then
     Mres = Abs(residuals(i, 1))
    End If
   Next i
   'Look for the position of the max
   i = 1
   Do Until Mres = Abs(residuals(i, 1))
    i = i + 1
   Loop
      
   If n = 6 Or n = 5 Then
    Set Y = ws.Range(ws.Cells(4 + iStep + (j + 1), 12), ws.Cells(4 + iStep + (n0 - k), 12))
    Set X = ws.Range(ws.Cells(4 + iStep + (j + 1), 11), ws.Cells(4 + iStep + (n0 - k), 11))
     rstats() = Application.WorksheetFunction.LinEst(Y, X, , True)
     R2a = rstats(3, 1)
     pVARa = rstats(5, 2) / (n - 1)
    Set Y = ws.Range(ws.Cells(4 + iStep + j, 12), ws.Cells(4 + iStep + (n0 - (k + 1)), 12))
    Set X = ws.Range(ws.Cells(4 + iStep + j, 11), ws.Cells(4 + iStep + (n0 - (k + 1)), 11))
     rstats() = Application.WorksheetFunction.LinEst(Y, X, , True)
     R2b = rstats(3, 1)
     pVARb = rstats(5, 2) / (n - 1)
    If R2a > R2b And pVARa < pVARb Then
     j = j + 1
    Else
     k = k + 1
    End If
    
   Else
   If i = 1 Then
    j = j + 1
    up = 0 'In the case the range was previously reduced by the upper limit (or "i=1" in the first iteration)
   ElseIf i = n Then
    k = k + 1
   ElseIf up = 1 Then
    j = j + 1
    up = 0 'If the range was reduced by the lower limit
   Else
    k = k + 1
    up = 1 'If the range was reduced by the upper limit
   End If
   End If
   n = n0 - it 'The number of observations has been reduced by 1 after this iteration
   
   Set Y = ws.Range(ws.Cells(4 + iStep + j, 12), ws.Cells(4 + iStep + (n0 - k), 12))
   Set nrY = ws.Range(ws.Cells(4 + iStep + j, 13), ws.Cells(4 + iStep + (n0 - k), 13))
   Set dY = ws.Range(ws.Cells(4 + iStep + j, 14), ws.Cells(4 + iStep + (n0 - k), 14))
   Set X = ws.Range(ws.Cells(4 + iStep + j, 11), ws.Cells(4 + iStep + (n0 - k), 11))
    
   'Calculate the regression stats and residuals with the new ranges:
    rstats() = Application.WorksheetFunction.LinEst(Y, X, , True)
    residuals() = Application.WorksheetFunction.Trend(Y, X, X)
    For i = 1 To n
     residuals(i, 1) = ws.Cells(3 + iStep + j + i, 12) - residuals(i, 1) ' "j" for considering the removed lower limits
    Next i
    
    R2 = rstats(3, 1) 'Saves the value for the new R2 coefficient
    pVAR = rstats(5, 2) / n 'Saves the value for the new pVAR coefficient (VAR considering a mean value = 0)

     If n = 4 And R2 < R2old And 0.7 * pVARold < pVAR Then 'Exit the loop before if the last iteration will give worse results
     Set Y = Yold
     Set nrY = nrYold
     Set dY = dYold
     Set X = Xold
     n = n + 1 '+1 as it is for the previous ranges
     'Calculate the regression stats and residuals with the previous ranges:
      rstats() = Application.WorksheetFunction.LinEst(Y, X, , True)
      residuals() = Application.WorksheetFunction.Trend(Y, X, X)
      For i = 1 To n
       residuals(i, 1) = ActiveSheet.Cells(3 + iStep + j + i, 12) - residuals(i, 1) ' "j" for considering the removed lower limits
      Next i
     Exit Do
    End If
    
    '===============CREATE A NEW SHEET AND WRITE THE RESULTS (DEBUGGING)===============
    
    'Loop for writing the "rstats":
     rs.Cells(2 + 10 * it, 1) = "rstats"
     For l = 1 To 2
      For i = 1 To 5
       rs.Cells(i + 2 + 10 * it, l) = rstats(i, l)
      Next i
     Next l
  
    'Loop for writing the residuals:
     rs.Cells(2 + 10 * it, 3) = "Residuals"
     For i = 1 To n
       rs.Cells(2 + 10 * it + i, 3) = residuals(i, 1)
     Next i
   
     rs.Cells(2 + 10 * it, 4) = "up"
     rs.Cells(3 + 10 * it, 4) = up
   
     rs.Cells(2 + 10 * it, 5) = "j"
     rs.Cells(3 + 10 * it, 5) = j
   
     rs.Cells(2 + 10 * it, 6) = "k"
     rs.Cells(3 + 10 * it, 6) = k
    
     rs.Cells(4 + 10 * it, 4) = "Lin. range:"
     rs.Cells(4 + 10 * it, 5) = X(1, 1)
     rs.Cells(4 + 10 * it, 6) = X(n, 1)
    '===================================================================
     
   it = it + 1 'Increase the number of completed iterations by 1
  Loop
'======================================================================
  
'=============CREATE A RANGE TO PLOT EACH LINEAR RANGE TO COMPARE===================
        Dim xname As String 'Declare the name variable as a String
        Dim yname As String 'Declare the name variable as a String
        Dim nrname As String 'Declare the name variable as a String
        Dim dyname As String 'Declare the name variable as a String

        xname = "xLIN_" & Mid(ws.Name, InStr(1, ws.Name, "(") + 1, 1)
        yname = "yLIN_" & Mid(ws.Name, InStr(1, ws.Name, "(") + 1, 1)
        nrname = "nrLIN_" & Mid(ws.Name, InStr(1, ws.Name, "(") + 1, 1)
        dyname = "dyLIN_" & Mid(ws.Name, InStr(1, ws.Name, "(") + 1, 1)
        wb.Worksheets(ws.Name).Names.Add Name:=xname, RefersTo:=X
        wb.Worksheets(ws.Name).Names.Add Name:=yname, RefersTo:=Y
        wb.Worksheets(ws.Name).Names.Add Name:=nrname, RefersTo:=nrY
        wb.Worksheets(ws.Name).Names.Add Name:=dyname, RefersTo:=dY
'=====================================================================
  
End Sub


