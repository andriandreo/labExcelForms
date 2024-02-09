Attribute VB_Name = "calcModule"

'=====================CREDITS======================'
'AUTHOR: Andres Alberto Andreo Acosta'
'GitHub: https://github.com/andriandreo'
'DATE (DD/MM/YY): 09/02/24'
'Version: v3.2'
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
            addt = ActiveSheet.Cells(1, 13)

        rStart = ActiveSheet.Cells(1, 12) - addt ' "-addt" in order to capture before the first step
        
        'As-measured Current:
        rname = Mid(ws.Name, 1, InStr(1, ws.Name, "(") - 1) & Mid(ws.Name, InStr(1, ws.Name, "(") + 1, 1)
        wb.Worksheets(ws.Name).Names.Add Name:=rname, RefersTo:=ActiveSheet.Range(ActiveSheet.Cells(rStart, 3), ActiveSheet.Cells(16000, 3)) ' Column#2 [A]; Column#3 [mA]
        
        'Baseline-substracted Current:
        rname = "d" & Mid(ws.Name, 1, InStr(1, ws.Name, "(") - 1) & Mid(ws.Name, InStr(1, ws.Name, "(") + 1, 1)
        wb.Worksheets(ws.Name).Names.Add Name:=rname, RefersTo:=ActiveSheet.Range(ActiveSheet.Cells(rStart, 4), ActiveSheet.Cells(16000, 4)) ' Column#3 [mA]; Column#4 dI [mA]

'=====================================================================

'Ask for converting [A] to [mA] in results:
uFlag = InputBox("Convert results from [A] to [mA] (y/n)?")
If uFlag = "y" Or uFlag = "Y" Or uFlag = "Yes" Or uFlag = "yes" Or uFlag = "1" Or uFlag = "True" Then
    iStep = Val(InputBox("What is the first readable step in the time trace? (1, 2, 3...):", "1st step?"))
    
    V = ActiveSheet.Cells(4, 12)
    ActiveSheet.Cells(4, 12) = V * 1000
    
    i = iStep
    While ActiveSheet.Cells(4 + i, 12) <> 0
        V = ActiveSheet.Cells(4 + i, 12)
        ActiveSheet.Cells(4 + i, 12) = V * 1000
        i = i + 1
    Wend
End If

'Loop for determine diffs' cutoff value and [mA] Current:
 i = 2
 While ActiveSheet.Cells(i, 2) <> 0
  ActiveSheet.Cells(i, 3) = ActiveSheet.Cells(i, 2) * 1000 'Current [A] to [mA]
  i = i + 1
 Wend

'Loop for determine Baseline-substracted Current [mA]:
 i = 2
 While ActiveSheet.Cells(i, 2) <> 0
  ActiveSheet.Cells(i, 4) = ActiveSheet.Cells(i, 3) - ActiveSheet.Cells(4, 12)
  i = i + 1
 Wend

End Sub

'=========================================='

Sub RCalcId()

j = 30 'Initialise reading variable (at a reasonable value where sweep started)

'Read data from the recording form (Id-Vd)
Dsteps = ActiveSheet.Cells(2, 1) 'The number of Vd steps specified in the form
For i = 1 To Dsteps
 Id1 = ActiveSheet.Cells(j, 2)
 Id2 = ActiveSheet.Cells(j + 1, 2)
 diff = Abs(Id2 - Id1) 'Difference between previous and next current
 
 While diff < 0.0001 'While no step, continue reading
  Id1 = ActiveSheet.Cells(j, 2)
  Id2 = ActiveSheet.Cells(j + 1, 2)
  diff = Abs(Id2 - Id1)
  j = j + 1
 Wend
 
 'Set the cells' range for calculating the average of currents:
 Set avRange = ActiveSheet.Range(ActiveSheet.Cells(j - 21, 2), ActiveSheet.Cells(j - 1, 2))
 'Write data to the plot cells:
 ActiveSheet.Cells(17 + i, 5) = Application.WorksheetFunction.Average(avRange)
 
 j = j + 30 'In order to ensure next step
 
Next i

End Sub

'=========================================='

Sub CCalcId()

'=================================================================================
'If you HAVE NOT added the analytes at the right time (NOT CONSTANT TIME INTERVAL)
'=================================================================================

'Determine the parameters as a function of the System State
If ActiveSheet.Cells(2, 13) = "QSS" Then 'Quasi Solid-State'
   Scoff = 0.3 'The cutoff correction factor for the steps
   Sstep = 53 'The increment in t for calculating the next step
ElseIf ActiveSheet.Cells(2, 13) = "LS" Then 'Liquid-State'
   Scoff = 0.1 'The cutoff correction factor for the steps
   Sstep = 23 'The increment in t for calculating the next step
End If

'=============CREATE A RANGE TO PLOT EACH TIME TRACE TO COMPARE=================
        Dim wb As Workbook
        Dim ws As Worksheet
        Dim rname As String 'Declare the name variable as a String
        Set wb = ActiveWorkbook
        Set ws = ActiveSheet

            'Read the time between additions (each "?" cells)
            addt = ActiveSheet.Cells(1, 13)

        rStart = ActiveSheet.Cells(1, 12) - addt ' "-addt" in order to capture before the first step
        
        'As-measured Current:
        rname = Mid(ws.Name, 1, InStr(1, ws.Name, "(") - 1) & Mid(ws.Name, InStr(1, ws.Name, "(") + 1, 1)
        wb.Worksheets(ws.Name).Names.Add Name:=rname, RefersTo:=ActiveSheet.Range(ActiveSheet.Cells(rStart, 3), ActiveSheet.Cells(16000, 3)) ' Column#2 [A]; Column#3 [mA]
        
        'Baseline-substracted Current:
        rname = "d" & Mid(ws.Name, 1, InStr(1, ws.Name, "(") - 1) & Mid(ws.Name, InStr(1, ws.Name, "(") + 1, 1)
        wb.Worksheets(ws.Name).Names.Add Name:=rname, RefersTo:=ActiveSheet.Range(ActiveSheet.Cells(rStart, 4), ActiveSheet.Cells(16000, 4)) ' Column#3 [mA]; Column#4 dI [mA]
'===============================================================================

'As "non-visible" steps are not readable:
iStep = Val(InputBox("What is the first readable step in the time trace? (1, 2, 3...):", "1st step?"))

'In the case you waited (or missed to add) for a longer time in a step:
spStep = Val(InputBox("Any special step? In affirmative case, write the step number (otherwise, type 0):", "Special step?")) '[!!!!!!]
If spStep <> 0 Then
 spaddt = Val(InputBox("Type the duration time for this step (s):", "Special addition time (s)"))
End If

'Read the time between additions (each "?" cells)
addt = ActiveSheet.Cells(1, 13)

'=====================BASELINE CURRENT DETERMINATION========================
'Calculate the average for the Io (stable current before the first addition)
 j = ActiveSheet.Cells(1, 12) 'Set the variable at the time for the first addition
 'Set the cells' range for calculating the average of currents:
  Set avRange = ActiveSheet.Range(ActiveSheet.Cells(j - 22, 2), ActiveSheet.Cells(j - 2, 2))
 'Write data:
  ActiveSheet.Cells(4, 12) = Application.WorksheetFunction.Average(avRange) * 1000 ''*1000' to change A to mA
'===========================================================================

'=====================SENS. CURRENT CALCULATION LOOP========================
'k = iStep
For k = iStep To 3 'In the case the initial steps correspond to the very first, the calculation for the subsequent may not be accurate
'Because 1st step is not always readible
j = ActiveSheet.Cells(1, 12) + (k * addt - 20) 'Initialise reading variable (at a reasonable right before the next step)

'Loop for determine diffs' cutoff value and [mA] Current:
 i = 2
 While ActiveSheet.Cells(i, 2) <> 0
  ActiveSheet.Cells(i, 1) = Abs(ActiveSheet.Cells(i, 2) - ActiveSheet.Cells(i + 1, 2))
  ActiveSheet.Cells(i, 3) = ActiveSheet.Cells(i, 2) * 1000 'Current [A] to [mA]
  i = i + 1
 Wend

 'Set the cells' range for calculating the cutoff value for the diffs:
 Set coffRange = ActiveSheet.Range(ActiveSheet.Cells(j, 1), ActiveSheet.Cells(j + (addt - 250), 1))
 cutoff = WorksheetFunction.Max(coffRange) - Scoff * (WorksheetFunction.Max(coffRange)) 'In order to reduce a little the max diff
 ActiveSheet.Columns(1).EntireColumn.ClearContents 'Clear the generated diffs column

    '=====================DEBUGGING=======================
    'Write cutoff value in the sheet (debugging)
    ActiveSheet.Cells(1, 1) = "Cutoff:"
    ActiveSheet.Cells(2, 1) = cutoff
    '=====================================================

'Read data from the recording form (the active sheet)
j = j + 6 'For getting closer to the 1st step
nSteps = ActiveSheet.Cells(2, 12) 'The number of additions specified in the form

errflag = 0 'The flag to control the overflow error
For i = k To nSteps
 Id1 = ActiveSheet.Cells(j, 2)
 Id2 = ActiveSheet.Cells(j + 1, 2)
 diff = Abs(Id2 - Id1) 'Difference between previous and next current

 Do While diff < cutoff 'While no step, continue reading
  Id1 = ActiveSheet.Cells(j, 2)
  Id2 = ActiveSheet.Cells(j + 1, 2)
  diff = Abs(Id2 - Id1)
  j = j + 1
  If j = 1000000 Then 'In the case the overflow error appears for the reading variable ("j") and the considered "cutoff"
   errflag = 1
   Exit Do
  End If
 Loop

 If errflag = 1 Then Exit For 'In the case the overflow error appears for the reading variable ("j") and the considered "cutoff"
 
 'Set the cells' range for calculating the average of currents:
 Set avRange = ActiveSheet.Range(ActiveSheet.Cells(j - 21, 2), ActiveSheet.Cells(j - 1, 2))
 'Write data:
 ActiveSheet.Cells(4 + i, 12) = Application.WorksheetFunction.Average(avRange) * 1000 ''*1000' to change A to mA

 If (i + 1) = spStep Then 'Considering the special step
  j = j + (spaddt - Sstep) 'To ensure next diff is a step (for the specified special step)
 Else
 j = j + (addt - Sstep) 'To ensure next diff is a step (each "?" defined cells)
 End If

Next i
Next k
'===========================================================================

'Loop for determine Baseline-substracted Current [mA]:
 i = 2
 While ActiveSheet.Cells(i, 2) <> 0
  ActiveSheet.Cells(i, 4) = ActiveSheet.Cells(i, 3) - ActiveSheet.Cells(4, 12)
  i = i + 1
 Wend

End Sub

'=========================================='

Sub ctCCalcId()

'=========================================================================
'IF YOU HAVE ADDED THE ANALYTES AT THE RIGHT TIME (CONSTANT TIME INTERVAL)
'=========================================================================

'Read the time for the 1st addition (cells)
add0 = ActiveSheet.Cells(1, 12)
'Read the time between additions (each "?" cells)
addt = ActiveSheet.Cells(1, 13)

'Because 1st step is not always readible:
j = add0 + (addt - 3) 'Initialise reading variable (right before the 2nd sweep started)

'Read data from the recording form (the active sheet)
Dsteps = ActiveSheet.Cells(2, 12) 'The number of additions specified in the form
For i = 1 To Dstep

 'Set the cells' range for calculating the average of currents:
 Set avRange = ActiveSheet.Range(ActiveSheet.Cells(j - 21, 2), ActiveSheet.Cells(j - 1, 2))

 'Write data:
 ActiveSheet.Cells(3 + i, 12) = Application.WorksheetFunction.Average(avRange) * 1000 ''*1000' to change A to mA
 j = add0 + (i + 1) * addt - 3 'Going right before the next step (restarting "j" to the previous step)

Next i

End Sub

'=========================================='

Sub DCalc()

ActiveSheet.Cells(3, 18) = "Drift (µA/min)" '

'Set the reading variable at the time for the first addition
j0 = ActiveSheet.Cells(1, 12)
'Read the time between additions (each "?" cells)
addt = ActiveSheet.Cells(1, 13)

'Starting at the first big step delivered (same as for response time):
iStep = Val(InputBox("What is the first reasonably big step in the time trace? (1, 2, 3...):", "1st step?"))

'In the case you waited (or missed to add) for a longer time in a step:
spStep = Val(InputBox("Any special step? In affirmative case, write the step number (otherwise, type 0):", "Special step?")) '[!!!!!!]
If spStep <> 0 Then
 spaddt = Val(InputBox("Type the duration time for this step (s):", "Special addition time (s)"))
End If

'===================== BASELINE DRIFT CALCULATION ========================
j = ActiveSheet.Cells(1, 12) 'Set the variable at the time for the first addition
Id1 = ActiveSheet.Cells(j - 5, 3)
Id2 = ActiveSheet.Cells(j - 5 - Application.WorksheetFunction.Round(0.8 * addt, 0), 3)
dt = ActiveSheet.Cells(j - 5, 1) - ActiveSheet.Cells(j - 5 - Application.WorksheetFunction.Round(0.8 * addt, 0), 1) 'Time interval in [s]
drift = Abs((Id2 - Id1) / dt) 'Drift for the baseline step
'Write data:
ActiveSheet.Cells(4, 18) = drift * 1000 * 60 ''*1000*60' to change [mA/s] to [µA/min]
'=========================================================================

'=====================SENS. DRIFT CALCULATION LOOP========================
i = iStep
If addt < 400 Then Jadj = 2 Else Jadj = 5 'Set proper back-adjustment for reading variable to calculation
While ActiveSheet.Cells(4 + i, 12) <> 0
 
    j = ActiveSheet.Cells(4 + i, 19)
    If i >= spStep And spStep <> 0 Then 'Considering the special step
        jf = j0 + (i - 1) * addt + spaddt
    Else
        jf = j0 + i * addt
    End If
    
    Id1 = ActiveSheet.Cells(j - Jadj, 3)
    Id2 = ActiveSheet.Cells(jf - Jadj, 3)
    dt = ActiveSheet.Cells(jf - Jadj, 1) - ActiveSheet.Cells(j - Jadj, 1)
   
    drift = Abs((Id2 - Id1) / dt) 'Drift for the current step
 
    'Write data:
    ActiveSheet.Cells(4 + i, 18) = drift * 1000 * 60 ''*1000*60' to change [mA/s] to [µA/min]
 
    'DEBUGGING:
    ActiveSheet.Cells(4 + i, 20) = jf
    
    i = i + 1
Wend
'===========================================================================

End Sub

'=========================================='

Sub old_gmCalcId()

j = 30 'Initialise reading variable (at a reasonable value where sweep started)

'Loop for determine diffs' cutoff value:
 i = 5
 While ActiveSheet.Cells(i, 2) <> 0
  ActiveSheet.Cells(i, 1) = Abs(ActiveSheet.Cells(i, 2) - ActiveSheet.Cells(i + 1, 2))
  i = i + 1
 Wend

 'Set the cells' range for calculating the cutoff value for the diffs:
 Set coffRange = ActiveSheet.Range(ActiveSheet.Cells(50, 1), ActiveSheet.Cells(150, 1))
 cutoff = WorksheetFunction.Max(coffRange) - 0.45 * WorksheetFunction.Max(coffRange) 'In order to reduce a little the max diff
 ActiveSheet.Columns(1).Range(Cells(5, 1), Cells(i, 1)).ClearContents 'Clear the generated diffs column


'Read data from the recording form (Id-Vd)
Dsteps = ActiveSheet.Cells(2, 1) 'The number of Vd steps specified in the form
Gsteps = ActiveSheet.Cells(4, 1) 'The number of Vg steps specified in the form
w = 17 'The start and counting writing variable

For k = 1 To Dsteps
 For i = 1 To Gsteps
  Id1 = ActiveSheet.Cells(j, 2)
  Id2 = ActiveSheet.Cells(j + 1, 2)
  diff = Abs(Id2 - Id1) 'Difference between previous and next current

  While diff < cutoff 'While no step, continue reading
   Id1 = ActiveSheet.Cells(j, 2)
   Id2 = ActiveSheet.Cells(j + 1, 2)
   diff = Abs(Id2 - Id1)
   j = j + 1
  Wend

  'Set the cells' range for calculating the average of currents:
  Set avRange = ActiveSheet.Range(ActiveSheet.Cells(j - 21, 2), ActiveSheet.Cells(j - 1, 2))
  'Write data to the plot cells:
  ActiveSheet.Cells(w + i, 7) = Application.WorksheetFunction.Average(avRange)

  'If i = (Gsteps - 1) Then 'Because the last step lasts 200s instead of 100s
   'j = j + 170 'In order to ensure next step
  'Else
   j = j + 80 'In order to ensure next step
  'End If

 Next i
 w = w + i - 1 'To start the next loop right where it ended (before "Next i")

Next k

End Sub

'=========================================='

Sub gmCalcId_v3()

Dim matchRange As String 'Declare the variable that will contain the range in which the max diff will be found by the MATCH Excel function

'Initialise reading variable (at a reasonable value where sweep started, usually = ~205s for the end of the 0V step):
t0 = ActiveSheet.Cells(6, 6)
j = t0

'=========================== DIFFS' VALUE LOOP =======================================
 i = 5 'Skip the cells for manually entering steps' data
 While ActiveSheet.Cells(i, 2) <> 0
  ActiveSheet.Cells(i, 1) = Abs(ActiveSheet.Cells(i, 2) - ActiveSheet.Cells(i + 1, 2))
  i = i + 1
 Wend
'=====================================================================================
  
'Read data from the recording form (gmCalc)
Dsteps = ActiveSheet.Cells(2, 1) 'The number of Vd steps specified in the form
Gsteps = ActiveSheet.Cells(4, 1) 'The number of Vg steps specified in the form
w = 17 'The start-and-counting writing variable

If ActiveSheet.Cells(18, 6) = 0 Then 'Check for the first Vg value considered
    oFlag = True
Else
    oFlag = False
End If

tDstep = ActiveSheet.Cells(9, 6) - 7 'The duration (# of cells) value for a regular Vd step (~1000; "-7" to compensate human delay before sweeping)
tStep = ActiveSheet.Cells(10, 6) 'The duration (# of cells) value for each normal Vg step
For k = 1 To Dsteps
 For i = 1 To Gsteps

  If i = Gsteps Then 'To jump to next Vd step at Vg = 0.10V
   j = k * tDstep
  Else
   j = j + Application.WorksheetFunction.Round(0.899 * tStep, 0) 'In order to ensure next step (usually ~103)
  End If
  
  'ActiveSheet.Cells(11 + i, 20 + k) = j 'DEBUGGING
'================================ MAX DIFF | Vg STEP =================================
  'Set the cells' range for calculating the max value for the diffs | each Vg step:
  t0coff = j - 30 'The inferior time value for the first Vg step considered for the cutoff
  tfcoff = j + 40 'The superior time value for the first Vg step considered for the cutoff
  Set coffRange = ActiveSheet.Range(ActiveSheet.Cells(t0coff, 1), ActiveSheet.Cells(tfcoff, 1))
  matchRange = "A" & t0coff & ":A" & tfcoff
  j = Application.WorksheetFunction.Match(WorksheetFunction.Max(coffRange), ActiveSheet.Range(matchRange), 0) + t0coff - 6
  '"+ t0coff" because MATCH gives relative position from the start of the range, "-(5+1)" to reduce a little the max diff position ("1" because of "+1" included by "t0coff")
  ActiveSheet.Cells(i, 20 + k) = j + 5 'DEBUGGING
'=====================================================================================
 
  'Set the cells' range for calculating the average of currents:
  If i = Gsteps Then
   Set avRange = ActiveSheet.Range(ActiveSheet.Cells(j - 21, 2), ActiveSheet.Cells(j - 1, 2))
  Else
   Set avRange = ActiveSheet.Range(ActiveSheet.Cells(j - 21, 2), ActiveSheet.Cells(j - 1, 2))
  End If
  'Write data to the plot cells:
  ActiveSheet.Cells(w + i, 7) = Application.WorksheetFunction.Average(avRange)
  
 Next i
 w = w + i - 1 'To start the next loop right where it ended (before "Next i")
 If oFlag = False Then j = j + Application.WorksheetFunction.Round(0.8 * t0, 0) 'In order to overcome the 0 Vg step within the next Vd step

Next k

'Clear the generated diffs column
'ActiveSheet.Columns(1).Range(Cells(5, 1), Cells(1000000, 1)).ClearContents '(Comment for debbuging)

End Sub

'=========================================='

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

'=========================================='

Sub tRCalc()
   
ActiveSheet.Cells(1, 1) = "Time (s)"
If (ActiveSheet.Cells(2, 1) <> 0 And Mid(ActiveSheet.Cells(2, 1), 3, 1) = ":") Then 'Correct time only if needed
'=============================== CORRECTED TIME CALCULATION ==================================
Dim raw_time As String 'Declare the variable for storaging the time string as recorded

'For the very first sampling time:
raw_time = ActiveSheet.Cells(2, 1) 'Store the as-recorded time string
If Mid(raw_time, 9, 1) = ":" Then
    corr_time = Left(raw_time, 8) & "." & Right(raw_time, 3) 'Rewrite as a proper time string
ElseIf Right(raw_time, 1) = ":" Then
    corr_time = Left(raw_time, 10) & "00" 'Rewrite as a proper time string
End If
    
t0_HH = Left(corr_time, 2) * 3600 'Convert to seconds
t0_MM = Mid(corr_time, 4, 2) * 60 'Convert to seconds
t0_SS = Right(corr_time, 6) 'Already in seconds
t0 = Val(t0_HH) + Val(t0_MM) + Val(t0_SS) 'Save the value for the very first sampling time
ActiveSheet.Cells(2, 1) = 0 'Set the zero for the time in seconds
ActiveSheet.Cells(2, 1).NumberFormat = "0.00" 'Set the cell format as float number

i = 3
While ActiveSheet.Cells(i, 1) <> 0
    
    raw_time = ActiveSheet.Cells(i, 1) 'Store the as-recorded time string

    If Mid(raw_time, 9, 1) = ":" Then
        corr_time = Left(raw_time, 8) & "." & Right(raw_time, 3) 'Rewrite as a proper time string
    ElseIf Right(raw_time, 1) = ":" Then
        corr_time = Left(raw_time, 10) & "00" 'Rewrite as a proper time string
    End If
    
    t1_HH = Left(corr_time, 2) * 3600 'Convert to seconds
    t1_MM = Mid(corr_time, 4, 2) * 60 'Convert to seconds
    t1_SS = Right(corr_time, 6) 'Already in seconds
    t1 = Val(t1_HH) + Val(t1_MM) + Val(t1_SS) 'Save the next (current) sampling time
    ActiveSheet.Cells(i, 1).NumberFormat = "0.00" 'Set the cell format as float number
    ActiveSheet.Cells(i, 1) = (t1 - t0) 'Calculate the time difference and write in seconds
    
    i = i + 1
Wend
'=============================================================================================
End If

'================================ RESPONSE TIME CALCULATION ==================================
Worksheets(ActiveSheet.Name).Cells(3, 17) = "Resp. time (s)"
Worksheets(ActiveSheet.Name).Cells(3, 16) = "Tot. R. (mA)"

'Starting at the first big step delivered:
iStep = Val(InputBox("What is the first reasonably big step in the time trace? (1, 2, 3...):", "1st step?"))


'In the case you waited (or missed to add) for a longer time in a step:
spStep = Val(InputBox("Any special step? In affirmative case, write the step number (otherwise, type 0):", "Special step?")) '[!!!!!!]
If spStep <> 0 Then
 spaddt = Val(InputBox("Type the duration time for this step (s):", "Special addition time (s)"))
End If

'Set the reading variable at the time for the first addition
j0 = Worksheets(ActiveSheet.Name).Cells(1, 12)
'Read the time between additions (each "?" cells)
addt = Worksheets(ActiveSheet.Name).Cells(1, 13)

    '====================== CALCULATION LOOOP ===================
    If addt < 400 Then coffPCENT = 0.95 Else coffPCENT = 0.9 'Set appropriate cutoff % to max tot. resp.
    i = iStep
    While Worksheets(ActiveSheet.Name).Cells(4 + i, 12) <> 0
    
        'Set the value for the time trace reading variable and initial time for each add.:
        If i >= spStep And spStep <> 0 Then 'Considering the special step
            j = j0 + (i - 2) * addt + spaddt + 1 '"+1" because of the title row
        Else
            j = j0 + (i - 1) * addt + 1 '"+1" because of the title row
        End If
        t0 = ActiveSheet.Cells(j, 1)
        
        'Calculate the tot. resp. for the selected addition (based on avg.) and the cutoff (90~95%) value for resp. time:
        dI = Abs(Worksheets(ActiveSheet.Name).Cells(3 + i, 12) - Worksheets(ActiveSheet.Name).Cells(4 + i, 12))
        ActiveSheet.Cells(4 + i, 16) = dI 'Write the value for the Tot. Resp. (mA)
        dI = coffPCENT * dI
        
        'Look for "stable" currents above the cutoff value:
        dI1 = Abs(Worksheets(ActiveSheet.Name).Cells(3 + i, 12) - ActiveSheet.Cells(j, 3))
        dI5 = Abs(Worksheets(ActiveSheet.Name).Cells(3 + i, 12) - ActiveSheet.Cells(j + 4, 3))
        While dI1 < dI Or dI5 < dI
            j = j + 1
            dI1 = Abs(Worksheets(ActiveSheet.Name).Cells(3 + i, 12) - ActiveSheet.Cells(j, 3))
            dI5 = Abs(Worksheets(ActiveSheet.Name).Cells(3 + i, 12) - ActiveSheet.Cells(j + 4, 3))
        Wend
        
        'Calculate and write the value of response time:
        ActiveSheet.Cells(4 + i, 17) = ActiveSheet.Cells(j, 1) - t0
        
        'DEBUGGING and Cells info for Drift calc.:
        ActiveSheet.Cells(4 + i, 19) = j
        
        i = i + 1
    Wend
    '============================================================
'=============================================================================================

End Sub
