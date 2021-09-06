Attribute VB_Name = "plotModule"
'=====================CREDITS======================'
'AUTHOR: Andres Alberto Andreo Acosta'
'GitHub: https://github.com/andriandreo'
'DATE (DD/MM/YY): 06/09/21'
'Version: v2.0'
'LICENSE: MIT License'
'================================================='

Dim V As String 'Series based on different "Vg"?

Sub plotUpdate()
 V = InputBox("Are the tests performed for several Vg? (Yes/No):", "Plot Series as Vg?")
 plotXY
 plotXnrY
 plotXdY
 plotCtt
 plotCC
 rtable
End Sub

Sub plotXY()
 n = Val(Mid(ActiveSheet.Name, InStr(1, ActiveSheet.Name, "(") + 1, 1)) 'Read the number of plot series based on the name of the "CCalc" Sheet
 
 ActiveSheet.ChartObjects("IdChart").Activate
 ActiveChart.SeriesCollection(1).Select
  ActiveChart.SeriesCollection(1).Name = "LIN"
  ActiveChart.SeriesCollection(1).XValues = "='CCalc(" & n & ")'!xLIN_" & n
  ActiveChart.SeriesCollection(1).Values = "='CCalc(" & n & ")'!yLIN_" & n
 ActiveChart.SeriesCollection(2).Select
  ActiveChart.SeriesCollection(2).Name = "CalcId"
End Sub

Sub plotXnrY()
 n = Val(Mid(ActiveSheet.Name, InStr(1, ActiveSheet.Name, "(") + 1, 1)) 'Read the number of plot series based on the name of the "CCalc" Sheet
 
 ActiveSheet.ChartObjects("NRChart").Activate
 ActiveChart.SeriesCollection(1).Select
  ActiveChart.SeriesCollection(1).Name = "LIN"
  ActiveChart.SeriesCollection(1).XValues = "='CCalc(" & n & ")'!xLIN_" & n
  ActiveChart.SeriesCollection(1).Values = "='CCalc(" & n & ")'!nrLIN_" & n
 ActiveChart.SeriesCollection(2).Select
  ActiveChart.SeriesCollection(2).Name = "CalcId"
End Sub

Sub plotXdY()
 n = Val(Mid(ActiveSheet.Name, InStr(1, ActiveSheet.Name, "(") + 1, 1)) 'Read the number of plot series based on the name of the "CCalc" Sheet
 
 ActiveSheet.ChartObjects("dIChart").Activate
 ActiveChart.SeriesCollection(1).Select
  ActiveChart.SeriesCollection(1).Name = "LIN"
  ActiveChart.SeriesCollection(1).XValues = "='CCalc(" & n & ")'!xLIN_" & n
  ActiveChart.SeriesCollection(1).Values = "='CCalc(" & n & ")'!dyLIN_" & n
 ActiveChart.SeriesCollection(2).Select
  ActiveChart.SeriesCollection(2).Name = "CalcId"
End Sub

Sub plotCtt()
 'Dim V As String
 'V = InputBox("Are the tests performed for several Vg? (Yes/No):", "Plot Series as Vg?")
 n = Val(Mid(ActiveSheet.Name, InStr(1, ActiveSheet.Name, "(") + 1, 1)) 'Read the number of plot series based on the name of the "CCalc" Sheet
 
 ActiveSheet.ChartObjects("CttChart").Activate
 M = ActiveChart.SeriesCollection.Count 'Count the total number of Series within the chart
 If V = "Yes" Or V = "Y" Or V = "yes" Or V = "y" Or V = "1" Then
  ActiveChart.ChartTitle.Text = "Time traces for " & n & " consecutive tests at different Vg"
  i = M
  Do
   ActiveChart.SeriesCollection(i).Select
   ActiveChart.SeriesCollection(i).Name = "0." & (M - i + 1) & "0 V"
   i = i - 1
  Loop Until i = 0
 End If
 
 If M > n Then
  For i = 1 To (M - n)
  ActiveChart.FullSeriesCollection(i).Select
   ActiveChart.FullSeriesCollection(i).IsFiltered = True
  Next i
 End If
 
 M = ActiveChart.SeriesCollection.Count 'Count the number of Series again after hiding the excess (Use FullSeriesCollection to count the actual total)
 i = M
 Do
   ActiveChart.SeriesCollection(i).Select
    ActiveChart.SeriesCollection(i).Values = "='CCalc(" & (M - i + 1) & ")'!CCalc" & (M - i + 1)
   i = i - 1
 Loop Until i = 0
End Sub

Sub plotCC()
 'Dim V As String
 'V = InputBox("Are the tests performed for several Vg? (Yes/No):", "Plot Series as Vg?")
 n = Val(Mid(ActiveSheet.Name, InStr(1, ActiveSheet.Name, "(") + 1, 1)) 'Read the number of plot series based on the name of the "CCalc" Sheet
 
 ActiveSheet.ChartObjects("CCChart").Activate
 M = ActiveChart.SeriesCollection.Count / 2 'Count the total number of Series within the chart ("=/2" because they are the double)
 If V = "Yes" Or V = "Y" Or V = "yes" Or V = "y" Or V = "1" Then
  i = M
  Do
   ActiveChart.SeriesCollection(2 * i - 1).Select
   ActiveChart.SeriesCollection(2 * i - 1).Name = "LIN(0." & (M - i + 1) & "0 V)"
   ActiveChart.SeriesCollection(2 * i).Select
   ActiveChart.SeriesCollection(2 * i).Name = "0." & (M - i + 1) & "0 V"
   i = i - 1
  Loop Until i = 0
 End If
 
 If M > n Then
  For i = 1 To (M - n)
   ActiveChart.FullSeriesCollection(2 * i - 1).Select
    ActiveChart.FullSeriesCollection(2 * i - 1).IsFiltered = True
   ActiveChart.FullSeriesCollection(2 * i).Select
    ActiveChart.FullSeriesCollection(2 * i).IsFiltered = True
  Next i
 End If
 
 M = ActiveChart.SeriesCollection.Count / 2 'Count the total number of Series again after hiding the excess ("=/2" because they are the double)
 i = M
 Do
   ActiveChart.SeriesCollection(2 * i - 1).Select
    ActiveChart.SeriesCollection(2 * i - 1).XValues = "='CCalc(" & (M - i + 1) & ")'!xLIN_" & (M - i + 1)
    ActiveChart.SeriesCollection(2 * i - 1).Values = "='CCalc(" & (M - i + 1) & ")'!yLIN_" & (M - i + 1)
   ActiveChart.SeriesCollection(2 * i).Select
    ActiveChart.SeriesCollection(2 * i).XValues = "='CCalc(" & (M - i + 1) & ")'!$K$5:$K$14"
    ActiveChart.SeriesCollection(2 * i).Values = "='CCalc(" & (M - i + 1) & ")'!$L$5:$L$14"
   i = i - 1
  Loop Until i = 0

End Sub

Sub rtable()

'Declare two variables as matrices for registering the regression stats (and the y, x ranges), respectively:
Dim rstats() As Variant
Dim Y, nrY, dY As Range
Dim X As Range

j = 0 'In the case "V" is negative
If V = "Yes" Or V = "Y" Or V = "yes" Or V = "y" Or V = "1" Then
 ActiveSheet.Cells(57, 4) = "Vg (V)"
 ActiveSheet.Cells(57, 5) = "N-Sens (A/logM/V)"
 ActiveSheet.Cells(57, 6) = "Sens (A/logM)"
 ActiveSheet.Cells(57, 7) = "Sens (log-1M)"
  ActiveSheet.Cells(57, 7).Characters(Start:=10, Length:=2).Font.Superscript = True
 ActiveSheet.Cells(57, 8) = "R2"
  ActiveSheet.Cells(57, 8).Characters(Start:=2, Length:=1).Font.Superscript = True
 ActiveSheet.Cells(57, 9) = "Lin. Range (logM)"
 
 j = 1 'To move 1 column to the right for writing the remaining data
End If

n = Val(Mid(ActiveSheet.Name, InStr(1, ActiveSheet.Name, "(") + 1, 1)) 'Read the number of plot series based on the name of the "CCalc" Sheet
 For i = 1 To n
   Set X = Worksheets("CCalc(" & i & ")").Range("xLIN_" & i)
   Set Y = Worksheets("CCalc(" & i & ")").Range("yLIN_" & i)
   Set nrY = Worksheets("CCalc(" & i & ")").Range("nrLIN_" & i)
   Set dY = Worksheets("CCalc(" & i & ")").Range("dyLIN_" & i)
  
  'Calculate a 2x5 matrix with the linear regression and its stats (Non-normalised Response):
   rstats() = Application.WorksheetFunction.LinEst(Y, X, , True)
   
  'Write the results within the table
   M = X.Count 'To count the elements within the "X" range
   ActiveSheet.Cells(66 - i, 8 + j) = "[" & X(1, 1) & ", " & X(M, 1) & "]" 'To write the Linear Range
    ActiveSheet.Cells(66 - i, 8 + j).HorizontalAlignment = xlCenter 'To format (center) the data of the "Lin. Range" column
   ActiveSheet.Cells(66 - i, 7 + j) = rstats(3, 1) 'To write the "R2" coefficient
    ActiveSheet.Cells(66 - i, 7 + j).NumberFormat = "0.0000" 'To format the data in the "R2" column
    ActiveSheet.Cells(66 - i, 7 + j).HorizontalAlignment = xlRight 'To formar (right) the data of the "R2" column
   ActiveSheet.Cells(66 - i, 5 + j) = rstats(1, 1) 'To write the value for the slope (Sensitivity)
   
  'Calculate a 2x5 matrix with the linear regression and its stats (Normalised Response):
   rstats() = Application.WorksheetFunction.LinEst(nrY, X, , True)
   'Write the sensibility for this case:
    ActiveSheet.Cells(66 - i, 6 + j) = rstats(1, 1) 'To write the value for the slope (Normalised Sens.)
     ActiveSheet.Cells(66 - i, 6 + j).NumberFormat = "0.0000E+00" 'To format the data in the "Sens" column
     
  'In the case the series correspond to several "Vg":
  If j = 1 Then
    ActiveSheet.Cells(66 - i, 4) = i / 10 'In order to extract the actual "Vg" value based on the "CCalc" Sheet number
    ActiveSheet.Cells(66 - i, 5).Formula = "=F" & (66 - i) & "/D" & (66 - i) 'To write the value for the "Vg"-Normalised  Sensitivity ("N-Sens")
  End If
 Next i
 
End Sub
