Attribute VB_Name = "Project2VaR"
Option Explicit

' FRE 6251 Project II. All money amounts are USD.
' A lower-tail P&L quantile is negative; reported loss VaR is its negative.
Private rngState As Double
Private spareNormal As Double
Private hasSpare As Boolean
Private Const PM_MOD As Double = 2147483647#
Private Const PI As Double = 3.1415926535897931

' This module is self-contained: import into a blank .xlsm and run RunProject2.
' VBA creates the input/output sheets, calculates every result and draws the chart.
Private Function ProjectSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set ProjectSheet = ws
End Function

Private Sub PutRow(ByVal ws As Worksheet, ByVal address As String, ByVal items As Variant)
    Dim j As Long
    For j = LBound(items) To UBound(items)
        ws.Range(address).Offset(0, j - LBound(items)).Value = items(j)
    Next j
End Sub

Private Sub HeaderStyle(ByVal cells As Range)
    With cells
        .Interior.Color = RGB(36, 59, 83)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 32
    End With
End Sub

Private Sub PrepareProject2()
    Dim inp As Worksheet, res As Worksheet, runs As Worksheet, checks As Worksheet
    Dim ws As Worksheet, names As Variant, labels As Variant, name As Variant, i As Long
    Set res = ProjectSheet("Results")
    Set inp = ProjectSheet("Inputs")
    Set runs = ProjectSheet("Runs")
    Set checks = ProjectSheet("Validation")
    names = Array("Results", "Inputs", "Runs", "Validation")
    For Each name In names
        Set ws = ThisWorkbook.Worksheets(CStr(name))
        With ws.Range("A1:M50")
            .Font.Name = "Arial"
            .Font.Size = 10
            .RowHeight = 22
            .VerticalAlignment = xlCenter
        End With
        ws.Columns("A:M").ColumnWidth = 18
        ws.Range("A2").Font.Size = 16
        ws.Range("A2").Font.Bold = True
        ws.Range("A2").Font.Color = RGB(36, 59, 83)
    Next name
    res.Columns("A").ColumnWidth = 29
    res.Columns("D").ColumnWidth = 22
    inp.Columns("A").ColumnWidth = 32
    checks.Columns("A").ColumnWidth = 40
    checks.Columns("B").ColumnWidth = 24
    checks.Columns("C").ColumnWidth = 25
    checks.Columns("D").ColumnWidth = 66
    If Len(CStr(inp.Range("A6").Value2)) = 0 Then
        PutRow inp, "A6", Array("C", 47#, 0.4, 100000&)
        PutRow inp, "A7", Array("GS", 144#, 0.45, -40000&)
        PutRow inp, "A8", Array("MSFT", 31#, 0.3, 10000&)
        PutRow inp, "A11", Array("Horizon in years", 1# / 252#)
        PutRow inp, "A12", Array("Lower-tail probability", 0.01)
        PutRow inp, "A13", Array("Repetitions per N", 20&)
        PutRow inp, "A14", Array("Base random seed", 20261019#)
        PutRow inp, "A19", Array("C", 1#, 0.6, 0.2)
        PutRow inp, "A20", Array("GS", 0.6, 1#, 0.4)
        PutRow inp, "A21", Array("MSFT", 0.2, 0.4, 1#)
    End If
    inp.Range("A2").Value = "Portfolio inputs and VBA analytic calculation"
    inp.Range("A4").Value = "Source: Project II Revised Yet Again.pdf; Notes about the projects.docx, Project II."
    PutRow inp, "A5", Array("Ticker", "USD price / share", "Annual volatility", "Shares held", "Market value", "Annual USD exposure")
    HeaderStyle inp.Range("A5:F5")
    inp.Range("B6:D8").Font.Color = vbBlue
    inp.Range("B11:B14").Font.Color = vbBlue
    inp.Range("B19:D21").Font.Color = vbBlue
    inp.Range("B6:B8").NumberFormat = "0.00"
    inp.Range("C6:C8").NumberFormat = "0.0%"
    inp.Range("D6:D8").NumberFormat = "#,##0;(#,##0)"
    inp.Range("E6:F8").NumberFormat = "#,##0.00;(#,##0.00)"
    inp.Range("B11").NumberFormat = "0.00000000"
    inp.Range("B12").NumberFormat = "0.0%"
    inp.Range("B13:B14").NumberFormat = "0"
    inp.Range("A16").Value = "Edit blue inputs, then run RunProject2. All calculations are performed in VBA."
    PutRow inp, "A18", Array("Correlation matrix", "C", "GS", "MSFT")
    PutRow inp, "A24", Array("VBA Cholesky factor L", "C", "GS", "MSFT")
    For i = 0 To 2: inp.Cells(25 + i, 1).Value = inp.Cells(6 + i, 1).Value2: Next i
    inp.Range("B19:D21").NumberFormat = "0.00"
    inp.Range("B25:D27").NumberFormat = "0.00000000"
    inp.Range("F24").Value = "Daily covariance contributions (USD^2)"
    inp.Range("G25:I27").NumberFormat = "#,##0"
    inp.Range("A30").Value = "Initial portfolio value (USD)"
    inp.Range("A31").Value = "Daily variance (USD^2)"
    inp.Range("A32").Value = "Daily standard deviation (USD)"
    inp.Range("A33").Value = "Analytic P&L 1% quantile"
    inp.Range("A34").Value = "Standard-normal 1% quantile"
    inp.Range("A35").Value = "Density at quantile (1/USD)"
    inp.Range("A36").Value = "MSE coefficient C (USD^2)"
    inp.Range("B30:B33").NumberFormat = "#,##0.00;(#,##0.00)"
    inp.Range("B34").NumberFormat = "0.00000000"
    inp.Range("B35:B36").NumberFormat = "0.000000E+00"
    inp.Range("J4").Value = "Last simulation input snapshot"
    PutRow inp, "J5", Array("Price", "Volatility", "Shares")
    inp.Range("J10").Value = "Horizon / tail / runs / seed"
    inp.Range("J18").Value = "Correlation snapshot"
    HeaderStyle inp.Range("J5:L5")

    res.Range("A2").Value = "Project II  Monte Carlo VaR in VBA"
    res.Range("A4").Value = "Analytic loss VaR (USD)"
    res.Range("A5").Value = "Daily P&L stdev (USD)"
    res.Range("D4").Value = "1% P&L quantile (USD)"
    res.Range("D5").Value = "Net initial value (USD)"
    res.Range("B4:B5").NumberFormat = "#,##0.00;(#,##0.00)"
    res.Range("F4:F5").NumberFormat = "#,##0.00;(#,##0.00)"
    PutRow res, "A8", Array("Scenarios N", "Rank k", "Mean loss VaR", "Bias of P&L q", "MSE (USD^2)", "RMSE (USD)", "Theory MSE", "MSE / theory")
    HeaderStyle res.Range("A8:H8")
    res.Range("A9:B13").NumberFormat = "#,##0"
    res.Range("C9:G13").NumberFormat = "#,##0.00;(#,##0.00)"
    res.Range("H9:H13").NumberFormat = "0.000"
    res.Range("A15").Value = "VBA: MSE = average of 20 squared quantile errors; RMSE = Sqr(MSE)."
    res.Range("A17").Value = "Log-log MSE slope"
    res.Range("B17").NumberFormat = "0.000"
    res.Range("D17").Value = "Theory: MSE slope -1; RMSE slope -0.5."
    res.Range("A35").Value = "Run information"
    labels = Array("Execution status", "Base seed", "Generated (local time)", "Run time (seconds)", "Saved analytic P&L q", "Simulation inputs", "Calculation engine")
    For i = 0 To 6: res.Cells(36 + i, 1).Value = labels(i): Next i
    res.Range("B38").NumberFormat = "yyyy-mm-dd hh:mm"
    res.Range("B39:B40").NumberFormat = "#,##0.00;(#,##0.00)"
    res.Range("A44").Value = "RunProject2 creates the sheets, runs all 100 estimates, calculates summaries and draws the log-log chart."
    res.Range("A45").Value = "No Python, external data or precomputed simulation results are needed."
    res.Range("A46").Value = "Edit blue inputs on Inputs and rerun. Numerical results update only when the VBA macro runs."
    res.Range("A47").Value = "To export the completed Results sheet as PDF, run ExportProject2PDF after RunProject2."

    runs.Range("A2").Value = "Individual VBA simulation estimates"
    runs.Range("A4").Value = "Each row is calculated in VBA; error = signed quantile minus analytic benchmark."
    PutRow runs, "A5", Array("Scenarios N", "Repetition", "Starting RNG state", "P&L quantile", "Loss VaR", "Signed error", "Squared error", "Ending RNG state")
    HeaderStyle runs.Range("A5:H5")
    runs.Range("A6:H105").Font.Name = "Arial"
    runs.Range("A6:H105").Font.Size = 10
    runs.Range("D6:G105").NumberFormat = "#,##0.00;(#,##0.00)"
    runs.Range("A6:C105").NumberFormat = "0"
    runs.Range("H6:H105").NumberFormat = "0"

    checks.Range("A2").Value = "VBA algorithm validation"
    checks.Range("A4").Value = "RunProject2 executes every check below in VBA before generating the 100 VaR estimates."
    PutRow checks, "A5", Array("Check", "Observed", "Expected / tolerance", "Meaning")
    HeaderStyle checks.Range("A5:D5")
    PutRow checks, "A6", Array("Max absolute L L^T residual", "", "<= 1E-12", "Reconstructs the supplied correlation matrix")
    PutRow checks, "A7", Array("Park-Miller state at step 10,000", "", 1043618065#, "Initial state = 1")
    PutRow checks, "A8", Array("Random quicksort ordered", "", True, "10,000 generated normal values")
    PutRow checks, "A9", Array("Sort sum difference", "", "<= 1E-8", "Preservation diagnostic")
    PutRow checks, "A10", Array("Sort sum of squares difference", "", "<= 1E-6", "Preservation diagnostic")
    PutRow checks, "A11", Array("Duplicates and negative values", "", True, "Expected: -1, -1, 0, 3, 3, 9")
    PutRow checks, "A12", Array("Reverse, sorted, singleton arrays", "", True, "Exact equality to expected array")
    PutRow checks, "A13", Array("1% order statistic of 1..1,000", "", 10, "One-based rank k = 10")
    PutRow checks, "A14", Array("Normal generator sample mean", "", 0, "100,000 draws; statistical diagnostic")
    PutRow checks, "A15", Array("Normal generator sample variance", "", 1, "100,000 draws; statistical diagnostic")
    PutRow checks, "A17", Array("Validation engine", "Not run")
    PutRow checks, "A19", Array("Independent sort comparison", "", "Exact match", "Insertion sort versus quicksort on 128 random values")
    checks.Range("A21").Value = "VBA correlation diagnostic"
    checks.Range("A22").Value = "100,000 generated correlated vectors; separate seed 97531."
    PutRow checks, "A23", Array("Ticker", "C", "GS", "MSFT")
    HeaderStyle checks.Range("A23:D23")
    For i = 0 To 2: checks.Cells(24 + i, 1).Value = inp.Cells(6 + i, 1).Value2: Next i
    checks.Range("B24:D26").ClearContents
    checks.Range("B24:D26").NumberFormat = "0.0000"
    checks.Range("B6").NumberFormat = "0.000E+00"
    checks.Range("B9:B10").NumberFormat = "0.000E+00"
    checks.Range("B14:B15").NumberFormat = "0.000000"
    checks.Range("A28").Value = "Quicksort is checked against an independently implemented insertion sort."
    checks.Range("A29").Value = "Statistical moments and correlations are diagnostics, not exact pass/fail identities."
    ApplyCompactLayout
End Sub

Private Sub TableBorders(ByVal cells As Range)
    With cells.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 192, 203)
    End With
End Sub

' Keep assignment outputs first; retain execution metadata out of the main view.
Private Sub ApplyCompactLayout()
    Dim res As Worksheet, inp As Worksheet, runs As Worksheet, checks As Worksheet
    Dim i As Long, g As Long, r As Long, sizes As Variant
    Set res = ThisWorkbook.Worksheets("Results")
    Set inp = ThisWorkbook.Worksheets("Inputs")
    Set runs = ThisWorkbook.Worksheets("Runs")
    Set checks = ThisWorkbook.Worksheets("Validation")
    runs.Move After:=res
    res.Range("A2").Value = "Project II  Monte Carlo VaR"
    res.Range("A2").Font.Size = 14
    res.Columns("A").ColumnWidth = 29
    res.Columns("C").ColumnWidth = 21
    res.Columns("E").ColumnWidth = 22
    res.Columns("G").ColumnWidth = 22
    res.Columns("B").Hidden = True
    res.Columns("D").Hidden = True
    res.Columns("F").Hidden = True
    res.Columns("H").Hidden = True
    res.Rows("4:5").Hidden = True
    res.Rows("15:18").Hidden = True
    res.Rows("35:49").Hidden = True
    res.Rows("57:65").Hidden = True
    res.Range("A6").Value = "Analytic loss VaR (USD)"
    res.Range("C6").Formula = "=$B$4"
    res.Range("C6").NumberFormat = "#,##0.00;(#,##0.00)"
    res.Range("A6:H6").Font.Bold = True
    res.Range("A6:H6").Interior.Color = RGB(233, 239, 245)
    res.Range("A8").Value = "Scenarios N"
    res.Range("C8").Value = "Mean loss VaR (USD)"
    res.Range("E8").Value = "MSE (USD^2)"
    res.Range("G8").Value = "Theory MSE (USD^2)"
    res.Rows("8:8").RowHeight = 34
    TableBorders res.Range("A6:C6")
    TableBorders res.Range("A8:H13")

    runs.Columns("A:H").Hidden = True
    runs.Columns("I:J").ColumnWidth = 3
    runs.Columns("K").ColumnWidth = 8
    runs.Columns("L:P").ColumnWidth = 19
    runs.Range("K2:P25").Font.Name = "Arial"
    runs.Range("K2:P25").Font.Size = 10
    runs.Range("K2").Value = "Individual 1% P&L quantiles (USD)"
    runs.Range("K2").Font.Bold = True
    runs.Range("K2").Font.Size = 14
    runs.Range("K5").Value = "Run"
    sizes = Array(1000&, 5000&, 10000&, 50000&, 100000&)
    For g = 0 To 4
        runs.Cells(5, g + 12).Value = sizes(g)
        For i = 1 To 20
            r = 6 + g * 20 + i - 1
            runs.Cells(i + 5, g + 12).Formula = "=$D$" & r
        Next i
    Next g
    For i = 1 To 20: runs.Cells(i + 5, 11).Value = i: Next i
    HeaderStyle runs.Range("K5:P5")
    runs.Range("L5:P5").NumberFormat = "#,##0"
    runs.Range("K6:K25").NumberFormat = "0"
    runs.Range("L6:P25").NumberFormat = "#,##0.00;(#,##0.00)"
    TableBorders runs.Range("K5:P25")

    inp.Range("A2").Value = "Inputs"
    inp.Range("A2").Font.Size = 14
    inp.Range("A4").Value = "Source: Project II Revised Yet Again.pdf"
    inp.Range("A16").Value = "Edit inputs and run RunProject2 to recalculate."
    inp.Columns("E:L").Hidden = True
    inp.Rows("30:36").Hidden = True
    TableBorders inp.Range("A5:D8")
    TableBorders inp.Range("A11:B14")
    TableBorders inp.Range("A18:D21")
    TableBorders inp.Range("A24:D27")
    TableBorders checks.Range("A5:D15")
    TableBorders checks.Range("A19:D19")
    TableBorders checks.Range("A23:D26")
End Sub

Public Sub RunProject2()
    Dim inp As Worksheet, runs As Worksheet, result As Worksheet
    Dim corr() As Double, lower() As Double, exposure(1 To 3) As Double
    Dim sizes As Variant, pnl() As Double, observations(1 To 100, 1 To 8) As Variant
    Dim n As Long, i As Long, j As Long, g As Long, trial As Long, outRow As Long
    Dim x1 As Double, x2 As Double, x3 As Double
    Dim z1 As Double, z2 As Double, z3 As Double, dt As Double, p As Double
    Dim oldScreen As Boolean, oldCalc As XlCalculation, started As Double
    Dim v As Double, sigma As Double, q As Double, startSeed As Double
    Dim baseSeed As Double, runCount As Long, msg As String, oldCancel As XlEnableCancelKey
    Dim initialValue As Double, normalQ As Double, density As Double, coefficient As Double
    Dim summary(1 To 5, 1 To 8) As Variant
    Dim errorSum As Double, errorSqSum As Double, quantileSum As Double, estimate As Double
    Dim logN(1 To 5) As Double, logMSE(1 To 5) As Double, rootDt As Double
    On Error GoTo Failed
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldCancel = Application.EnableCancelKey
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    PrepareProject2
    Set inp = ThisWorkbook.Worksheets("Inputs")
    Set runs = ThisWorkbook.Worksheets("Runs")
    Set result = ThisWorkbook.Worksheets("Results")
    result.Range("B36").Value = "Running"
    result.Range("A9:H13").ClearContents
    result.Range("B17").ClearContents
    result.Range("K5:M10").ClearContents
    runs.Range("A6:H105").ClearContents
    dt = inp.Range("B11").Value2
    p = inp.Range("B12").Value2
    runCount = inp.Range("B13").Value2
    baseSeed = inp.Range("B14").Value2
    If dt <= 0 Or p <= 0 Or p >= 0.5 Or runCount <> 20 Then
        Err.Raise 5, , "Use a positive horizon, a tail probability between 0 and 0.5, and 20 runs."
    End If
    If baseSeed < 1 Or baseSeed >= PM_MOD Or baseSeed <> Fix(baseSeed) Then
        Err.Raise 5, , "Seed must be an integer between 1 and 2147483646."
    End If
    ReDim corr(1 To 3, 1 To 3)
    For i = 1 To 3
        If inp.Cells(i + 5, 2).Value2 <= 0 Or inp.Cells(i + 5, 3).Value2 < 0 Then
            Err.Raise 5, , "Prices must be positive and volatilities nonnegative."
        End If
        exposure(i) = inp.Cells(i + 5, 2).Value2 * inp.Cells(i + 5, 3).Value2 _
                    * inp.Cells(i + 5, 4).Value2
        inp.Cells(i + 5, 5).Value = inp.Cells(i + 5, 2).Value2 * inp.Cells(i + 5, 4).Value2
        inp.Cells(i + 5, 6).Value = exposure(i)
        initialValue = initialValue + inp.Cells(i + 5, 5).Value2
        For j = 1 To 3
            corr(i, j) = inp.Cells(i + 18, j + 1).Value2
        Next j
    Next i
    Cholesky corr, lower
    For i = 1 To 3
        For j = 1 To 3
            v = v + exposure(i) * corr(i, j) * exposure(j) * dt
            inp.Cells(i + 24, j + 6).Value = exposure(i) * corr(i, j) * exposure(j) * dt
            inp.Cells(i + 24, j + 1).Value = lower(i, j)
        Next j
    Next i
    If v <= 0 Then Err.Raise 5, , "Portfolio variance must be positive."
    sigma = Sqr(v)
    normalQ = Application.WorksheetFunction.Norm_S_Inv(p)
    q = normalQ * sigma
    density = Exp(-0.5 * normalQ * normalQ) / Sqr(2# * PI) / sigma
    coefficient = p * (1# - p) / (density * density)
    inp.Range("B30").Value = initialValue
    inp.Range("B31").Value = v
    inp.Range("B32").Value = sigma
    inp.Range("B33").Value = q
    inp.Range("B34").Value = normalQ
    inp.Range("B35").Value = density
    inp.Range("B36").Value = coefficient
    result.Range("B4").Value = -q
    result.Range("B5").Value = sigma
    result.Range("F4").Value = q
    result.Range("F5").Value = initialValue
    ValidateProject2
    SetSeed baseSeed
    sizes = Array(1000&, 5000&, 10000&, 50000&, 100000&)
    started = Timer
    rootDt = Sqr(dt)
    For g = 0 To 4
        n = sizes(g)
        errorSum = 0#: errorSqSum = 0#: quantileSum = 0#
        For trial = 1 To runCount
            Application.StatusBar = "Project II: N=" & n & ", run " & trial & "/20"
            startSeed = rngState
            ReDim pnl(1 To n)
            For i = 1 To n
                x1 = StandardNormal(): x2 = StandardNormal(): x3 = StandardNormal()
                z1 = lower(1, 1) * x1
                z2 = lower(2, 1) * x1 + lower(2, 2) * x2
                z3 = lower(3, 1) * x1 + lower(3, 2) * x2 + lower(3, 3) * x3
                pnl(i) = rootDt * (exposure(1) * z1 + exposure(2) * z2 + exposure(3) * z3)
            Next i
            QuickSort pnl, 1, n
            outRow = outRow + 1
            observations(outRow, 1) = n
            observations(outRow, 2) = trial
            observations(outRow, 3) = startSeed
            observations(outRow, 4) = LowerQuantile(pnl, p)
            observations(outRow, 5) = -observations(outRow, 4)
            observations(outRow, 6) = observations(outRow, 4) - q
            observations(outRow, 7) = observations(outRow, 6) ^ 2
            observations(outRow, 8) = rngState
            estimate = observations(outRow, 4)
            quantileSum = quantileSum + estimate
            errorSum = errorSum + (estimate - q)
            errorSqSum = errorSqSum + (estimate - q) ^ 2
        Next trial
        summary(g + 1, 1) = n
        summary(g + 1, 2) = -Int(-Round(p * n, 9))
        summary(g + 1, 3) = -quantileSum / runCount
        summary(g + 1, 4) = errorSum / runCount
        summary(g + 1, 5) = errorSqSum / runCount
        summary(g + 1, 6) = Sqr(errorSqSum / runCount)
        summary(g + 1, 7) = coefficient / n
        summary(g + 1, 8) = summary(g + 1, 5) / summary(g + 1, 7)
        logN(g + 1) = Log(CDbl(n)) / Log(10#)
        logMSE(g + 1) = Log(CDbl(summary(g + 1, 5))) / Log(10#)
    Next g
    runs.Range("A6:H105").Value = observations
    result.Range("A9:H13").Value = summary
    result.Range("B17").Value = RegressionSlope(logN, logMSE)
    CreateConvergenceChart result
    WriteProjectDiscussion result, q, CDbl(result.Range("B17").Value2)
    result.Range("B37").Value = baseSeed
    result.Range("B38").Value = Now
    result.Range("B39").Value = ElapsedSeconds(started)
    result.Range("B40").Value = q
    result.Range("B42").Value = "Excel VBA"
    inp.Range("J6:L8").Value = inp.Range("B6:D8").Value
    inp.Range("J11:J14").Value = inp.Range("B11:B14").Value
    inp.Range("J19:L21").Value = inp.Range("B19:D21").Value
    SetFreshnessFormula result
    result.Range("B36").Value = "Complete - VBA executed"
    Application.Calculate
    Application.StatusBar = False
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableCancelKey = oldCancel
    If Len(ThisWorkbook.Path) > 0 Then ThisWorkbook.Save
    Exit Sub
Failed:
    msg = Err.Description
    On Error Resume Next
    result.Range("B36").Value = "FAILED: " & msg
    Application.StatusBar = False
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableCancelKey = oldCancel
    On Error GoTo 0
    Err.Raise vbObjectError + 201, "RunProject2", msg
End Sub

Private Function RegressionSlope(ByRef x() As Double, ByRef y() As Double) As Double
    Dim n As Long, i As Long, sx As Double, sy As Double, sxy As Double, sxx As Double
    n = UBound(x)
    For i = 1 To n
        sx = sx + x(i): sy = sy + y(i)
        sxy = sxy + x(i) * y(i): sxx = sxx + x(i) * x(i)
    Next i
    RegressionSlope = (n * sxy - sx * sy) / (n * sxx - sx * sx)
End Function

Private Sub SetFreshnessFormula(ByVal ws As Worksheet)
    Dim parts As String, r As Long, j As Long, cols As Variant, snap As Variant
    cols = Array("B", "C", "D"): snap = Array("J", "K", "L")
    For r = 6 To 8
        For j = 0 To 2
            parts = parts & "Inputs!" & cols(j) & r & "=Inputs!" & snap(j) & r & ","
        Next j
    Next r
    For r = 11 To 14: parts = parts & "Inputs!B" & r & "=Inputs!J" & r & ",": Next r
    For r = 19 To 21
        For j = 0 To 2
            parts = parts & "Inputs!" & cols(j) & r & "=Inputs!" & snap(j) & r & ","
        Next j
    Next r
    parts = Left$(parts, Len(parts) - 1)
    ws.Range("B41").Formula = "=IF(AND(" & parts & "),""Current"",""STALE - rerun RunProject2"")"
End Sub

Private Sub CreateConvergenceChart(ByVal ws As Worksheet)
    Dim box As ChartObject, series As Series, i As Long
    ws.Range("A19:H34").ClearContents
    For i = ws.ChartObjects.Count To 1 Step -1: ws.ChartObjects(i).Delete: Next i
    Set box = ws.ChartObjects.Add(ws.Range("A19").Left, ws.Range("A19").Top, _
              ws.Range("A19:H34").Width, ws.Range("A19:H34").Height)
    box.Name = "Project2Convergence"
    With box.Chart
        .ChartType = xlXYScatterLines
        .HasTitle = True
        .ChartTitle.Text = "Monte Carlo MSE convergence"
        Set series = .SeriesCollection.NewSeries
        series.Name = "Empirical MSE (20 VBA runs)"
        series.XValues = ws.Range("A9:A13")
        series.Values = ws.Range("E9:E13")
        series.Format.Line.ForeColor.RGB = RGB(23, 99, 166)
        Set series = .SeriesCollection.NewSeries
        series.Name = "Asymptotic theory C / N"
        series.XValues = ws.Range("A9:A13")
        series.Values = ws.Range("G9:G13")
        series.Format.Line.ForeColor.RGB = RGB(138, 108, 58)
        series.Format.Line.DashStyle = 4
        .Axes(xlCategory).ScaleType = xlLogarithmic
        .Axes(xlCategory).MinimumScale = 1000#
        .Axes(xlCategory).MaximumScale = 100000#
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Scenarios N (log scale)"
        .Axes(xlValue).ScaleType = xlLogarithmic
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "MSE in USD^2 (log scale)"
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop
    End With
End Sub

Private Sub WriteProjectDiscussion(ByVal ws As Worksheet, ByVal q As Double, ByVal slope As Double)
    Dim areas As Variant, paragraphs As Variant, i As Long
    ws.Range("A50:H65").UnMerge
    ws.Range("A50:H65").ClearContents
    ws.Range("A50").Value = "Answers"
    ws.Range("A50").Font.Bold = True
    areas = Array("A52:H53", "A55:H56")
    paragraphs = Array( _
        "Does the error agree with the formula? The fitted log-log MSE slope is " & Format(slope, "0.000") & _
        ", close to -1. This agrees with MSE approximately proportional to 1/N, allowing for fluctuation from 20 runs.", _
        "Why does the logarithm base not matter in O(N log N)? Since log_b(N) = ln(N)/ln(b), changing the base only changes a constant factor, which does not affect the order of complexity.")
    For i = 0 To 1
        With ws.Range(CStr(areas(i)))
            .Merge
            .Cells(1, 1).Value = paragraphs(i)
            .WrapText = True
            .VerticalAlignment = xlCenter
            .Font.Name = "Arial"
            .Font.Size = 10
            .RowHeight = 27
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(180, 192, 203)
        End With
    Next i
End Sub

Public Sub ExportProject2PDF()
    Dim ws As Worksheet, filePath As String
    Set ws = ThisWorkbook.Worksheets("Results")
    If ws.Range("B36").Value2 <> "Complete - VBA executed" Then
        Err.Raise 5, , "Run RunProject2 successfully before exporting the PDF."
    End If
    If ws.Range("B41").Value2 <> "Current" Then Err.Raise 5, , "Inputs changed; rerun before exporting."
    If Len(ThisWorkbook.Path) = 0 Then Err.Raise 5, , "Save the workbook before exporting."
    filePath = ThisWorkbook.Path & Application.PathSeparator & "Project II VBA Results.pdf"
    With ws.PageSetup
        .PrintArea = "$A$2:$H$65"
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False
End Sub

' Lower triangular L with L L^T = correlation. No Excel formula shortcut.
Public Sub Cholesky(ByRef a() As Double, ByRef lower() As Double)
    Dim n As Long, i As Long, j As Long, k As Long, s As Double
    n = UBound(a, 1)
    If LBound(a, 1) <> 1 Or LBound(a, 2) <> 1 Or UBound(a, 2) <> n Then Err.Raise 5
    ReDim lower(1 To n, 1 To n)
    For i = 1 To n
        If Abs(a(i, i) - 1#) > 0.000000000001 Then Err.Raise 5, , "Correlation diagonal must equal one."
        For j = 1 To n
            If Abs(a(i, j) - a(j, i)) > 0.000000000001 Or Abs(a(i, j)) > 1# Then
                Err.Raise 5, , "Correlation must be symmetric and entries must be between -1 and 1."
            End If
        Next j
        For j = 1 To i
            s = a(i, j)
            For k = 1 To j - 1
                s = s - lower(i, k) * lower(j, k)
            Next k
            If i = j Then
                If s <= 0# Then Err.Raise 5, , "Correlation must be positive definite."
                lower(i, j) = Sqr(s)
            Else
                lower(i, j) = s / lower(j, j)
            End If
        Next j
    Next i
End Sub

' In-place quicksort. Recurse only into the smaller partition, bounding
' recursion depth even if partitions are unbalanced. Average O(N log N).
Public Sub QuickSort(ByRef a() As Double, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long, pivot As Double, temp As Double
    Do While first < last
        i = first: j = last
        pivot = a(first + (last - first) \ 2)
        Do
            Do While a(i) < pivot: i = i + 1: Loop
            Do While a(j) > pivot: j = j - 1: Loop
            If i <= j Then
                temp = a(i): a(i) = a(j): a(j) = temp
                i = i + 1: j = j - 1
            End If
        Loop While i <= j
        If j - first < last - i Then
            If first < j Then QuickSort a, first, j
            first = i
        Else
            If i < last Then QuickSort a, i, last
            last = j
        End If
    Loop
End Sub

Public Function LowerQuantile(ByRef sortedValues() As Double, ByVal p As Double) As Double
    Dim n As Long, k As Long, rank As Double
    n = UBound(sortedValues) - LBound(sortedValues) + 1
    If p <= 0 Or p >= 1 Or n < 1 Then Err.Raise 5
    rank = p * n
    ' Remove floating point noise at an exact integer; otherwise use ceiling.
    If Abs(rank - Round(rank, 0)) < 0.000000001 Then rank = Round(rank, 0)
    k = -Int(-rank)
    LowerQuantile = sortedValues(LBound(sortedValues) + k - 1)
End Function

Private Sub SetSeed(ByVal seed As Double)
    rngState = seed
    hasSpare = False
End Sub

' Park-Miller minimal-standard uniform generator. Double arithmetic is
' exact for these integer products (< 2^53); states are in 1..m-1.
' Unlike RND(), this uniform is strictly inside (0,1), so Log(0) cannot occur.
Private Function UniformOpen() As Double
    Dim product As Double
    product = 16807# * rngState
    rngState = product - Fix(product / PM_MOD) * PM_MOD
    UniformOpen = rngState / PM_MOD
End Function

' Box-Muller returns independent N(0,1) pairs and caches the second value.
Private Function StandardNormal() As Double
    Dim radius As Double, angle As Double
    If hasSpare Then
        hasSpare = False
        StandardNormal = spareNormal
    Else
        radius = Sqr(-2# * Log(UniformOpen()))
        angle = 2# * PI * UniformOpen()
        StandardNormal = radius * Cos(angle)
        spareNormal = radius * Sin(angle)
        hasSpare = True
    End If
End Function

Private Function ElapsedSeconds(ByVal started As Double) As Double
    ElapsedSeconds = Timer - started
    If ElapsedSeconds < 0 Then ElapsedSeconds = ElapsedSeconds + 86400#
End Function

Public Sub ValidateProject2()
    Dim ws As Worksheet, a() As Double, lower() As Double, x() As Double
    Dim i As Long, j As Long, k As Long, s As Double, residual As Double
    Dim beforeSum As Double, afterSum As Double, beforeSq As Double, afterSq As Double
    Dim ordered As Boolean, fixed(1 To 6) As Double, normalSum As Double, normalSq As Double
    Dim independent(1 To 128) As Double, held As Double
    Dim sums(1 To 3) As Double, products(1 To 3, 1 To 3) As Double
    Dim normals(1 To 3) As Double, correlated(1 To 3) As Double
    Dim cov As Double, varianceI As Double, varianceJ As Double
    Set ws = ThisWorkbook.Worksheets("Validation")
    ReDim a(1 To 3, 1 To 3)
    For i = 1 To 3
        For j = 1 To 3
            a(i, j) = ThisWorkbook.Worksheets("Inputs").Cells(i + 18, j + 1).Value2
        Next j
    Next i
    Cholesky a, lower
    For i = 1 To 3
        For j = 1 To 3
            s = 0
            For k = 1 To 3: s = s + lower(i, k) * lower(j, k): Next k
            If Abs(s - a(i, j)) > residual Then residual = Abs(s - a(i, j))
        Next j
    Next i
    ws.Range("B6").Value = residual
    If residual > 0.000000000001 Then Err.Raise 5, , "Cholesky reconstruction test failed."
    SetSeed 1
    For i = 1 To 10000: s = UniformOpen(): Next i
    ws.Range("B7").Value = rngState
    If rngState <> 1043618065# Then Err.Raise 5, , "Random generator reference test failed."
    SetSeed 13579
    ReDim x(1 To 10000)
    For i = 1 To 10000
        x(i) = StandardNormal()
        beforeSum = beforeSum + x(i): beforeSq = beforeSq + x(i) * x(i)
    Next i
    QuickSort x, 1, 10000
    ordered = True
    For i = 1 To 10000
        afterSum = afterSum + x(i): afterSq = afterSq + x(i) * x(i)
        If i > 1 Then
            If x(i) < x(i - 1) Then ordered = False
        End If
    Next i
    ws.Range("B8").Value = ordered
    ws.Range("B9").Value = Abs(beforeSum - afterSum)
    ws.Range("B10").Value = Abs(beforeSq - afterSq)
    If Not ordered Or Abs(beforeSum - afterSum) > 0.00000001 Or Abs(beforeSq - afterSq) > 0.000001 Then
        Err.Raise 5, , "Random-array quicksort test failed."
    End If
    fixed(1) = 3: fixed(2) = -1: fixed(3) = 3
    fixed(4) = 0: fixed(5) = -1: fixed(6) = 9
    QuickSort fixed, 1, 6
    ordered = (fixed(1) = -1 And fixed(2) = -1 And fixed(3) = 0 And fixed(4) = 3 And fixed(5) = 3 And fixed(6) = 9)
    ws.Range("B11").Value = ordered
    If Not ordered Then Err.Raise 5, , "Duplicate quicksort test failed."
    ' Reverse sorted, already sorted and singleton cases.
    ReDim x(1 To 1000)
    For i = 1 To 1000: x(i) = 1001 - i: Next i
    QuickSort x, 1, 1000: QuickSort x, 1, 1000
    For i = 1 To 1000
        If x(i) <> i Then Err.Raise 5, , "Sorted/reversed quicksort test failed."
    Next i
    QuickSort x, 1, 1
    ws.Range("B12").Value = True
    ws.Range("B13").Value = LowerQuantile(x, 0.01)
    If ws.Range("B13").Value <> 10 Then Err.Raise 5, , "Quantile indexing test failed."
    SetSeed 24680
    For i = 1 To 100000
        s = StandardNormal(): normalSum = normalSum + s: normalSq = normalSq + s * s
    Next i
    ws.Range("B14").Value = normalSum / 100000#
    ws.Range("B15").Value = (normalSq - normalSum * normalSum / 100000#) / 99999#
    SetSeed 86420
    ReDim x(1 To 128)
    For i = 1 To 128
        x(i) = StandardNormal(): independent(i) = x(i)
    Next i
    QuickSort x, 1, 128
    ' Independent insertion sort supplies an exact reference for a small sample.
    For i = 2 To 128
        held = independent(i): j = i - 1
        Do While j >= 1
            If independent(j) <= held Then Exit Do
            independent(j + 1) = independent(j): j = j - 1
        Loop
        independent(j + 1) = held
    Next i
    For i = 1 To 128
        If x(i) <> independent(i) Then Err.Raise 5, , "Independent sort comparison failed."
    Next i
    ws.Range("B19").Value = True
    SetSeed 97531
    For k = 1 To 100000
        For i = 1 To 3: normals(i) = StandardNormal(): Next i
        For i = 1 To 3
            correlated(i) = 0#
            For j = 1 To i: correlated(i) = correlated(i) + lower(i, j) * normals(j): Next j
            sums(i) = sums(i) + correlated(i)
        Next i
        For i = 1 To 3
            For j = 1 To 3
                products(i, j) = products(i, j) + correlated(i) * correlated(j)
            Next j
        Next i
    Next k
    For i = 1 To 3
        For j = 1 To 3
            cov = products(i, j) - sums(i) * sums(j) / 100000#
            varianceI = products(i, i) - sums(i) ^ 2 / 100000#
            varianceJ = products(j, j) - sums(j) ^ 2 / 100000#
            ws.Cells(i + 23, j + 1).Value = cov / Sqr(varianceI * varianceJ)
        Next j
    Next i
    ws.Range("B17").Value = "VBA executed"
End Sub
