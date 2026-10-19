Attribute VB_Name = "Project1ImpVol"
Option Explicit
Option Base 0

' American put, no dividends. Volatilities are decimal fractions.
' The optional arguments link the numerical controls to the template.
Public Function AmericanPutImpVol(MktPrice As Double, S As Double, T As Double, _
        X As Double, r As Double, Tol As Double, _
        Optional TimeSteps As Long = 100, Optional PriceSteps As Long = 200, _
        Optional MinPrice As Double = 0, Optional MaxPrice As Double = 100) As Double

    Dim nt As Long, nLeft As Long, nRight As Long, iteration As Long
    Dim dx As Double, europeanLimit As Double
    Dim x0 As Double, x1 As Double, x2 As Double
    Dim f0 As Double, f1 As Double, f2 As Double
    Dim lowerVol As Double, upperVol As Double, intrinsic As Double

    intrinsic = PositivePart(X - S)
    If MktPrice <= intrinsic Or MktPrice >= X Then
        Err.Raise 5, "AmericanPutImpVol", "The option price must be above intrinsic value and below the strike."
    End If
    If S <= 0 Or T <= 0 Or X <= 0 Or Tol <= 0 Or r < 0 Then Err.Raise 5
    If MinPrice < 0 Or MinPrice >= S Or MaxPrice <= S Then Err.Raise 5
    If TimeSteps < 1 Or PriceSteps < 2 Then Err.Raise 5

    ' At equal volatility the American put is at least as expensive as
    ' the European put. Use its implied volatility as the initial upper
    ' bound when it exists. European prices cannot reach X * Exp(-r*T),
    ' whereas a valid American price can exceed that limit.
    europeanLimit = X * Exp(-r * T)
    If MktPrice < europeanLimit Then
        x1 = impPutVol(MktPrice, S, T, X, r)
    Else
        x1 = 0.5
    End If

    ' Establish a bracket in numerical prices, including the case where
    ' no European implied volatility exists and small finite-grid errors.
    For iteration = 1 To 20
        MakeLogGrid S, T, X, r, x1, TimeSteps, PriceSteps, MinPrice, MaxPrice, _
                    nt, nLeft, nRight, dx
        f1 = AmericanPutCN(S, x1, T, X, r, nt, nLeft, nRight, dx) - MktPrice
        If f1 >= 0 Then Exit For
        If -f1 < 0.01 Then
            x1 = x1 * 1.01 + Tol
        Else
            x1 = x1 * 2
        End If
    Next iteration
    If f1 < 0 Then Err.Raise 5, "AmericanPutImpVol", "Unable to bracket the numerical root."
    x0 = 0
    f0 = intrinsic - MktPrice
    lowerVol = x0
    upperVol = x1

    For iteration = 1 To 100
        If f1 = f0 Then
            x2 = (lowerVol + upperVol) / 2
        Else
            x2 = x1 - f1 * (x1 - x0) / (f1 - f0)
        End If
        ' A bisection step keeps the secant iterate inside the valid bounds.
        If x2 <= lowerVol Or x2 >= upperVol Then x2 = (lowerVol + upperVol) / 2
        f2 = AmericanPutCN(S, x2, T, X, r, nt, nLeft, nRight, dx) - MktPrice
        ' The additional price check avoids stopping on the nearly flat bend.
        If Abs(x2 - x1) < Tol And Abs(f2) < 0.00000001 Then
            AmericanPutImpVol = Round(x2, 4)
            Exit Function
        End If
        If f2 < 0 Then
            lowerVol = x2
        Else
            upperVol = x2
        End If
        x0 = x1
        f0 = f1
        x1 = x2
        f1 = f2
    Next iteration
    Err.Raise 5, "AmericanPutImpVol", "The secant iteration did not converge."
End Function

' Fixed log-price grid for one implied-volatility solve. Both the spot
' and strike are nodes, so the current option value needs no interpolation.
' Six standard deviations plus drift cover the numerical domain. The
' template's step counts and price bounds are minimum coverage controls.
Private Sub MakeLogGrid(S As Double, T As Double, X As Double, r As Double, _
        upperVol As Double, TimeSteps As Long, PriceSteps As Long, _
        MinPrice As Double, MaxPrice As Double, ByRef nt As Long, _
        ByRef nLeft As Long, ByRef nRight As Long, ByRef dx As Double)

    Dim sd As Double, mu As Double, leftSpan As Double, rightSpan As Double
    Dim strikeGap As Double, minSteps As Long
    sd = upperVol * Sqr(T)
    mu = r - upperVol * upperVol / 2
    strikeGap = Abs(Log(X / S))
    leftSpan = 6 * sd + PositivePart(-mu * T) + strikeGap
    rightSpan = 6 * sd + PositivePart(r * T) + strikeGap
    If MinPrice > 0 Then
        If Log(S / MinPrice) > leftSpan Then leftSpan = Log(S / MinPrice)
    End If
    If Log(MaxPrice / S) > rightSpan Then rightSpan = Log(MaxPrice / S)

    minSteps = PriceSteps
    If minSteps < 3200 Then minSteps = 3200
    dx = (leftSpan + rightSpan) / minSteps
    If dx > 0.004 Then dx = 0.004
    If dx > sd / 150 Then dx = sd / 150
    If strikeGap > 0.0000000001 Then
        dx = strikeGap / CeilingPositive(strikeGap / dx)
    End If
    nLeft = CeilingPositive(leftSpan / dx)
    nRight = CeilingPositive(rightSpan / dx)
    nt = TimeSteps
    If nt < 1600 Then nt = 1600
    If nt < CeilingPositive(200 * T) Then nt = CeilingPositive(200 * T)
End Sub

' Crank-Nicolson in x = Log(S): V_tau = sigma^2/2 V_xx
' + (r - sigma^2/2) V_x - r V. Reverse tridiagonal elimination and
' forward projected substitution enforce early exercise at every step.
' Four implicit half steps replace the first two CN steps to smooth
' the payoff kink (Rannacher start). Their implicit matrix is the same
' as the remaining CN steps, so its factorization is reused.
Public Function AmericanPutCN(S As Double, sigma As Double, T As Double, _
        X As Double, r As Double, nt As Long, nLeft As Long, _
        nRight As Long, dx As Double) As Double

    Dim v() As Double, payoff() As Double, pivot() As Double
    Dim multiplier() As Double, rhs() As Double
    Dim dt As Double, mu As Double, a As Double, b As Double, c As Double
    Dim spot As Double, z As Double
    Dim i As Long, timeStep As Long, ns As Long

    If sigma = 0 Then
        AmericanPutCN = PositivePart(X - S)
        Exit Function
    End If
    ns = nLeft + nRight
    dt = T / nt
    mu = r - sigma * sigma / 2
    a = dt / 4 * (sigma * sigma / (dx * dx) - mu / dx)
    b = -dt / 2 * (sigma * sigma / (dx * dx) + r)
    c = dt / 4 * (sigma * sigma / (dx * dx) + mu / dx)
    ReDim v(0 To ns), payoff(0 To ns)
    ReDim pivot(1 To ns - 1), multiplier(1 To ns - 1), rhs(1 To ns - 1)

    ' Terminal payoff and American boundary values.
    For i = 0 To ns
        spot = S * Exp((i - nLeft) * dx)
        payoff(i) = PositivePart(X - spot)
        v(i) = payoff(i)
    Next i
    ' M has diagonals -a, 1-b, -c; N has diagonals a, 1+b, c.
    pivot(ns - 1) = 1 - b
    For i = ns - 2 To 1 Step -1
        multiplier(i) = -c / pivot(i + 1)
        pivot(i) = 1 - b + multiplier(i) * a
    Next i
    For timeStep = 1 To nt + 2
        If timeStep <= 4 Then
            For i = 1 To ns - 1
                rhs(i) = v(i)
            Next i
        Else
            For i = 1 To ns - 1
                rhs(i) = a * v(i - 1) + (1 + b) * v(i) + c * v(i + 1)
            Next i
        End If
        rhs(ns - 1) = rhs(ns - 1) + c * v(ns)
        For i = ns - 2 To 1 Step -1
            rhs(i) = rhs(i) - multiplier(i) * rhs(i + 1)
        Next i
        For i = 1 To ns - 1
            z = (rhs(i) + a * v(i - 1)) / pivot(i)
            If z < payoff(i) Then z = payoff(i)
            v(i) = z
        Next i
    Next timeStep
    AmericanPutCN = v(nLeft)
End Function

' The template's European helpers, with bisection rather than Newton.
Public Function bsput(S As Double, sigma As Double, T As Double, _
        X As Double, r As Double) As Double
    Dim d1 As Double, d2 As Double
    If sigma = 0 Then
        bsput = PositivePart(X * Exp(-r * T) - S)
        Exit Function
    End If
    d1 = (Log(S / X) + (r + sigma * sigma / 2) * T) / (sigma * Sqr(T))
    d2 = d1 - sigma * Sqr(T)
    bsput = X * Exp(-r * T) * Application.NormSDist(-d2) - S * Application.NormSDist(-d1)
End Function

Public Function impPutVol(MktPrice As Double, S As Double, T As Double, _
        X As Double, r As Double) As Double
    Dim lo As Double, hi As Double, mid As Double
    If MktPrice >= X * Exp(-r * T) Then Err.Raise 5, "impPutVol", "No European implied volatility exists."
    lo = 0
    hi = 2
    Do While bsput(S, hi, T, X, r) < MktPrice
        hi = 2 * hi
    Loop
    Do While hi - lo > 0.0000000001
        mid = (lo + hi) / 2
        If bsput(S, mid, T, X, r) < MktPrice Then
            lo = mid
        Else
            hi = mid
        End If
    Loop
    impPutVol = hi
End Function

Private Function PositivePart(value As Double) As Double
    If value > 0 Then PositivePart = value Else PositivePart = 0
End Function

Private Function CeilingPositive(value As Double) As Long
    CeilingPositive = -Int(-value)
End Function
