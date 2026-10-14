Attribute VB_Name = "Project1ImpVol"
Option Explicit
Option Base 0

' American put, no dividends. Volatilities are decimal fractions.
' The optional arguments link the numerical controls to the template.
Public Function AmericanPutImpVol(MktPrice As Double, S As Double, T As Double, _
        X As Double, r As Double, Tol As Double, _
        Optional TimeSteps As Long = 100, Optional PriceSteps As Long = 200, _
        Optional MinPrice As Double = 0, Optional MaxPrice As Double = 100) As Double

    Dim nt As Long, ns As Long, iteration As Long
    Dim h As Double, upperPrice As Double, europeanVol As Double
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

    ' More than the template's minimum steps are used for basis-point accuracy.
    nt = TimeSteps
    If nt < 1600 Then nt = 1600
    ns = PriceSteps
    If ns < 3200 Then ns = 3200
    h = (MaxPrice - MinPrice) / ns

    ' At equal volatility the American put is at least as expensive as
    ' the European put, so the European implied volatility is an upper bound.
    europeanVol = impPutVol(MktPrice, S, T, X, r)
    upperPrice = MaxPrice
    ' Keep the original price spacing when extending the far boundary.
    If X * Exp(2.5 * europeanVol * Sqr(T)) > upperPrice Then
        upperPrice = X * Exp(2.5 * europeanVol * Sqr(T))
    End If
    ns = -Int(-(upperPrice - MinPrice) / h)
    upperPrice = MinPrice + ns * h

    x0 = 0
    x1 = europeanVol
    f0 = intrinsic - MktPrice
    f1 = AmericanPutCN(S, x1, T, X, r, nt, ns, MinPrice, upperPrice) - MktPrice
    ' The European bound applies to the exact prices. At zero or very
    ' small rates, finite-grid error can put the CN price just below the
    ' market price there. Ensure the numerical prices actually bracket
    ' the root before applying the bounded secant iteration.
    For iteration = 1 To 20
        If f1 >= 0 Then Exit For
        x1 = x1 * 1.01 + Tol
        f1 = AmericanPutCN(S, x1, T, X, r, nt, ns, MinPrice, upperPrice) - MktPrice
    Next iteration
    If f1 < 0 Then Err.Raise 5, "AmericanPutImpVol", "Unable to bracket the numerical root."
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
        f2 = AmericanPutCN(S, x2, T, X, r, nt, ns, MinPrice, upperPrice) - MktPrice
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

' Crank-Nicolson in the underlying price S, theta = 1/2.
' Reverse tridiagonal elimination and forward projected substitution
' enforce the American put's early-exercise obstacle at every time step.
Public Function AmericanPutCN(S As Double, sigma As Double, T As Double, _
        X As Double, r As Double, nt As Long, ns As Long, _
        MinPrice As Double, MaxPrice As Double) As Double

    Dim v() As Double, payoff() As Double, a() As Double
    Dim b() As Double, c() As Double, pivot() As Double
    Dim multiplier() As Double, rhs() As Double
    Dim dt As Double, ds As Double, spot As Double, q As Double
    Dim z As Double, location As Double, weight As Double
    Dim i As Long, n As Long, j As Long

    If sigma = 0 Then
        AmericanPutCN = PositivePart(X - S)
        Exit Function
    End If
    dt = T / nt
    ds = (MaxPrice - MinPrice) / ns
    ReDim v(0 To ns), payoff(0 To ns)
    ReDim a(1 To ns - 1), b(1 To ns - 1), c(1 To ns - 1)
    ReDim pivot(1 To ns - 1), multiplier(1 To ns - 1), rhs(1 To ns - 1)

    ' Terminal payoff and American boundary values.
    For i = 0 To ns
        spot = MinPrice + i * ds
        payoff(i) = PositivePart(X - spot)
        v(i) = payoff(i)
    Next i
    For i = 1 To ns - 1
        q = (MinPrice + i * ds) / ds
        a(i) = dt / 4 * (sigma * sigma * q * q - r * q)
        b(i) = -dt / 2 * (sigma * sigma * q * q + r)
        c(i) = dt / 4 * (sigma * sigma * q * q + r * q)
    Next i

    ' M has diagonals -a, 1-b, -c; N has diagonals a, 1+b, c.
    pivot(ns - 1) = 1 - b(ns - 1)
    For i = ns - 2 To 1 Step -1
        multiplier(i) = -c(i) / pivot(i + 1)
        pivot(i) = 1 - b(i) + multiplier(i) * a(i + 1)
    Next i
    For n = 1 To nt
        For i = 1 To ns - 1
            rhs(i) = a(i) * v(i - 1) + (1 + b(i)) * v(i) + c(i) * v(i + 1)
        Next i
        rhs(ns - 1) = rhs(ns - 1) + c(ns - 1) * v(ns)
        For i = ns - 2 To 1 Step -1
            rhs(i) = rhs(i) - multiplier(i) * rhs(i + 1)
        Next i
        For i = 1 To ns - 1
            z = (rhs(i) + a(i) * v(i - 1)) / pivot(i)
            If z < payoff(i) Then z = payoff(i)
            v(i) = z
        Next i
    Next n

    ' Linear interpolation if the current underlying price is off-grid.
    location = (S - MinPrice) / ds
    j = Int(location)
    weight = location - j
    AmericanPutCN = (1 - weight) * v(j) + weight * v(j + 1)
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
