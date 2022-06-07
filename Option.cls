Option Explicit

Sub OptionReplicable()
    Dim s As Double, div As Double, dividendDeath As Double, k As Double, sDiv As Double, T As Double, r As Double, sigma As Double, n As Integer, h As Double
    Dim u As Double, d As Double, p As Double
    's = 20
    'div = 2
    'dividendDeath = 1.5
    'k = 20
    'T = 0.25
    'n = 3
    'r = 0.03
    'sigma = 0.25
    s = InputBox("Your stock price? ")
    div = InputBox("Your dividend yield? ")
    dividendDeath = InputBox("For how long the dividend will be cut off? ")
    k = InputBox("The call's strike price? ")
    T = InputBox("The call's maturity? ")
    n = InputBox("The analysis period? ")
    r = InputBox("The risk free rate?")
    sigma = InputBox("The volatility? ")

    On Error GoTo Here:
    h = T / n
    sDiv = s - div * Exp(-r * dividendDeath / 12)
    u = Exp(sigma * Sqr(h))
    d = Exp(-sigma * Sqr(h))
    p = (Exp(r * h) - d) / (u - d)

    range("A1").Value = "Stock Price"
    range("A2").Value = "Strike Price"
    range("A3").Value = "Maturity"
    range("A4").Value = "Risk-free rate"
    range("A5").Value = "Volatility"
    range("A6").Value = "Number of Steps"
    range("A7").Value = "h"
    range("D1").Value = "Dividend"
    range("D2").Value = "Dividend Death"
    range("A10").Value = "Adjusted Stock Price"
    range("G1").Value = "up"
    range("G2").Value = "down"
    range("G3").Value = "probability"
    
    range("A1").Offset(0, 1).Value = s
    range("A2").Offset(0, 1).Value = k
    range("A3").Offset(0, 1).Value = T
    range("A4").Offset(0, 1).Value = r
    range("A5").Offset(0, 1).Value = sigma
    range("A6").Offset(0, 1).Value = n
    range("A7").Offset(0, 1).Value = h
    range("D1").Offset(0, 1).Value = div
    range("D2").Offset(0, 1).Value = dividendDeath / 12
    range("A10").Offset(0, 1).Value = sDiv
    range("G1").Offset(0, 1).Value = u
    range("G2").Offset(0, 1).Value = d
    range("G3").Offset(0, 1).Value = p



    range("A12").Value = "Dynamic of the ex-Dividend stock price S*"
    range("A13").Value = "Time"
    range("A14").Value = "Period"

    Dim cell As Integer
    cell = n
    range("A" & 12 + cell + 4).Value = "Dynamic of the Stock Price"
    range("A" & 12 + 2 * (cell + 3) + 1).Value = "American call option price"

    Dim stockPriceExDiv As Double, stockPriceValue As Double, i As Integer, j As Integer
    For i = 0 To n
        range("A" & 13).Offset(0, i + 1).Value = i * h
        range("A" & 14).Offset(0, i + 1).Value = i
        For j = 0 To i
            stockPriceExDiv = sDiv * u ^ j * d ^ i * 1 / d ^ j
            range("A" & 12 + 3 + cell).Offset(-j, 1 + i).Value = stockPriceExDiv
            If i <= dividendDeath Then
               stockPriceValue = stockPriceExDiv + div * Exp(-r * (dividendDeath - i) / 12)
            Else
                stockPriceValue = stockPriceExDiv
            End If
            range("A" & 12 + 2 * cell + 6).Offset(-j, 1 + i).Value = stockPriceValue
        Next j
    Next i

    Dim americanCallValue As Double, inductionCall As Double, differenceCall As Double, nextCallUp As Double, nextCallDown As Double
    For j = 0 To n
        range("A" & 12 + 3 * cell + 8).Offset(-n + j, n + 1).Value = max(0, range("A" & 12 + 2 * cell + 6).Offset(-n + j, n + 1).Value - k)
    Next j
    nextCallUp = max(0, range("A" & 12 + 2 * cell + 6).Offset(-n, n + 1).Value - k)
    nextCallDown = max(0, range("A" & 12 + 2 * cell + 6).Offset(-n + 1, n + 1).Value - k)
    For i = 1 To n
        For j = i To n
            stockPriceValue = range("A" & 12 + 2 * cell + 6).Offset(-n + j, n + 1 - i).Value
            inductionCall = Exp(-r * h) * (p * nextCallUp + (1 - p) * nextCallDown)
            differenceCall = max(stockPriceValue - k, 0)
            americanCallValue = max(inductionCall, differenceCall)
            range("A" & 12 + 3 * cell + 8).Offset(-n + j, n + 1 - i).Value = americanCallValue
            nextCallUp = nextCallDown
            nextCallDown = range("A" & 12 + 3 * cell + 8).Offset(-n + j + 1, n + 1 - i).Value
        Next j
    Next i
    americanCallValue = Exp(-r * h) * (p * range("A" & 12 + 3 * cell + 8).Offset(-1, 2).Value + range("A" & 12 + 3 * cell + 8).Offset(0, 2))
    range("A" & 12 + 3 * cell + 8).Offset(0, 1).Value = americanCallValue
    MsgBox ("Your american Call value is " & americanCallValue & ". Press Ok to view the analysis on the sheet.")
    Exit Sub
Here:
    MsgBox ("There was an error")
End Sub


Function max(x As Double, y As Double) As Double
    If x >= y Then
        max = x
    Else
        max = y
    End If
End Function


