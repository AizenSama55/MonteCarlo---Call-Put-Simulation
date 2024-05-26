Imports System
Imports System.Linq
Imports System.Threading.Tasks
Imports System.Text
Imports YahooFinanceApi
Imports System.IO

Module MonteCarloSimulation
    Sub Main()
        ' Register the encoding provider
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
        RunMonteCarloSimulation().Wait()
    End Sub

    Async Function RunMonteCarloSimulation() As Task
        Do
            Try
                ' Prompt for stock symbol
                Console.Write("Enter the stock symbol for historical data (or type 'exit' to quit): ")
                Dim stockSymbol As String = Console.ReadLine()
                If stockSymbol.ToLower() = "exit" Then Exit Do

                ' Prompt for call or put
                Console.Write("Is this a Call or Put option? (Enter 'Call' or 'Put'): ")
                Dim optionType As String = Console.ReadLine()
                If optionType.ToLower() <> "call" AndAlso optionType.ToLower() <> "put" Then
                    Console.WriteLine("Invalid option type.")
                    Continue Do
                End If

                ' Prompt for target price
                Console.Write("Enter the target price: ")
                Dim targetPrice As Double
                If Not Double.TryParse(Console.ReadLine(), targetPrice) Then
                    Console.WriteLine("Invalid target price.")
                    Continue Do
                End If

                ' Prompt for days until expiry
                Console.Write("Enter the number of days until expiry: ")
                Dim days As Integer
                If Not Integer.TryParse(Console.ReadLine(), days) Then
                    Console.WriteLine("Invalid number of days.")
                    Continue Do
                End If

                ' Load historical data and calculate indicators
                Dim historicalData As List(Of Double) = Await LoadHistoricalData(stockSymbol)
                If historicalData Is Nothing OrElse historicalData.Count = 0 Then
                    Console.WriteLine("No historical data loaded.")
                    Continue Do
                End If

                Dim dailyReturns As List(Of Double) = CalculateDailyReturns(historicalData)
                If dailyReturns Is Nothing OrElse dailyReturns.Count = 0 Then
                    Console.WriteLine("No daily returns calculated.")
                    Continue Do
                End If

                Dim meanReturn As Double = dailyReturns.Average()
                Dim volatility As Double = CalculateVolatility(dailyReturns)

                ' Parameters
                Dim currentPrice As Double = historicalData.Last()
                Dim numSimulations As Integer = 1000000 ' Adjusted for better performance

                ' Allocate arrays for the simulation
                Dim finalPrices(numSimulations - 1) As Double

                ' Set parallel options
                Dim parallelOptions As New ParallelOptions()
                parallelOptions.MaxDegreeOfParallelism = Environment.ProcessorCount

                ' Thread-safe random number generator
                Dim rand = New Random()

                ' Perform the simulations in parallel
                Parallel.For(0, numSimulations, parallelOptions, Sub(i)
                                                                     Try
                                                                         Dim price As Double = currentPrice
                                                                         Dim localRand As New Random(rand.Next())

                                                                         Dim tempHistoricalPrices = New List(Of Double)(historicalData)
                                                                         For t As Integer = 1 To days
                                                                             Dim rsi As Double = CalculateRSI(tempHistoricalPrices, 14)
                                                                             Dim macd As Double = CalculateMACD(tempHistoricalPrices, 12, 26, 9)
                                                                             Dim randomShock As Double = meanReturn + volatility * (localRand.NextDouble() * 2 - 1)
                                                                             If rsi > 70 Then
                                                                                 randomShock -= volatility * 0.5 ' Adjust for overbought conditions
                                                                             ElseIf rsi < 30 Then
                                                                                 randomShock += volatility * 0.5 ' Adjust for oversold conditions
                                                                             End If
                                                                             price *= (1 + randomShock)
                                                                             tempHistoricalPrices.Add(price)
                                                                         Next
                                                                         finalPrices(i) = price
                                                                     Catch ex As Exception
                                                                         Console.WriteLine($"Error in simulation {i}: {ex.Message}")
                                                                         Console.WriteLine(ex.StackTrace)
                                                                     End Try
                                                                 End Sub)

                ' Calculate the number of simulations above or below the target price based on option type
                Dim targetCount As Integer
                If optionType.ToLower() = "call" Then
                    targetCount = finalPrices.Count(Function(price) price > targetPrice)
                ElseIf optionType.ToLower() = "put" Then
                    targetCount = finalPrices.Count(Function(price) price < targetPrice)
                Else
                    Console.WriteLine("Invalid option type.")
                    Continue Do
                End If

                ' Output results
                Console.WriteLine($"Number of simulations where the price meets the target: {targetCount}")
                Console.WriteLine($"Probability of meeting the target: {(targetCount / numSimulations * 100).ToString("F2")}%")

                ' Save final prices to a CSV file
                SaveFinalPricesToFile(finalPrices, "final_prices.csv")

            Catch ex As Exception
                Console.WriteLine($"An error occurred: {ex.Message}")
                Console.WriteLine(ex.StackTrace)
            End Try

        Loop While True

        ' Wait for user input before closing
        Console.WriteLine("Press Enter to exit...")
        Console.ReadLine()
    End Function

    Async Function LoadHistoricalData(stockSymbol As String) As Task(Of List(Of Double))
        Dim data As New List(Of Double)
        Try
            ' Fetch historical data using YahooFinanceApi
            Dim historicalData = Await Yahoo.GetHistoricalAsync(stockSymbol, DateTime.Now.AddYears(-1), DateTime.Now, Period.Daily)

            For Each quote In historicalData
                data.Add(quote.Close)
            Next
        Catch ex As Exception
            Console.WriteLine($"Error loading historical data: {ex.Message}")
        End Try
        Return data
    End Function

    Function CalculateDailyReturns(prices As List(Of Double)) As List(Of Double)
        Dim returns As New List(Of Double)
        For i As Integer = 1 To prices.Count - 1
            returns.Add((prices(i) - prices(i - 1)) / prices(i - 1))
        Next
        Return returns
    End Function

    Function CalculateVolatility(returns As List(Of Double)) As Double
        Dim meanReturn As Double = returns.Average()
        Dim squaredDifferences As New List(Of Double)
        For Each ret As Double In returns
            squaredDifferences.Add((ret - meanReturn) ^ 2)
        Next
        Return Math.Sqrt(squaredDifferences.Average())
    End Function

    Function CalculateSMA(prices As List(Of Double), period As Integer) As Double
        If prices.Count < period Then Return 0
        Dim sum As Double = 0
        For i As Integer = prices.Count - period To prices.Count - 1
            sum += prices(i)
        Next
        Return sum / period
    End Function

    Function CalculateEMA(prices As List(Of Double), period As Integer) As Double
        If prices.Count < period Then Return 0
        Dim multiplier As Double = 2 / (period + 1)
        Dim ema As Double = prices(prices.Count - period)
        For i As Integer = prices.Count - period + 1 To prices.Count - 1
            ema = (prices(i) - ema) * multiplier + ema
        Next
        Return ema
    End Function

    Function CalculateRSI(prices As List(Of Double), period As Integer) As Double
        If prices.Count < period Then Return 50 ' Return neutral value if not enough data
        Dim gains As Double = 0
        Dim losses As Double = 0
        For i As Integer = prices.Count - period To prices.Count - 1
            If i <= 0 Then Continue For ' Ensure index is non-negative
            Dim change As Double = prices(i) - prices(i - 1)
            If change > 0 Then
                gains += change
            Else
                losses -= change
            End If
        Next
        Dim avgGain As Double = gains / period
        Dim avgLoss As Double = losses / period
        If avgLoss = 0 Then Return 100
        Dim rs As Double = avgGain / avgLoss
        Return 100 - (100 / (1 + rs))
    End Function

    Function CalculateMACD(prices As List(Of Double), shortPeriod As Integer, longPeriod As Integer, signalPeriod As Integer) As Double
        If prices.Count < Math.Max(shortPeriod, longPeriod) Then Return 0 ' Return 0 if not enough data
        Dim shortEMA As Double = CalculateEMA(prices, shortPeriod)
        Dim longEMA As Double = CalculateEMA(prices, longPeriod)
        Dim macd As Double = shortEMA - longEMA
        If prices.Count < signalPeriod Then Return macd ' Return macd if not enough data for signal line
        Dim signalLine As Double = CalculateEMA(prices.Take(prices.Count - signalPeriod).ToList(), signalPeriod)
        Return macd - signalLine
    End Function

    Sub SaveFinalPricesToFile(finalPrices As Double(), filePath As String)
        Try
            Using writer As New StreamWriter(filePath)
                For Each price In finalPrices
                    writer.WriteLine(price)
                Next
            End Using
        Catch ex As Exception
            Console.WriteLine($"Error saving final prices to file: {ex.Message}")
        End Try
    End Sub
End Module
