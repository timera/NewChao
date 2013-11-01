Public Class Calculations
    Shared Lwr As Integer = 100
    'Background calculations
    Shared Function BackCalc(ByRef matrix As List(Of List(Of Double)), ByVal r As Double, ByRef warning As String)
        If IsNothing(matrix) Then
            Return False
        End If
        Dim rowCount = matrix.Count

        If rowCount > 6 Then
            For i = 6 To matrix.Count - 1
                Dim temp As List(Of Double) = matrix(i)
                matrix.Remove(temp)
            Next
        End If
        Dim LAverage As Double = Average(GetColumn(matrix, 0))
        ' but not adding to matrix
        Dim splRss As Double = (100 - 20 * Math.Log10(r) - 8) - LAverage
        If splRss >= 7 Then
            warning += "注意RSS背景噪音過高"
            Return False
        End If
        Return True
    End Function

    'RSS calculations

    Enum Modes
        A1
        A2
        A3
        A4
    End Enum
    ''3 time calculations
    'Shared Function ThreeTimeCalculation(ByRef matrix As List(Of List(Of Double)), ByVal mode As Modes, ByVal r As Double, ByRef warning As String)
    '    If mode = Modes.A1 Then
    '        Return A1Calc(matrix, r, warning)
    '    ElseIf mode = Modes.A2 Then
    '        Return A2Calc(matrix, r, warning)
    '    ElseIf mode = Modes.A3 Then
    '        Return A3Calc(matrix, r, warning)
    '    ElseIf mode = Modes.A4 Then
    '        Return A4Calc(matrix, r, warning)
    '    End If
    '    Return False
    'End Function

    'calculates the average
    Shared Function Average(ByVal numbers() As Double)
        If Not IsNothing(numbers) And Not numbers.Length = 0 Then
            Dim sum As Double = 0
            For i = 0 To numbers.Length - 1
                sum += 10 ^ (0.1 * numbers(i))
            Next
            Return 10 * Math.Log10(sum / numbers.Length)
        End If
        Return Nothing
    End Function

    'get the column from a matrix made of list of lists
    Shared Function GetColumn(ByRef matrix As List(Of List(Of Double)), ByVal cNum As Integer)
        Dim temp(matrix.Count) As Double
        For i = 0 To matrix.Count - 1
            temp(i) = matrix(i)(cNum)
        Next
        Return temp
    End Function

    'Get Environmental values
    Shared Function GetEnvVals(ByRef matrix As List(Of List(Of Double)), ByVal r As Double)
        If IsNothing(matrix) Then
            Return False
        End If
        Dim rowCount As Integer = matrix.Count
        Dim columnCount As Integer

        If matrix(0).Count > 0 Then
            columnCount = matrix(0).Count
        Else
            Return False
        End If

        Dim LstarW As Double = matrix(rowCount - 1)(0)
        Dim K2A As Double = LstarW - Lwr
        Dim LWA As List(Of Double) = New List(Of Double)
        LWA.Add(0)
        Dim max As Double = 0
        For i = 1 To columnCount - 1
            Dim LAverage As Double = matrix(rowCount - 4)(i)
            Dim K1A As Double = matrix(rowCount - 3)(i)
            Dim temp As Double = LAverage - K1A - K2A + 10 * Math.Log10(2 * Math.PI * r * r)
            LWA.Add(temp)
            If temp > max Then
                max = temp
            End If
        Next
        matrix.Add(LWA)
        Dim LWAchoice = New List(Of Double)
        LWAchoice.Add(max)
        matrix.Add(LWAchoice)
        Return max

    End Function


    'A1 calculation up to K1A
    Shared Function A1Calc(ByRef matrix As List(Of List(Of Double)), ByVal r As Double, ByRef warning As String)
        If IsNothing(matrix) Then
            Return False
        End If
        Dim rowCount = matrix.Count
        Dim columnCount As Integer

        If matrix(0).Count > 0 Then
            columnCount = matrix(0).Count
        Else
            Return False
        End If

        If rowCount > 6 Then
            For i = 6 To matrix.Count - 1
                Dim temp As List(Of Double) = matrix(i)
                matrix.Remove(temp)
            Next
        End If

        'All averages
        Dim LAverage As List(Of Double) = New List(Of Double)
        'background average and averages for diff trials

        For i = 0 To matrix(0).Count - 1
            LAverage.Add(Average(GetColumn(matrix, i)))
        Next
        matrix.Add(LAverage)

        'Delta LA
        Dim Ldelta As List(Of Double) = New List(Of Double)

        For i = 0 To columnCount
            Ldelta.Add(LAverage(i) - LAverage(0))
        Next
        matrix.Add(Ldelta)

        'K1A
        Dim K1A As List(Of Double) = New List(Of Double)

        For i = 0 To columnCount
            If Not i = 0 Then
                If Ldelta(i) >= 10 Then
                    K1A.Add(0)
                Else
                    K1A.Add(-10 * Math.Log10(1 - 10 ^ (-0.1 * Ldelta(i))))
                End If
                If Ldelta(i) <= 3 Then
                    warning += "注意背景噪音過高"
                End If
            Else
                K1A.Add(-10 * Math.Log10(1 - 10 ^ (-0.1 * Ldelta(i))))
            End If
        Next
        matrix.Add(K1A)

        'test how many passes are there- we need at least 2 to return true
        Dim pass As Integer = 0
        For i = 1 To columnCount
            If LAverage(i) + K1A(i) <= 1 Then
                pass += 1
            End If
            If pass = 2 Then
                Return True
            End If
        Next
        Return False
    End Function

    ' pretty much the same as A1 but adding manual times
    Shared Function A2Calc(ByRef matrix As List(Of List(Of Double)), ByVal r As Double, ByRef warning As String, ByVal times As List(Of Double))
        Dim success As Boolean = A1Calc(matrix, r, warning)
        If success Then
            matrix.Insert(matrix.Count - 3, times)
            Return True
        Else
            Return False
        End If
    End Function


    Shared Function A3Calc(ByRef matrix As List(Of List(Of Double)), ByVal r As Double, ByRef warning As String)

    End Function

    Shared Function A4Calc(ByRef matrix As List(Of List(Of Double)), ByVal r As Double, ByRef warning As String)

    End Function
End Class
