
Sub GenrePieChart()

    'Declarations
    Dim Rng As Range
    Dim xColIndex As Integer
    Dim xRowIndex As Integer
    Dim TotalGenreCells As Range
    Dim GenreChart As Object
    Dim ChartName As String
    
    'Variables are initialized as 0 by default
    Dim Comedy As Integer
    Dim Thriller As Integer
    Dim Horror As Integer
    Dim Drama As Integer
    Dim Documentary As Integer
    Dim War As Integer
    Dim Action As Integer
    Dim SciFi As Integer
    Dim Western As Integer
    Dim Biography As Integer
    Dim Animated As Integer
    Dim Fantasy As Integer
    Dim Romance As Integer
    Dim Short As Integer
    Dim Other As Integer
    Dim DoubleGenre() As String
    
    'Data Selection
    Worksheets("Movie Log").Activate
    xIndex = 3
    xRowIndex = Application.ActiveSheet.Cells(Rows.Count, xIndex).End(xlUp).Row
    Range(Cells(3, xIndex), Cells(xRowIndex, xIndex)).Select
    Set TotalGenreCells = Range(Cells(3, xIndex), Cells(xRowIndex, xIndex))
    
    'Genre Sorting
    'Loops through selected cells and places and counts cells for each genre
    Dim i As Integer
    For i = 1 To TotalGenreCells.Rows.Count
        
        DoubleGenre() = Split(TotalGenreCells.Cells(i).Value, "/")
        
        If DoubleGenre(0) = "Comedy" Then
            Comedy = Comedy + 1
        ElseIf DoubleGenre(0) = "Thriller" Then
            Thriller = Thriller + 1
        ElseIf DoubleGenre(0) = "Horror" Then
            Horror = Horror + 1
        ElseIf DoubleGenre(0) = "Drama" Then
            Drama = Drama + 1
        ElseIf DoubleGenre(0) = "Documentary" Then
            Documentary = Documentary + 1
        ElseIf DoubleGenre(0) = "War" Then
            War = War + 1
        ElseIf DoubleGenre(0) = "Action" Then
            Action = Action + 1
        ElseIf DoubleGenre(0) = "SciFi" Or DoubleGenre(0) = "Sci-Fi" Then
            SciFi = SciFi + 1
        ElseIf DoubleGenre(0) = "Western" Then
            Western = Western + 1
        ElseIf DoubleGenre(0) = "Biography" Then
            Biography = Biography + 1
        ElseIf DoubleGenre(0) = "Animated" Or DoubleGenre(0) = "Animation" Then
            Animated = Animated + 1
        ElseIf DoubleGenre(0) = "Fantasy" Then
            Fantasy = Fantasy + 1
        ElseIf DoubleGenre(0) = "Romance" Then
            Romance = Romance + 1
        ElseIf DoubleGenre(0) = "Short" Then
            Short = Short + 1
        Else
            Other = Other + 1
        End If
    Next i
    
    'Worksheet Selection and Data Plotting
    Worksheets("Graphs and Stats").Activate
    Cells(2, 2) = Comedy
    Cells(3, 2) = Thriller
    Cells(4, 2) = Horror
    Cells(5, 2) = Drama
    Cells(6, 2) = Documentary
    Cells(7, 2) = War
    Cells(8, 2) = Action
    Cells(9, 2) = SciFi
    Cells(10, 2) = Western
    Cells(11, 2) = Biography
    Cells(12, 2) = Animated
    Cells(13, 2) = Fantasy
    Cells(14, 2) = Romance
    Cells(15, 2) = Short
    Cells(16, 2) = Other
    
    For i = 1 To ActiveSheet.ChartObjects.Count
        If ActiveSheet.ChartObjects(i).Chart.Name = "Graphs and Stats Genre" Then
            ActiveSheet.ChartObjects("Genre").Delete
        End If
    Next i
    
    'Graph Creation
    Set Rng = ActiveSheet.Range("A2:B16")
    Set GenreChart = ActiveSheet.Shapes.AddChart
    GenreChart.Name = "Genre"
    With GenreChart.Chart
        .SetSourceData Source:=Rng
        .ChartType = xlPie
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = "Genre Breakdown"
    End With
    Dim ChartSize As ChartObject
    Set ChartSizing = Worksheets("Graphs and Stats").ChartObjects("Genre")
    With ChartSizing
        .Top = Range("D2").Top
        .Width = Range("D2:H23").Width
        .Height = Range("D2:H23").Height
        .Left = Range("D2").Left
    End With

End Sub

Sub AverageRating()

    'Declarations
    Dim xColIndex As Integer
    Dim xRowIndex As Integer
    Dim TotalRatingCells As Range
    Dim RatingSum As Double
    
    'Data Selection
    Worksheets("Movie Log").Activate
    xIndex = 4
    xRowIndex = Application.ActiveSheet.Cells(Rows.Count, xIndex).End(xlUp).Row
    Range(Cells(3, xIndex), Cells(xRowIndex, xIndex)).Select
    Set TotalRatingCells = Range(Cells(3, xIndex), Cells(xRowIndex, xIndex))
    
    'Ratings Sum
    Dim i As Integer
    For i = 1 To TotalRatingCells.Rows.Count
        RatingSum = TotalRatingCells.Cells(i).Value + RatingSum
    Next i
    
    'Averaging
    RatingSum = Round((RatingSum / TotalRatingCells.Rows.Count), 2)
    
    'Worksheet Selection and Data Plotting
    Worksheets("Graphs and Stats").Activate
    Cells(25, 2) = RatingSum
    
End Sub

Sub HomeOrTheater()

    'Declarations
    Dim xColIndex As Integer
    Dim xRowIndex As Integer
    Dim TotalLocationCells As Range
    Dim Home As Integer
    Dim Theater As Integer
    
    'Data Selection
    Worksheets("Movie Log").Activate
    xIndex = 7
    xRowIndex = Application.ActiveSheet.Cells(Rows.Count, xIndex).End(xlUp).Row
    Range(Cells(3, xIndex), Cells(xRowIndex, xIndex)).Select
    Set TotalLocationCells = Range(Cells(3, xIndex), Cells(xRowIndex, xIndex))
    
    'Home and Theater Sums
    Dim i As Integer
    For i = 1 To TotalLocationCells.Rows.Count
        
        If TotalLocationCells.Cells(i).Value = "H" Then
            Home = Home + 1
        ElseIf TotalLocationCells.Cells(i).Value = "T" Then
            Theater = Theater + 1
        End If
    Next i
    
    'Worksheet Selection and Data Plotting
    Worksheets("Graphs and Stats").Activate
    Cells(28, 2) = Home
    Cells(29, 2) = Theater

End Sub

Sub HighLowRatings()

    'Declarations
    Dim xColIndex As Integer
    Dim xRowIndex As Integer
    Dim TotalRatingCells As Range
    Dim HighRating As Double
    Dim LowRating As Double
    LowRating = 10
    
    'Data Selection
    Worksheets("Movie Log").Activate
    xIndex = 4
    xRowIndex = Application.ActiveSheet.Cells(Rows.Count, xIndex).End(xlUp).Row
    Range(Cells(3, xIndex), Cells(xRowIndex, xIndex)).Select
    Set TotalRatingCells = Range(Cells(3, xIndex), Cells(xRowIndex, xIndex))
    
    'Finding Min and Max
    Dim i As Integer
    For i = 1 To TotalRatingCells.Rows.Count
    
        If TotalRatingCells.Cells(i).Value > HighRating Then
            HighRating = TotalRatingCells.Cells(i).Value
        ElseIf TotalRatingCells.Cells(i).Value < LowRating Then
            LowRating = TotalRatingCells.Cells(i).Value
        End If
    Next i
    
    'Worksheet Selection and Data Plotting
    Worksheets("Graphs and Stats").Activate
    Cells(26, 2) = HighRating
    Cells(27, 2) = LowRating
    
End Sub

Sub MoviesPerYear()
    'CURRENTLY ONLY SUMMING ENTRIES WATCHED SINCE THE TIME OF THIS SPREADSHEET'S CREATION
    'Could (and probably should) turn this into a dynamic array
    
    'Declarations
    Dim xColIndex As Integer
    Dim xRowIndex As Integer
    Dim TotalDateCells As Range
    
    'Variables are initialized as 0 by default
    Dim YR2016 As Integer
    Dim YR2017 As Integer
    Dim YR2018 As Integer
    Dim DateSplit() As String
    
    'Data Selection
    Worksheets("Movie Log").Activate
    xIndex = 2
    xRowIndex = Application.ActiveSheet.Cells(Rows.Count, xIndex).End(xlUp).Row
    Range(Cells(3, xIndex), Cells(xRowIndex, xIndex)).Select
    Set TotalDateCells = Range(Cells(3, xIndex), Cells(xRowIndex, xIndex))
    
    'Year Summing
    Dim i As Integer
    For i = 1 To TotalDateCells.Rows.Count
        
        DateSplit() = Split(TotalDateCells.Cells(i).Value, "/")
        If UBound(DateSplit()) = 2 Then
            If DateSplit(2) = "2016" Then
                YR2016 = YR2016 + 1
            ElseIf DateSplit(2) = "2017" Then
                YR2017 = YR2017 + 1
            ElseIf DateSplit(2) = "2018" Then
                YR2018 = YR2018 + 1
            End If
        End If
    Next i
    
    'Worksheet Selection and Data Plotting
    Worksheets("Graphs and Stats").Activate
    Cells(26, 5) = YR2016
    Cells(27, 5) = YR2017
    Cells(28, 5) = YR2018
End Sub

Sub MoviesPerMonth()
    '*DANGER* UNDER CONSTRUCTION *DANGER*

    'Declarations
    Dim xColIndex As Integer
    Dim xRowIndex As Integer
    Dim TotalDateCells As Range
    Dim MonthChart As Object
    Dim Rng As Range
    
    'Variables are initialized as 0 by default
    Dim Months(1 To 12) As Integer
    Dim MaxMonthMovies As Integer   'Holds the max value
    Dim MinMonthMovies As Integer
    Dim MaxMonth As Integer         'Holds month where max value occurred
    Dim MinMonth As Integer
    Dim Trend As Integer            'For now, the viewing change compared to the previous month
    Dim CurrentDateSplit() As String
    Dim DateSplit() As String
    
    'Data Selection
    Worksheets("Movie Log").Activate
    xIndex = 2
    xRowIndex = Application.ActiveSheet.Cells(Rows.Count, xIndex).End(xlUp).Row
    Range(Cells(3, xIndex), Cells(xRowIndex, xIndex)).Select
    Set TotalDateCells = Range(Cells(3, xIndex), Cells(xRowIndex, xIndex))
    
    'Month Summing
    Dim i As Integer
    For i = 1 To TotalDateCells.Rows.Count
        
        CurrentDateSplit() = Split(Date, "/")
        DateSplit() = Split(TotalDateCells.Cells(i).Value, "/")
        
        'Checks that the date is complete and that it is the current year
        If UBound(DateSplit()) = 2 Then
            If DateSplit(2) = CurrentDateSplit(2) Then
                If DateSplit(0) = "1" Then
                    Months(1) = Months(1) + 1
                ElseIf DateSplit(0) = "2" Then
                    Months(2) = Months(2) + 1
                ElseIf DateSplit(0) = "3" Then
                    Months(3) = Months(3) + 1
                ElseIf DateSplit(0) = "4" Then
                    Months(4) = Months(4) + 1
                ElseIf DateSplit(0) = "5" Then
                    Months(5) = Months(5) + 1
                ElseIf DateSplit(0) = "6" Then
                    Months(6) = Months(6) + 1
                ElseIf DateSplit(0) = "7" Then
                    Months(7) = Months(7) + 1
                ElseIf DateSplit(0) = "8" Then
                    Months(8) = Months(8) + 1
                ElseIf DateSplit(0) = "9" Then
                    Months(9) = Months(9) + 1
                ElseIf DateSplit(0) = "10" Then
                    Months(10) = Months(10) + 1
                ElseIf DateSplit(0) = "11" Then
                    Months(11) = Months(11) + 1
                ElseIf DateSplit(0) = "12" Then
                    Months(12) = Months(12) + 1
                End If
            End If
        End If
    Next i
    
    'Min and Max Months for Quantity of Movies
    MaxMonthMovies = WorksheetFunction.Max(Months)
    MaxMonth = WorksheetFunction.Match(MaxMonthMovies, Months, 0)
    MinMonthMovies = WorksheetFunction.Min(Months)
    MinMonth = WorksheetFunction.Match(MinMonthMovies, Months, 0)
        
    'Worksheet Selection and Data Plotting
    Worksheets("Graphs and Stats").Activate
    Cells(26, 8) = MonthName(MaxMonth) & " (" & MaxMonthMovies & ")"
    Cells(27, 8) = MonthName(MinMonth) & " (" & MinMonthMovies & ")"
    
    'Trends
    Trend = (Months(CInt(CurrentDateSplit(0))) - Months(CInt(CurrentDateSplit(0)) - 1))
    If (Months(CInt(CurrentDateSplit(0)) - 1)) = 0 Then
        Trend = Trend * 100
    Else
        Trend = Trend / (Months(CInt(CurrentDateSplit(0)) - 1)) * 100
    End If
           
    If Months(CurrentDateSplit(0)) > Months(CurrentDateSplit(0) - 1) Then
        Cells(28, 8) = "+" & Trend & "% this month"
    ElseIf Months(CurrentDateSplit(0)) < Months(CurrentDateSplit(0) - 1) Then
        Cells(28, 8) = "-" & Trend & "% this month"
    ElseIf Months(CurrentDateSplit(0)) = Months(CurrentDateSplit(0) - 1) Then
        Cells(28, 8) = "No change"
    End If
    
    'Worksheet Selection and Data Plotting
    Worksheets("Graphs and Stats").Activate
    Cells(2, 11) = Months(1)
    Cells(3, 11) = Months(2)
    Cells(4, 11) = Months(3)
    Cells(5, 11) = Months(4)
    Cells(6, 11) = Months(5)
    Cells(7, 11) = Months(6)
    Cells(8, 11) = Months(7)
    Cells(9, 11) = Months(8)
    Cells(10, 11) = Months(9)
    Cells(11, 11) = Months(10)
    Cells(12, 11) = Months(11)
    Cells(13, 11) = Months(12)
    
    'Test comment pls ignore
    
    'Deleting Old Monthly Chart
    For i = 1 To ActiveSheet.ChartObjects.Count
        If ActiveSheet.ChartObjects(i).Chart.Name = "Graphs and Stats Month" Then
            ActiveSheet.ChartObjects("Month").Delete
        End If
    Next i
     
    'Graph Creation
    Set Rng = ActiveSheet.Range("J2:K13")
    Set MonthChart = ActiveSheet.Shapes.AddChart
    MonthChart.Name = "Month"
    With MonthChart.Chart
        .SetSourceData Source:=Rng
        .ChartType = xlLineMarkers
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = "Viewing By Month"
        .HasLegend = False
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Movies Watched"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Month"
    End With
    
    Dim ChartSize As ChartObject
    Set ChartSizing = Worksheets("Graphs and Stats").ChartObjects("Month")
    With ChartSizing
        .Top = Range("M2").Top
        .Width = Range("M2:U16").Width
        .Height = Range("M2:U16").Height
        .Left = Range("M2").Left
    End With
    
End Sub

