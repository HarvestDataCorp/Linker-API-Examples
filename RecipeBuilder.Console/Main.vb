Imports System.Security
Imports CommandLineParser.Arguments
Imports Harvest.Walmart
Imports System.Console
Imports System.IO
Imports Harvest.Shared

Module Main
    Private _filename As String
    Private _retailLink As RetailLink

    Sub Main(args() As String)
#Region "Argument Setup"
        'config letters used, do not repeat a letter - upkrfcs
        Dim parser As New CommandLineParser.CommandLineParser
        parser.CheckMandatoryArguments = True

        Dim username As New ValueArgument(Of String)(CChar("u"), "username", "Retail Link Username")
        username.Example = "-u User1234"
        username.Optional = False
        parser.Arguments.Add(username)

        Dim password As New ValueArgument(Of SecureString)(CChar("p"), "password", "Retail Link Password")
        password.Example = "-p MyPassword123"
        password.Optional = False
        password.ConvertValueHandler = Function(p As String)
            Dim secured As New SecureString

            If IsNothing(p) Or p = "" Then Return Nothing

            For Each c As Char In p
                secured.AppendChar(c)
            Next
            Return secured
                                       End Function
        parser.Arguments.Add(password)

        Dim pause As New SwitchArgument(CChar("k"), "key", "Require the user to hit any key at the end of the download process", False)
        pause.Optional = True
        parser.Arguments.Add(pause)

        Dim recipes As New ValueArgument(Of String)(CChar("r"), "recipes", "Reports you would like recipes saved for (case sensitive), must be used with files parameter")
        recipes.Example = "-r ""My Scorecard"""
        recipes.FullDescription =
            vbTab & "Please see http://msdn.microsoft.com/EN-US/library/swf8kaxw(v=VS.120,d=hv.2).aspx for more details of format" & vbCrLf & vbCrLf &
            vbTab & "Typical use may include:" & vbCrLf &
            vbTab & vbTab & "?               Any single character" & vbCrLf &
            vbTab & vbTab & "*               Zero or more characters" & vbCrLf &
            vbTab & vbTab & "#               Any single digit" & vbCrLf &
            vbTab & vbTab & "[charlist]      Any single character in charlist" & vbCrLf &
            vbTab & vbTab & "[!charlist]     Any single character not in charlist" & vbCrLf &
            vbTab & vbTab & "-               May be used in charlist to indicate a range [a-z]"
        recipes.Optional = True
        parser.Arguments.Add(recipes)

        Dim files As New ValueArgument(Of String)(CChar("f"), "files", "files you would like the report recipes saved to, must be used with recipes parameter")
        files.Example = "-f D:\Data\{Name}.txt"
        files.FullDescription =
            vbTab & "Possible parameters (case sensitive) are as follows, enclose all parameters in either {} or %%" & vbCrLf & vbCrLf &
            vbTab & "Name      Report name" & vbCrLf &
            vbTab & "User      The username you passed into the application" & vbCrLf &
            vbTab & "Today     Today's date generated from the local computer. Formattable, see http://tinyurl.com/format99" & vbCrLf &
            vbTab & "Now       Today's date and time generated from the local computer. Formattable, see http://tinyurl.com/format99" & vbCrLf &
            vbTab & "TW        Walmart fuzzy date that represents This Week based on the date the report was run in Retail Link" & vbCrLf &
            vbTab & "LW        Walmart fuzzy date that represents Last Week based on the date the report was run in Retail Link" & vbCrLf &
            vbTab & "LYTW      Walmart fuzzy date that represents This Week This Year based on the date the report was run in Retail Link." & vbCrLf &
            vbTab & "LYLW      Walmart fuzzy date that represents This Week Last Year based on the date the report was run in Retail Link."
        files.Optional = False
        parser.Arguments.Add(files)

        Dim config As New FileArgument(CChar("c"), "config", "Config file to submit report recipes for")
        config.Example = "-f D:\Data\myConfig.txt"
        config.FullDescription = 
            vbTab & "" & vbCrLf & vbCrLf &
            vbTab & "" & vbCrLf & vbCrLf &
            vbTab & "" & vbCrLf & vbCrLf &
            vbTab & "" & vbCrLf & vbCrLf &
            vbTab & "" & vbCrLf & vbCrLf &
            vbTab & "" & vbCrLf & vbCrLf &
            vbTab & "" & vbCrLf & vbCrLf &
            vbTab & "" & vbCrLf & vbCrLf &
            vbTab & "" & vbCrLf & vbCrLf &
            vbTab & "" & vbCrLf & vbCrLf &
            vbTab & "" & vbCrLf & vbCrLf &
            vbTab & "" & vbCrLf & vbCrLf
        config.Optional = True
        parser.Arguments.Add(config)

        Dim resubmit As New SwitchArgument(CChar("s"), "resubmit", "Do not resubmit erred reports", False)
        resubmit.Optional = True
        parser.Arguments.Add(resubmit)

#End Region

        
        Try
            parser.ParseCommandLine(args)
#If DEBUG
            parser.ShowParsedArguments()
#End If
            'Creating the RetailLink object and adding handlers for the console output
            _retailLink = New RetailLink(username.Value, password.Value)
            AddHandler _retailLink.DownloadStatus, AddressOf DownloadStatus
            AddHandler _retailLink.StatusUpdate, AddressOf StatusUpdate
            AddHandler _retailLink.StatusFound, AddressOf StatusFound

            'If we're supposed to save the recipes to files
            If recipes.Parsed AndAlso Not files.Parsed Then Throw New Exception("If you set the recipes parameter, you must set the files parameter")
            If recipes.Parsed AndAlso files.Parsed Then
                Dim savedReports = From savedReport In _retailLink.GetSavedReportsAsync().Result
                        Where savedReport.Name Like recipes.Value
                
                For Each savedReport in savedReports
                    Try
                        Dim reportRecipe = _retailLink.GetSavedReportRecipeAsync(savedReport).Result

                        'Here we use a Harvest built function called ReplaceToken, we can also used some date conversions from Harvest for Walmart Calendar Weeks
                        Dim filename = files.StringValue.
                                ReplaceToken("Name", reportRecipe.ReportName).
                                ReplaceToken("User", username.Value).
                                ReplaceToken("Today", Today.ToString("yyyyMMdd")).
                                ReplaceToken("Now", Now.ToString("yyyyyMMdd HHmmss")).
                                ReplaceToken("TW", Today.GetWmWeek()).
                                ReplaceToken("LW", Today.GetWmLastWeek()).
                                ReplaceToken("LYTW", Today.GetWmWeekLastYear()).
                                ReplaceToken("LYLW", Today.GetWmLastWeekLastYear())

                        'Write the report recipe to the file
                        Using file As New StreamWriter(filename)
                            With reportRecipe
                                file.WriteLine($"ReportName{vbTab}{.ReportName}")
                                file.WriteLine($"ApplicationId{vbTab}{.ApplicationId}")
                                file.WriteLine($"Compress{vbTab}{.Compress}")
                                file.WriteLine($"Country{vbTab}{.Country}")
                                file.WriteLine($"DivisionId{vbTab}{.DivisionId}")
                                file.WriteLine($"EncodedReport{vbTab}{.EncodedReport}")
                                file.WriteLine($"ExeId{vbTab}{.ExeId}")
                                file.WriteLine($"QuestionId{vbTab}{.QuestionId}")
                                file.WriteLine($"ReportFormat{vbTab}{.ReportFormat}")
                            End With
                        End Using
                    Catch ex As Exception

                    End Try
                Next
            End If

            'If we're supposed to submit recipes
            If config.Parsed
                Using configFile As New StreamReader(config.Value.FullName)
                    Dim reportRecipe as New ReportRecipe
                    Dim currentLine As String, variableName as String, variableValue as String
                    Dim upcs As New List(Of String)
                    Dim itemNums As New List(Of Long)
                    Dim storeStart as Long = -1, storeEnd as Long = -1
                    Dim day As Date = Nothing, dayStart as Date = Nothing, dayEnd As Date = Nothing

                    Do While configFile.Peek() >= 0
                        currentLine = configFile.ReadLine()
                        variableName = currentLine.Split(CChar(vbTab))(0)
                        variableValue = currentLine.Split(CChar(vbTab))(1)

                        Select Case variableName
                            Case "SubmitName"
                                reportRecipe.ReportName = variableValue
                            Case "UPC"
                                upcs.Add(variableValue)
                            Case "ItemNum"
                                itemNums.Add(CLng(variableValue))
                            Case "StoreStart"
                                storeStart = CLng(variableValue)
                            Case "StoreEnd"
                                storeEnd = CLng(variableValue)
                            Case "Day"
                                day = CDate(variableValue)
                            Case "dayStart"
                                dayStart = CDate(variableValue)
                            Case "dayEnd"
                                dayEnd = CDate(variableValue)
                            Case "Zip"
                                reportRecipe.Compress = CBool(variableValue)
                            Case "Format"
                                reportRecipe.ReportFormat = ReportType(variableValue)
                            Case "Filename"
                                Using recipeFile As New StreamReader(variableValue)
                                    Dim recipeLine As String, recipeName As String, recipeValue as String
                                    Do While recipeFile.Peek() >= 0
                                        recipeLine = recipeFile.ReadLine()
                                        recipeName = recipeLine.Split(CChar(vbTab))(0)
                                        recipeValue = recipeLine.Split(CChar(vbTab))(1)

                                        Select Case recipeName
                                            Case "ReportName"
                                                reportRecipe.ReportName = recipeValue
                                            Case "ApplicationId"
                                                reportRecipe.ApplicationId = CType(recipeValue, ApplicationType)
                                            Case "Compress"
                                                reportRecipe.Compress = CBool(recipeValue)
                                            Case "Country"
                                                reportRecipe.Country = CType(recipeValue, CountryCodes)
                                            Case "DivisionId"
                                                reportRecipe.DivisionId = CShort(recipeValue)
                                            Case "EncodedReport"
                                                reportRecipe.EncodedReport = recipeValue
                                            Case "ExeId"
                                                reportRecipe.ExeId = CShort(recipeValue)
                                            Case "QuestionId"
                                                reportRecipe.QuestionId = recipeValue
                                            Case "ReportFormat"
                                                reportRecipe.ReportFormat = ReportType(recipeValue)
                                        End Select

                                    Loop
                                End Using
                            Case Else
                                With reportRecipe
                                    If upcs.Count > 0 Then .ReplaceUpc(upcs)
                                    If itemNums.Count > 0 Then .ReplaceItems(itemNums)
                                    If storeStart > -1 AndAlso storeEnd > -1 Then .ReplaceStoreNumbers(storeStart,storeEnd)
                                    If Not IsNothing(day) Then .ReplaceTimeframe(day)
                                    If Not IsNothing(dayStart) AndAlso IsNothing(dayEnd) Then .ReplaceTimeframe(dayStart,dayEnd)
                                End With

                                WriteLine("Attempting to submit report: " & reportRecipe.ReportName)
                                WriteLine("Successfully submitted report, Job #: " & _retailLink.SubmitReportAsync(reportRecipe).ToString())
                        End Select
                    Loop

                End Using
            End If

#Region "Resubmit"
            If Not resubmit.Value Then
                Dim task2 = _retailLink.ResubmitErroredReportsAsync({ReportStatus.ErrorNormal, ReportStatus.ErrorFormatting}, True)

                Try
                    Task.WaitAll(task2)
                Catch ae As AggregateException
                    ForegroundColor = ConsoleColor.Red
                    WriteLine("Problem encountered while resubmitting reports: ")
                    WriteLine(From exception In ae.InnerExceptions Select exception.Message)
                    ForegroundColor = ConsoleColor.Gray
                End Try

                For Each Result In task2.Result
                    WriteLine("Unable to resubmit Failed Report: " & Result.SavedReport.Name)
                Next
            End If
#End Region

        Catch ex3 As ExpiredPasswordException
            ForegroundColor = ConsoleColor.Red
            WriteLine("Expired Password - Log into retail link, update your password and try again")
        Catch ex2 As BadPasswordException
            ForegroundColor = ConsoleColor.Red
            WriteLine("Bad Password")
        Catch ex As Exception
            ForegroundColor = ConsoleColor.Red
            WriteLine("Error Encountered: " & vbCrLf & vbCrLf & ex.Source & vbCrLf & ex.Message & vbCrLf & vbCrLf)
            ForegroundColor = ConsoleColor.Gray
            parser.ShowUsage
        Finally
            ForegroundColor = ConsoleColor.Gray

            If pause.Value Then
                WriteLine("Press any key to exit")
                ReadKey()
            End If
        End Try
    End Sub

    Function ReportType(input As String) As ReportFormat
        Select Case input
            Case "Excel"
                Return ReportFormat.Excel
            Case "Excel2007"
                Return ReportFormat.Excel2007
            Case "Access"
                Return ReportFormat.Access
            Case "Html"
                Return ReportFormat.Html
            Case "QuickView"
                Return ReportFormat.QuickView
            Case "Text"
                Return ReportFormat.Text
            Case Else
                Return Nothing
        End Select
    End Function

    Sub StatusUpdate(userId As String, status As String)
        WriteLine(Now.ToString & vbTab & status)
    End Sub

    Sub DownloadStatus(userId As String, filename As String, bytesDownloaded As Double)
        If _filename <> filename Then
            WriteLine()
            Write(filename)
            _filename = filename
        End If

        Write(vbCr & filename.PadRight(30) & Format(bytesDownloaded, "#,##0 bytes").PadLeft(20))
    End Sub

    Private Sub StatusFound(userId As String, status As Status)
        WriteLine(Now.ToString & vbTab & "Report Status: " & vbTab & status.CountryId & vbTab & status.JobId & vbTab & status.Name & vbTab & status.Status.ToString & vbTab & status.DateCompleted & vbTab & status.FileSize)
    End Sub
End Module