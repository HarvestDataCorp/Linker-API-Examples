Imports System.Security
Imports CommandLineParser.Arguments
Imports Harvest.Walmart
Imports System.Console

Module Main
    Private _filename As String

    Sub Main(args() As String)
#Region "Argument Setup"
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

        Dim reports As New ValueArgument(Of String)(CChar("r"), "reports", "Reports you would like downloaded (case sensitive)")
        reports.Example = "-r *"
        reports.FullDescription =
            vbTab & "Please see http://msdn.microsoft.com/EN-US/library/swf8kaxw(v=VS.120,d=hv.2).aspx for more details of format" & vbCrLf & vbCrLf &
            vbTab & "Typical use may include:" & vbCrLf &
            vbTab & vbTab & "?               Any single character" & vbCrLf &
            vbTab & vbTab & "*               Zero or more characters" & vbCrLf &
            vbTab & vbTab & "#               Any single digit" & vbCrLf &
            vbTab & vbTab & "[charlist]      Any single character in charlist" & vbCrLf &
            vbTab & vbTab & "[!charlist]     Any single character not in charlist" & vbCrLf &
            vbTab & vbTab & "-               May be used in charlist to indicate a range [a-z]"
        reports.Optional = False
        parser.Arguments.Add(reports)

        Dim folder As New ValueArgument(Of String)(CChar("f"), "folder", "folder you would like the report saved to")
        folder.Example = "-f D:\Data\{Name}.{FileType}"
        folder.FullDescription =
            vbTab & "Possible parameters (case sensitive) are as follows, enclose all parameters in either {} or %%" & vbCrLf & vbCrLf &
            vbTab & "Name      Report name" & vbCrLf &
            vbTab & "User      The username you passed into the application" & vbCrLf &
            vbTab & "Filename  Random filename assigned by Retail Link, typically username followed by underscore, random 9 characters, underscore, & 36 characters" & vbCrLf &
            vbTab & "FileType  The report's filetype, xls, xlxs, mdb, htm, txt" & vbCrLf &
            vbTab & "Today     Today's date generated from the local computer. Formattable, see http://tinyurl.com/format99" & vbCrLf &
            vbTab & "Now       Today's date and time generated from the local computer. Formattable, see http://tinyurl.com/format99" & vbCrLf &
            vbTab & "Date      Date completed generated from retail link time. Formattable, see http://tinyurl.com/format99" & vbCrLf &
            vbTab & "DateTime  Date & time completed generated from retail link time. Formattable, see http://tinyurl.com/format99" & vbCrLf &
            vbTab & "TW        Walmart fuzzy date that represents This Week based on the date the report was run in Retail Link" & vbCrLf &
            vbTab & "LW        Walmart fuzzy date that represents Last Week based on the date the report was run in Retail Link" & vbCrLf &
            vbTab & "LYTW      Walmart fuzzy date that represents This Week This Year based on the date the report was run in Retail Link." & vbCrLf &
            vbTab & "LYLW      Walmart fuzzy date that represents This Week Last Year based on the date the report was run in Retail Link."
        folder.Optional = False
        parser.Arguments.Add(folder)

        Dim unzip As New SwitchArgument(CChar("z"), "unzip", "Do not have automatically unzip downloaded files", True)
        unzip.Optional = True
        parser.Arguments.Add(unzip)

        Dim resubmit As New SwitchArgument(CChar("s"), "resubmit", "Do not resubmit erred reports", False)
        resubmit.Optional = True
        parser.Arguments.Add(resubmit)

        Dim mark As New SwitchArgument(CChar("m"), "mark", "Do not mark reports as retrieved after they successfully download", True)
        mark.Optional = True
        parser.Arguments.Add(mark)

#End Region

        
        Try
            parser.ParseCommandLine(args)
#If DEBUG
            parser.ShowParsedArguments()
#End If
            Dim retailLink As New RetailLink(username.Value, password.Value)
            AddHandler retailLink.DownloadStatus, AddressOf DownloadStatus
            AddHandler retailLink.StatusUpdate, AddressOf StatusUpdate
            AddHandler retailLink.StatusFound, AddressOf StatusFound

            Dim task0 = retailLink.GetReportStatusAsync
            Try
                Task.WaitAll(task0)
            Catch ae As AggregateException
                Throw ae.InnerException
            End Try

            Dim status = From report In task0.Result
                    Where report.Name Like reports.Value And report.Status = ReportStatus.Done

            For Each item In status
                Try
                    Dim task1 = retailLink.DownloadReportAsync(item, folder.Value, unzip.Value)
                    Task.WaitAll(task1)
                    WriteLine()

                    If task1.Result AndAlso mark.Value Then
                        task1 = retailLink.MarkAsRetrievedAsync(item)
                        Task.WaitAll(task1)
                    End If

                    WriteLine()
                Catch ae As AggregateException
                    WriteLine("Error Occurred with report: " & item.Name & " " & item.Filename & vbCrLf & ae.InnerException.Message)
                Catch ex As Exception
                    WriteLine("Error Occurred with report: " & item.Name & " " & item.Filename & vbCrLf & ex.Message)
                End Try
            Next

            If Not resubmit.Value Then
                Dim task2 = retailLink.ResubmitErroredReportsAsync({ReportStatus.ErrorNormal, ReportStatus.ErrorFormatting}, reports.Value, True)

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

            WriteLine()
            WriteLine()
            WriteLine("Processing completed")


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