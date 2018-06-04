Imports System.Security
Imports CommandLineParser.Arguments
Imports Harvest.Walmart
Imports System.Console
Imports System.IO

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

        Dim headers As New SwitchArgument(CChar("h"), "headers", "Add headers to the file", True)
        pause.Optional = True
        parser.Arguments.Add(headers)

        Dim field As New ValueArgument(Of String)(CChar("f"), "field", "Field terminator")
        field.Example = "-r t"
        field.FullDescription =
            vbTab & "Specify the seperator between fields" & vbCrLf &
            vbTab & vbTab & "t               tab" & vbCrLf &
            vbTab & vbTab & "c               comma" & vbCrLf &
            vbTab & vbTab & "s               space" & vbCrLf &
            vbTab & vbTab & "|               pipe" & vbCrLf &
            vbTab & vbTab & "-               dash"
        field.Optional = False
        parser.Arguments.Add(field)

        Dim row As New ValueArgument(Of String)(CChar("r"), "row", "End of row separator")
        row.Example = "-r crlf"
        row.FullDescription =
            vbTab & "Specify the seperator between rows" & vbCrLf & vbCrLf &
            vbTab & vbTab & "cr              carriage return" & vbCrLf &
            vbTab & vbTab & "lf              line feed" & vbCrLf &
            vbTab & vbTab & "crlf            carriage return + line feed" & vbCrLf &
            vbTab & vbTab & "t               tab" & vbCrLf &
            vbTab & vbTab & "|               pipe" & vbCrLf &
            vbTab & vbTab & "-               dash"
        row.Optional = False
        parser.Arguments.Add(row)

        Dim items As New ValueArgument(Of String)(CChar("i"), "items", "File containing the list of items to search for")
        items.Example = "-i D:\Data\items.txt"
        items.FullDescription =
            vbTab & "File containing a list of items. One item number per line, each line separated by a carriage return and line feed"
        items.Optional = False
        parser.Arguments.Add(items)

        Dim output As New ValueArgument(Of String)(CChar("o"), "output", "File that containing the results")
        output.Example = "-o D:\Data\{LW}.txt"
        output.FullDescription =
            vbTab & "Possible parameters (case sensitive) are as follows, enclose all parameters in either {} or %%" & vbCrLf & vbCrLf &
            vbTab & "User      The username you passed into the application" & vbCrLf &
            vbTab & "Today     Today's date generated from the local computer. Formattable, see http://tinyurl.com/format99" & vbCrLf &
            vbTab & "Now       Today's date and time generated from the local computer. Formattable, see http://tinyurl.com/format99" & vbCrLf &
            vbTab & "TW        Walmart fuzzy date that represents This Week based on the date the report was run in Retail Link" & vbCrLf &
            vbTab & "LW        Walmart fuzzy date that represents Last Week based on the date the report was run in Retail Link" & vbCrLf &
            vbTab & "LYTW      Walmart fuzzy date that represents This Week This Year based on the date the report was run in Retail Link." & vbCrLf &
            vbTab & "LYLW      Walmart fuzzy date that represents This Week Last Year based on the date the report was run in Retail Link."
        output.Optional = False
        parser.Arguments.Add(output)

#End Region


        Try
            parser.ParseCommandLine(args)
#If DEBUG Then
            parser.ShowParsedArguments()
#End If
            Dim retailLink As New RetailLink(username.Value, password.Value)
            AddHandler retailLink.DownloadStatus, AddressOf DownloadStatus
            AddHandler retailLink.StatusUpdate, AddressOf StatusUpdate
            AddHandler retailLink.StatusFound, AddressOf StatusFound

            Dim file As New StreamReader(items.Value)
            Dim outfile As New StreamWriter(output.Value)
            Dim task0 As Task(Of IEnumerable(Of ItemInquiry))

            Dim fieldTerminator As String = ""
            Dim rowTerminator As String = ""

            Select Case field.Value
                Case "t" : fieldTerminator = vbTab
                Case "c" : fieldTerminator = ","
                Case "s" : fieldTerminator = " "
                Case "|" : fieldTerminator = "|"
                Case "-" : fieldTerminator = "-"
            End Select

            Select Case row.Value
                Case "cr" : rowTerminator = vbCr
                Case "lf" : rowTerminator = vbLf
                Case "crlf" : rowTerminator = vbCrLf
                Case "t" : rowTerminator = vbTab
                Case "|" : rowTerminator = "|"
                Case "-" : rowTerminator = "-"
            End Select

            If headers.Value Then
                outfile.Write("WM Week" & fieldTerminator)
                outfile.Write("Item Nbr" & fieldTerminator)
                outfile.Write("Club" & fieldTerminator)
                outfile.Write("Status" & fieldTerminator)
                outfile.Write("Sub Division" & fieldTerminator)
                outfile.Write("Region" & fieldTerminator)
                outfile.Write("Market" & fieldTerminator)
                outfile.Write("Orderable Cost" & fieldTerminator)
                outfile.Write("Warehouse Pack Cost" & fieldTerminator)
                outfile.Write("Sell" & fieldTerminator)
                outfile.Write("Margin" & fieldTerminator)
                outfile.Write("OH" & fieldTerminator)
                outfile.Write("CONS OH" & fieldTerminator)
                outfile.Write("OO" & fieldTerminator)
                outfile.Write("Claims" & fieldTerminator)
                outfile.Write("CW" & fieldTerminator)
                outfile.Write("LW" & fieldTerminator)
                outfile.Write("2WA" & fieldTerminator)
                outfile.Write("3WA" & fieldTerminator)
                outfile.Write("L4W" & fieldTerminator)
                outfile.Write("Item Effective Date" & fieldTerminator)
                outfile.Write("Out of Stock Date" & fieldTerminator)
                outfile.Write("Markdown Status" & fieldTerminator)
                outfile.Write("Upcharge" & fieldTerminator)
                outfile.Write(rowTerminator)
            End If

            Do While file.Peek >= 0
                Dim ItemNumber As Long = CLng(file.ReadLine)
                task0 = retailLink.GetItemInquiryData(ItemNumber)

                Try
                    Task.WaitAll(task0)
                Catch ae As AggregateException
                    Throw ae.InnerException
                End Try

                If Not IsNothing(task0.Result) Then

                    For Each item In task0.Result
                        outfile.Write(item.WalmartWeek & fieldTerminator)
                        outfile.Write(item.ItemNbr & fieldTerminator)
                        outfile.Write(item.ClubNbr & fieldTerminator)
                        outfile.Write(item.Status & fieldTerminator)
                        outfile.Write(item.Subdivision & fieldTerminator)
                        outfile.Write(item.Region & fieldTerminator)
                        outfile.Write(item.Market & fieldTerminator)
                        outfile.Write(item.OrderableCost & fieldTerminator)
                        outfile.Write(item.WarehousePackCost & fieldTerminator)
                        outfile.Write(item.Sell & fieldTerminator)
                        outfile.Write(item.Margin & fieldTerminator)
                        outfile.Write(item.OnHands & fieldTerminator)
                        outfile.Write(item.ConsOnHands & fieldTerminator)
                        outfile.Write(item.OO & fieldTerminator)
                        outfile.Write(item.Claims & fieldTerminator)
                        outfile.Write(item.Week1 & fieldTerminator)
                        outfile.Write(item.Week2 & fieldTerminator)
                        outfile.Write(item.Week3 & fieldTerminator)
                        outfile.Write(item.Week4 & fieldTerminator)
                        outfile.Write(item.FourWeekTotal & fieldTerminator)
                        outfile.Write(item.EffectiveDate & fieldTerminator)
                        outfile.Write(item.OutOfStockDate & fieldTerminator)
                        outfile.Write(item.MarkdownStatus & fieldTerminator)
                        outfile.Write(item.Upcharge)
                        outfile.Write(rowTerminator)
                    Next

                End If
            Loop

            file.Close()
            outfile.Close()

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