Imports System.IO
Imports System.Reflection


Module Application_Code

    Private error_reporting_level As String = "minimal"
    Private finds As Long = 0

    Sub Main(ByVal sArgs() As String)
        Try

        
        If Not sArgs.Length = 4 Then
                Console.WriteLine("Email Address Harvester (20061120.1)")
                Console.WriteLine("-------------------------")
                Console.WriteLine("Usage: executable [affected filename] [email address pattern] [savefile] [error level]")
                Console.WriteLine("  where:")
                Console.WriteLine("   - [affected filename]: the text file to be parsed")
                Console.WriteLine("   - [email address pattern]: the pattern to be located e.g. @commerce.uct.ac.za")
                Console.WriteLine("   - [savefile]: file to which the located tokens are to be appended to")
                Console.WriteLine("   - [error level]: the level of error reporting (full, minimal, none)")
            Else
                'Log_Handler(Execute_Code(sArgs(0), sArgs(1), sArgs(2)))
                Console.WriteLine(Execute_Code(sArgs(0), sArgs(1), sArgs(2)))
            End If
            '  Console.ReadLine()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Log_Handler(ByVal identifier_msg As String)
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((ApplicationPath() & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((ApplicationPath() & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_TFSR_Activity_Log.txt", True)

            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg)


            filewriter.Flush()
            filewriter.Close()

        Catch exc As Exception
            Error_Handler(exc, "Activity Logger")
        End Try
    End Sub

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((ApplicationPath() & "\").Replace("\\", "\") & "Error Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((ApplicationPath() & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_TFSR_Error_Log.txt", True)
            If error_reporting_level = "minimal" Then
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.Message.ToString)
            End If
            If error_reporting_level = "full" Then
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.ToString)
            End If

            filewriter.Flush()
            filewriter.Close()

        Catch exc As Exception
            Console.WriteLine("An error occurred in Email Address Harvester's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Function Execute_Code(ByVal filename As String, ByVal searchstring As String, ByVal savefile As String) As String
        Dim result As String
        result = "Fail. Reason Unknown"
        Dim continue As Boolean = True
        Try
            Dim finfo As FileInfo
            If continue = True Then
                finfo = New FileInfo(filename)
                If finfo.Exists = False Then
                    continue = False
                    result = "Fail. Cannot Access Specified Text File"
                End If
                finfo = Nothing
            End If
            If continue = True Then
                Try


                    Dim filereader As StreamReader = New StreamReader(filename)
                    Dim filewriter As StreamWriter = New StreamWriter(savefile, True)
                    Dim linetocheck, linetowrite As String
                    While filereader.Peek > -1
                        linetocheck = filereader.ReadLine
                        Try

                        
                        If linetocheck.ToLower.IndexOf(searchstring.ToLower) <> -1 Then

                                For Each token As String In linetocheck.ToLower.Split(" ")
                                    For Each token2 As String In token.ToLower.Split("<")
                                        For Each token3 As String In token2.ToLower.Split(">")
                                            For Each token4 As String In token3.ToLower.Split(":")
                                                If token4.ToLower.IndexOf(searchstring.ToLower) <> -1 Then

                                                    filewriter.WriteLine(token4.Replace("""", "").Replace("'", ""))
                                                End If
                                            Next
                                        Next
                                    Next
                                Next
                            End If
                        Catch ex As Exception
                            Error_Handler(ex, "Line Read")
                        End Try
                        'linetowrite = FastReplace(linetocheck, searchstring, replacestring)
                        'filewriter.WriteLine(linetowrite)
                    End While
                    filereader.Close()
                    filewriter.Close()
                Catch ex As Exception
                    Error_Handler(ex, "File Read Write")
                    result = "Fail. Read/Write Error Encountered. [Check Log for Details]"
                    continue = False
                End Try
            End If
            If continue = True Then
                Try
                    'If File.Exists(filename & "_XTEMPX") = True Then
                    '    File.Delete(filename)
                    '    File.Move(filename & "_XTEMPX", filename)
                        result = "Success."
                    ' End If
                Catch ex As Exception
                    Error_Handler(ex, "File Rename")
                    result = "Fail. File Rename Error Encountered. [Check Log for Details]"
                    continue = False
                End Try
            End If
        Catch ex As Exception
            Error_Handler(ex, "Execute Code")
            result = "Fail. Critical Error Encountered. [Check Log for Details]"
        End Try
        Return result
    End Function

    Private Function ApplicationPath() As String
        Return _
        Path.GetDirectoryName([Assembly].GetEntryAssembly().Location)
    End Function

    Private Function FastReplace(ByVal Expr As String, ByVal Find As String, ByVal Replacement As String) As String
        Dim builder As System.Text.StringBuilder
        Dim upCaseExpr, upCaseFind As String
        Dim lenOfFind, lenOfReplace As Integer
        Dim currentIndex, prevIndex As Integer

        builder = New System.Text.StringBuilder
        upCaseExpr = Expr.ToUpper()
        upCaseFind = Find.ToUpper()
        lenOfFind = Find.Length
        lenOfReplace = Replacement.Length
        currentIndex = upCaseExpr.IndexOf(upCaseFind, 0)
        lenOfReplace = 0
        Do While currentIndex >= 0
            finds = finds + 1
            builder.Append(Expr, prevIndex, currentIndex - prevIndex)
            builder.Append(Replacement)
            prevIndex = currentIndex + lenOfFind
            currentIndex = upCaseExpr.IndexOf(upCaseFind, prevIndex)
        Loop
        If prevIndex < Expr.Length Then
            builder.Append(Expr, prevIndex, Expr.Length - prevIndex)
        End If
        Return builder.ToString()
    End Function


End Module
