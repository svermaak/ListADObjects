Imports System.DirectoryServices
Imports System.Environment
Imports System.IO
Module modMain
    Private Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
    Dim strDomainFQDN As String
    Dim strdefaultNamingContext As String
    Dim strPath As String
    Sub Main()
        Dim arrCommandLineArgs As String()

        Console.WriteLine("")
        Console.WriteLine("ListADObjects v1.00 by Shaun Vermaak {22E1AA44-6A30-4e11-9F6E-45CD0A627B5C}")
        Console.WriteLine("")

        Try
            arrCommandLineArgs = GetCommandLineArgs()
            If arrCommandLineArgs.GetUpperBound(0) = 2 Then
                strDomainFQDN = arrCommandLineArgs(1)
                strdefaultNamingContext = arrCommandLineArgs(2)

                Dim cki As System.ConsoleKeyInfo
                Dim strCommand As String
                Dim strTemp As String

                Do
                    cki = Console.ReadKey(True)
                    If cki.Key = ConsoleKey.Backspace Then
                        If strCommand.Length > 0 Then
                            strCommand = strCommand.Remove(strCommand.Length - 1)
                            Console.Write(cki.KeyChar)
                            Console.Write(" ")
                            Console.Write(cki.KeyChar)
                        End If
                    ElseIf cki.Key >= ConsoleKey.A And cki.Key <= ConsoleKey.Z Then
                        strCommand &= cki.KeyChar
                        Console.Write(cki.KeyChar)
                    ElseIf cki.Key >= ConsoleKey.D0 And cki.Key <= ConsoleKey.D9 Then
                        strCommand &= cki.KeyChar
                        Console.Write(cki.KeyChar)
                    ElseIf cki.Key = ConsoleKey.Spacebar Then
                        strCommand &= cki.KeyChar
                        Console.Write(cki.KeyChar)
                    ElseIf cki.Key = ConsoleKey.OemPeriod Then
                        strCommand &= cki.KeyChar
                        Console.Write(cki.KeyChar)
                    ElseIf cki.KeyChar = "\" Then
                        strCommand &= cki.KeyChar
                        Console.Write(cki.KeyChar)
                    ElseIf cki.KeyChar = ":" Then
                        strCommand &= cki.KeyChar
                        Console.Write(cki.KeyChar)
                    ElseIf cki.KeyChar = "-" Then
                        strCommand &= cki.KeyChar
                        Console.Write(cki.KeyChar)
                    ElseIf cki.KeyChar = "[" Then
                        strCommand &= cki.KeyChar
                        Console.Write(cki.KeyChar)
                    ElseIf cki.KeyChar = "]" Then
                        strCommand &= cki.KeyChar
                        Console.Write(cki.KeyChar)
                    ElseIf cki.Key = ConsoleKey.Enter Then
                        If strCommand.Trim.Length > 0 Then
                            'strCommand = strCommand.ToUpper
                            If strCommand.ToUpper.StartsWith("SHOW") = True Then
                                Select Case strCommand.ToUpper
                                    Case "SHOW OUS"
                                        List(strDomainFQDN & strPath, "organizationalUnit", "ou", SearchScope.OneLevel)
                                    Case "SHOW ALL OUS"
                                        List(strDomainFQDN & strPath, "organizationalUnit", "ou", SearchScope.Subtree)
                                    Case "SHOW COMPUTERS"
                                        List(strDomainFQDN & strPath, "computer", "cn", SearchScope.OneLevel)
                                    Case "SHOW ALL COMPUTERS"
                                        List(strDomainFQDN & strPath, "computer", "cn", SearchScope.Subtree)
                                    Case Else
                                        Console.WriteLine(vbCrLf & "Not complete command")
                                End Select
                            ElseIf strCommand.ToUpper.StartsWith("EXPORT") = True Then
                                If strCommand.ToUpper.StartsWith("EXPORT ALL OUS TO ") = True Then
                                    strTemp = strCommand.Substring(18)
                                    List(strDomainFQDN & strPath, "organizationalUnit", "ou", SearchScope.Subtree, strTemp)
                                ElseIf strCommand.ToUpper.StartsWith("EXPORT OUS TO ") = True Then
                                    strTemp = strCommand.Substring(14)
                                    List(strDomainFQDN & strPath, "organizationalUnit", "ou", SearchScope.OneLevel, strTemp)
                                ElseIf strCommand.ToUpper.StartsWith("EXPORT ALL COMPUTERS TO ") = True Then
                                    strTemp = strCommand.Substring(24)
                                    List(strDomainFQDN & strPath, "computer", "cn,operatingSystem,operatingSystemServicePack", SearchScope.Subtree, strTemp)
                                ElseIf strCommand.ToUpper.StartsWith("EXPORT ALL USERS TO ") = True Then
                                    strTemp = strCommand.Substring(20)
                                    List(strDomainFQDN & strPath, "user", "cn,distinguishedName,homeMDB", SearchScope.Subtree, strTemp)
                                Else
                                    Console.WriteLine(vbCrLf & "Not complete command")
                                End If
                            ElseIf strCommand.ToUpper.StartsWith("GOTO") = True Then
                                If strCommand.ToUpper.StartsWith("GOTO OU ") = True Then
                                    If strPath <> "" Then
                                        strPath = "/OU=" & strCommand.Substring(8) & "," & strPath.Substring(1)
                                    Else
                                        strPath = "/OU=" & strCommand.Substring(8) & "," & strdefaultNamingContext
                                    End If
                                    Console.WriteLine(vbCrLf & strPath)
                                ElseIf strCommand.ToUpper.StartsWith("GOTO CN ") = True Then
                                    If strPath <> "" Then
                                        strPath = "/CN=" & strCommand.Substring(8) & "," & strPath.Substring(1)
                                    Else
                                        strPath = "/CN=" & strCommand.Substring(8) & "," & strdefaultNamingContext
                                    End If
                                    Console.WriteLine(vbCrLf & strPath)
                                End If
                            Else
                                Console.WriteLine(vbCrLf & "Invalid command")
                            End If
                        End If
                        strCommand = ""
                    Else

                    End If
                Loop Until cki.Key = ConsoleKey.Escape
            Else
                ShowUsage()
            End If
        Catch ex As Exception
            ShowUsage()
        End Try





    End Sub
    Private Sub ShowUsage()
        Console.WriteLine("Usage: ListADObjects.exe SERVER[\PATH] OBJECTCLASS ATTRIBUTE1,ATTRIBUTE2,ATTRIBUTE3")
        Console.WriteLine("Example: ListADObjects.exe dc1.domain.local\cn=Computers,dc=domain,dc=local computer cn,operatingSystem,operatingSystemServicePack")
    End Sub

    Private Sub List(ByVal strADsPath As String, Optional ByVal strobjectClass As String = "organizationalUnit", Optional ByVal strAttributes As String = "cn", Optional ByVal objScope As SearchScope = SearchScope.OneLevel, Optional ByVal strOutputFile As String = "")
        'Additional Defaults
        Dim blnExport As Boolean = False
        Dim objOutputFile As StreamWriter

        If strOutputFile <> "" Then
            Try
                objOutputFile = New StreamWriter(strOutputFile)

                blnExport = True
            Catch ex As Exception

            End Try
        End If

        If strADsPath.StartsWith("LDAP://") = False Then
            strADsPath = "LDAP://" & strADsPath
        End If

        If strobjectClass = "organizationalUnit" Then
            strAttributes = "ou"
        End If

        Try
            Dim arrAttributes As String()
            Dim strAttribute As String

            arrAttributes = strAttributes.Split(",")

            Try
                Dim sr As SearchResult
                Dim src As SearchResultCollection

                Dim de As DirectoryEntry = New DirectoryEntry(strADsPath)
                Dim ds As DirectorySearcher = New DirectorySearcher(de)
                Dim rpvc As System.DirectoryServices.ResultPropertyValueCollection
                Dim item As String
                Dim strAttributeValue As String = ""
                Dim strHeader As String = ""
                Dim strLine As String = ""

                ds.PageSize = 1000
                ds.SearchScope = objScope

                For Each strAttribute In arrAttributes
                    ds.PropertiesToLoad.Add(strAttribute)
                    If strHeader <> "" Then
                        strHeader = strHeader & vbTab & strAttribute
                    Else
                        strHeader = strAttribute
                    End If
                Next

                If blnExport = False Then
                    Console.WriteLine(vbCrLf & strHeader)
                Else
                    Try
                        Console.WriteLine(vbCrLf & "Exporting to file")
                        objOutputFile.WriteLine(strHeader)
                    Catch ex As Exception

                    End Try
                End If

                ds.Filter = "(objectClass=" & strobjectClass & ")"

                src = ds.FindAll

                For Each sr In src
                    strLine = ""
                    For Each strAttribute In arrAttributes
                        strAttributeValue = ""
                        Try
                            rpvc = sr.Properties.Item(strAttribute)
                            If Not rpvc Is Nothing Then
                                For Each item In rpvc
                                    If strAttributeValue <> "" Then
                                        strAttributeValue = strAttributeValue & "|" & item
                                    Else
                                        strAttributeValue = item
                                    End If
                                Next
                            End If
                        Catch ex As Exception
                            strAttributeValue = ""
                        End Try
                        If strLine <> "" Then
                            strLine = strLine & vbTab & strAttributeValue
                        Else
                            strLine = strAttributeValue
                        End If
                    Next
                    If strLine <> "" Then
                        If blnExport = False Then
                            Console.WriteLine(strLine)
                        Else
                            Try
                                objOutputFile.WriteLine(strLine)
                            Catch ex As Exception

                            End Try
                        End If
                    End If
                Next
                If blnExport = True Then
                    Try
                        objOutputFile.Close()
                    Catch ex As Exception

                    End Try
                End If
                Console.WriteLine("Done")
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
            'Else
            'ShowUsage()
            'End If
        Catch ex As Exception
            ShowUsage()
        End Try
    End Sub
End Module
