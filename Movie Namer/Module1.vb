Imports System.Text.RegularExpressions
Imports System.IO

Module Module1

    Sub Main()

#If DEBUG Then
        Dim Path = "E:\Downloads\Videos\Movies To View"
#Else
        Dim Path As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location)
#End If
        Dim oRenamer As New Renamer(Path)
        oRenamer.Process()
        oRenamer.WriteHTML(Path, Path + "\Rotten Tomatoes.html")
        'Console.WriteLine("Press any key to continue...")
        'Console.ReadKey(True)
    End Sub

End Module

Public Class Renamer
    Private sPath As String
    Private sDirs() As String
    Private sFiles() As String
    Private VideoExtentions() As String = {"avi", "mpg", "mkv", "mp4", "m4v"}
    Private Substitues(,) As String = {{"[", "("}, {"]", ")"}, {"{", "("}, {"}", ")"}, {".", " "}, {"_", " "}, {"( www Torrenting com ) - ", ""}}
    Private Terminators() As String = {"DvDrip", "DvdRip", "DVDRip", "DVDRIP", "DVDScr", "DVDSCR", "BluRay", " TS ", "TVRiP", "720p", "x264", "XviD", "AC3", "LIMITED", "LiMiTED", "DiAMOND", "UnKnOwN", "KLAXXON"}
    Private regDone As New Regex("^[\w-!', ]+(\(?\d{4}\)?){1}$")
    Private regYear As New Regex("(\(?\d{4}\)?)")

    Public Sub New(ByVal Path As String)
        Dim Files() As String = Directory.GetFiles(Path)
        Dim DirName As String = ""
        For b As Integer = 0 To Files.Count - 1
            For Each ext As String In Me.VideoExtentions
                If Files(b).EndsWith(ext) Then
                    DirName = Files(b).Substring(Files(b).LastIndexOf("\"c) + 1)
                    DirName = DirName.Substring(0, DirName.Length - 4)
                    Try
                        If Not Directory.Exists(Path + "\" + DirName) Then
                            Directory.CreateDirectory(Path + "\" + DirName)
                        End If
                        File.Move(Files(b), Path + "\" + DirName + "\" + Files(b).Substring(Files(b).LastIndexOf("\"c) + 1))
                    Catch ex As Exception
                        Console.WriteLine("Filed to move " + Files(b) + " to new directory " + DirName)
                    End Try

                End If
            Next
        Next
        Dim Directories() As String = Directory.GetDirectories(Path)
        For a As Integer = 0 To Directories.Count - 1
            Directories(a) = Directories(a).Substring(Directories(a).LastIndexOf("\"c) + 1)
        Next a
        Me.sPath = Path
        Me.sDirs = Directories
    End Sub

    Public Sub Process()

        For Each DirName As String In Me.sDirs
            'If DirName.StartsWith("Harry Potter") Then Debugger.Break()
            If regDone.Match(DirName).Success Then
#If DEBUG Then
                Debug.WriteLine("Not Changing:" + DirName)
#End If
            Else
                Rename(DirName)
            End If
        Next

    End Sub

    Public Sub Rename(ByVal Directory As String)
        Dim NewName As String = Directory
        For a As Integer = 0 To Me.Substitues.GetUpperBound(0)
            NewName = NewName.Replace(Me.Substitues(a, 0), Me.Substitues(a, 1))
        Next

        'If NewName.StartsWith("Harry Potter") Then Debugger.Break()

        'Find Date
        Dim regMatch = Me.regYear.Match(NewName)
        If regMatch.Success Then
            NewName = NewName.Substring(0, regMatch.Index)
            NewName = NewName.Trim()
            Select Case regMatch.Value.Substring(0, 1)
                Case "("
                    NewName = NewName + " " + regMatch.Value
                Case Else
                    NewName = NewName + " (" + regMatch.Value + ")"
            End Select
        Else
            'Find Terminator
            For Each term In Me.Terminators
                If NewName.Contains(term) Then
                    NewName = NewName.Substring(0, NewName.IndexOf(term))
                    NewName = NewName.Trim()
                End If
            Next
        End If

        'Make proper
        NewName = StrConv(NewName, VbStrConv.ProperCase)

        While NewName.Contains("  ")
            NewName = NewName.Replace("  ", " ")
        End While

        If Directory <> NewName Then
#If DEBUG Then
            Debug.WriteLine("Renaming: {0} To: {1}", Directory, NewName)
#Else
            Console.WriteLine("Renaming: {0} To: {1}", Directory, NewName)
            Try
                IO.Directory.Move(Me.sPath + "\" + Directory, Me.sPath + "\" + NewName)
            Catch ex As Exception
                Console.WriteLine("Failed to Rename")
                Try
                    IO.Directory.Move(Me.sPath + "\" + Directory, Me.sPath + "\" + NewName + " (DUP)")
                Catch ex2 As Exception
                End Try
            End Try
#End If
        End If

    End Sub

    Private Encodes(,) As String = {{" ", "+"}, {"(", "%28"}, {")", "%29"}, {"'", "%27"}}
    Private Function Encode(ByVal text As String) As String
        For a As Integer = 0 To Me.Encodes.GetUpperBound(0)
            text = text.Replace(Me.Encodes(a, 0), Me.Encodes(a, 1))
        Next
        Return text
    End Function

    Private Function RemoveEndDate(Text As String) As String
        Dim Texta() As Char = Text.ToCharArray
        If Texta.Last = ")"c AndAlso Texta(Texta.Length - 6) = "("c AndAlso IsNumeric(Text.Substring(Text.Length - 5, 4)) Then
            Return Text.Substring(0, Text.Length - 7)
        End If
        Return Text
    End Function

    Public Sub WriteHTML(ByVal Path As String, ByVal Filename As String)

        Dim Template As String

        Dim oAssembly As System.Reflection.Assembly
        Dim oStream As System.IO.Stream
        Dim oStreamRead As StreamReader
        oAssembly = System.Reflection.Assembly.LoadFrom(System.Reflection.Assembly.GetExecutingAssembly.Location)
        oStream = oAssembly.GetManifestResourceStream("MovieNamer.html.htm")
        oStreamRead = New StreamReader(oStream)
        Template = oStreamRead.ReadToEnd()
        oStreamRead.Dispose()
        oStream.Dispose()

        Dim Directories() As String = Directory.GetDirectories(Path)

        Dim di As DirectoryInfo = New DirectoryInfo(Path)
        Dim dis() As DirectoryInfo = di.GetDirectories
        Dim ordered = dis.OrderByDescending(Function(f) f.CreationTime)

        Dim sb As New System.Text.StringBuilder()

        Dim current_letter As Char = ""
        Dim NameNoYear As String = ""
        Dim NameNoYearOrder As String = ""
        For a As Integer = 0 To Directories.Count - 1
            Directories(a) = Directories(a).Substring(Directories(a).LastIndexOf("\"c) + 1)
            sb.AppendLine("<tr>")
            'Do the sorted list
            'including year
            'sb.AppendLine(vbTab + vbTab + "<td><a href='http://www.rottentomatoes.com/search/?search=" + Encode(Directories(a)) + "&sitesearch=rt'>" + Directories(a) + "</a></td>")
            'withoutyear & and direct movie link
            NameNoYear = RemoveEndDate(Directories(a))
            NameNoYearOrder = RemoveEndDate(ordered(a).Name)
            sb.AppendLine(vbTab + vbTab + "<td><a href=""http://www.rottentomatoes.com/m/" + Encode(NameNoYear.ToLower).Replace("+"c, "_"c) + """>" + Directories(a) + "</a> ")
            sb.AppendLine(vbTab + vbTab + "<a href=""http://www.rottentomatoes.com/search/?search=" + Encode(NameNoYear) + "&sitesearch=rt"">Search</a></td>")

            'Do Unsorted
            'sb.AppendLine(vbTab + vbTab + "<td><a href='http://www.rottentomatoes.com/search/?search=" + Encode(ordered(a).Name) + "&sitesearch=rt'>" + ordered(a).Name + "</a></td>")
            sb.AppendLine(vbTab + vbTab + "<td><a href=""http://www.rottentomatoes.com/m/" + Encode(NameNoYearOrder.ToLower).Replace("+"c, "_"c) + """>" + ordered(a).Name + "</a> ")
            sb.AppendLine(vbTab + vbTab + "<a href=""http://www.rottentomatoes.com/search/?search=" + Encode(NameNoYearOrder) + "&sitesearch=rt"">Search</a> ")
            sb.AppendLine("</tr>")

        Next a

        Dim Output As String
        Output = Template.Replace("{{TABLE}}", sb.ToString)
        Dim out As New StreamWriter(Filename)
        out.Write(Output)
        out.Close()
        out.Dispose()

    End Sub

End Class