Imports System.Text
Imports System.IO

Public Class PageWriter
    Const AZ As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    Private Template As String
    Private sb As New StringBuilder
    Private ListAtoZ() As DirectoryInfo
    Private ListByDate() As DirectoryInfo

    Public Sub New(ByVal SearchPath As String, ByVal Filename As String)

        Dim oAssembly As System.Reflection.Assembly
        Dim oStream As System.IO.Stream
        Dim oStreamRead As IO.StreamReader

        oAssembly = System.Reflection.Assembly.LoadFrom(System.Reflection.Assembly.GetExecutingAssembly.Location)
        oStream = oAssembly.GetManifestResourceStream("MovieNamer.html.htm")
        oStreamRead = New IO.StreamReader(oStream)
        Me.Template = oStreamRead.ReadToEnd()

        oStreamRead.Dispose()
        oStream.Dispose()

        Dim di As DirectoryInfo = New DirectoryInfo(SearchPath)
        Me.ListAtoZ = di.GetDirectories
        ' Me.ListByDate = Me.ListAtoZ.OrderByDescending(Function(f) f.CreationTime)


    End Sub

    Private Sub AddHeaderAtoZ()
        sb.Append("<td>")
        For a = 0 To 26
            sb.Append("<a href='#" + AZ(a) + "'>" + AZ(a) + "</a> ")
        Next
        sb.Append("</td>")
    End Sub

    Private Sub AddHeaderTime()

    End Sub

    Public Sub Write(ByVal Path As String)



    End Sub
End Class
