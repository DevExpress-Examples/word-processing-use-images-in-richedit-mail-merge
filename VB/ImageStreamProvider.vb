Imports System.IO
Imports System.Data
Imports DevExpress.Office.Services

Namespace RichEditImageMailMerge

'#Region "#iuristreamprovider"
    Public Class ImageStreamProvider
        Implements IUriStreamProvider

        Private Shared ReadOnly prefix As String = "dbimg://"

        Private table As DataTable

        Private columnName As String

        Public Sub New(ByVal sourceTable As DataTable, ByVal imageColumn As String)
            table = sourceTable
            columnName = imageColumn
        End Sub

        Public Function GetStream(ByVal uri As String) As Stream Implements IUriStreamProvider.GetStream
            uri = uri.Trim()
            If Not uri.StartsWith(prefix) Then Return Nothing
            Dim strId As String = uri.Substring(prefix.Length).Trim()
            Dim id As Integer
            If Not Integer.TryParse(strId, id) Then Return Nothing
            Dim row As DataRow = table.Rows.Find(id)
            If row Is Nothing Then Return Nothing
            Dim bytes As Byte() = TryCast(row(columnName), Byte())
            If bytes Is Nothing Then Return Nothing

            Dim memoryStream As MemoryStream = New MemoryStream(bytes)
            Return memoryStream
        End Function
'#End Region  ' #iuristreamprovider
    End Class
End Namespace
