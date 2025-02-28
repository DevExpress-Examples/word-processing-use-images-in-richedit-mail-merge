Imports System.Data
Imports System.Data.OleDb

Namespace RichEditImageMailMerge

    Public Module NorthwindDataProvider

        Private connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\nwind.mdb;Persist Security Info=False;"

        Sub New()
            Using connection As OleDbConnection = New OleDbConnection(connectionString)
                Dim selectCommand As OleDbCommand = New OleDbCommand("SELECT * FROM Categories", connection)
                Dim da As OleDbDataAdapter = New OleDbDataAdapter(selectCommand)
                categoriesField = New DataTable("Categories")
                da.Fill(categoriesField)
                categoriesField.Constraints.Add("IDPK", categoriesField.Columns("CategoryID"), True)
                selectCommand.Dispose()
            End Using
        End Sub

        Private categoriesField As DataTable

        Public ReadOnly Property Categories As DataTable
            Get
                Return categoriesField
            End Get
        End Property
    End Module
End Namespace
