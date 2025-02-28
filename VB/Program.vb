Imports DevExpress.Office.Services
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System.Windows.Forms

Namespace RichEditImageMailMerge
    Friend Module Program
        ''' <summary>
        ''' The main entry point for the application.
        ''' </summary>
        <STAThread>
        Sub Main()
            Using wordProcessor As New RichEditDocumentServer()
                RegisterUriStreamService(wordProcessor)
                wordProcessor.LoadDocument(Application.StartupPath & "\MailMergeTemplate.rtf")
                wordProcessor.Options.MailMerge.DataSource = NorthwindDataProvider.Categories
                wordProcessor.Options.MailMerge.ViewMergedData = True
                MergeToNewDocument(wordProcessor)
            End Using
        End Sub

#Region "registerprovider"
        Private Sub RegisterUriStreamService(richEditDocumentServer As RichEditDocumentServer)
            Dim uriStreamService As IUriStreamService = richEditDocumentServer.GetService(Of IUriStreamService)()
            uriStreamService.RegisterProvider(New ImageStreamProvider(NorthwindDataProvider.Categories, "Picture"))
        End Sub
#End Region

#Region "Mail-merge the document"
        Private Sub MergeToNewDocument(richEditDocumentServer As RichEditDocumentServer)
            Dim options As MailMergeOptions = richEditDocumentServer.Document.CreateMailMergeOptions()
            options.MergeMode = MergeMode.NewSection
            Dim fileName As String = System.IO.Directory.GetCurrentDirectory() & "\MailMergeResult.rtf"

            richEditDocumentServer.Document.MailMerge(options, fileName, DocumentFormat.Rtf)

            Dim p As New Process()
            p.StartInfo = New ProcessStartInfo(fileName) With {
                .UseShellExecute = True
            }
            p.Start()
        End Sub
#End Region
    End Module
End Namespace
