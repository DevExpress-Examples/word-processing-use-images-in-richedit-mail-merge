using DevExpress.Office.Services;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace RichEditImageMailMerge {
    static class Program {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main() {

            using (RichEditDocumentServer wordProcessor = new RichEditDocumentServer())
            {
                RegisterUriStreamService(wordProcessor);
                wordProcessor.LoadDocument(Application.StartupPath + @"MailMergeTemplate.rtf");
                wordProcessor.Options.MailMerge.DataSource = NorthwindDataProvider.Categories;
                wordProcessor.Options.MailMerge.ViewMergedData = true;
                MergeToNewDocument(wordProcessor);
            }
        }

        #region #registerprovider
        private static void RegisterUriStreamService(RichEditDocumentServer richEditDocumentServer)
        {
            IUriStreamService uriStreamService = richEditDocumentServer.GetService<IUriStreamService>();
            uriStreamService.RegisterProvider(new ImageStreamProvider(NorthwindDataProvider.Categories, "Picture"));
        }
        #endregion #registerprovider
        #region Mail-merge the document

        private static void MergeToNewDocument(RichEditDocumentServer richEditDocumentServer)
        {
            MailMergeOptions options = richEditDocumentServer.Document.CreateMailMergeOptions();
            options.MergeMode = MergeMode.NewSection;
            string fileName = System.IO.Directory.GetCurrentDirectory() + @"MailMergeResult.rtf";

            richEditDocumentServer.Document.MailMerge(options, fileName, DocumentFormat.Rtf);

            var p = new Process();
            p.StartInfo = new ProcessStartInfo(fileName)
            {
                UseShellExecute = true
            };
            p.Start();
        }
        #endregion Mail-merge the document
    }
}