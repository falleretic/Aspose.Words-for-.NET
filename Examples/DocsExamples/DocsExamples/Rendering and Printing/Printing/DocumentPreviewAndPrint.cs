#if NET462
using System;
using System.Windows.Forms;
using Aspose.Words.Rendering;
using Aspose.Words;

namespace DocsExamples.Rendering_and_Printing.Printing
{
    //ExStart:ActivePrintPreviewDialogClass 
    internal class ActivePrintPreviewDialog : PrintPreviewDialog
    {
        /// <summary>
        /// Brings the Print Preview dialog on top when it is initially displayed.
        /// </summary>
        protected override void OnShown(EventArgs e)
        {
            Activate();
            base.OnShown(e);
        }
    }
    //ExEnd:ActivePrintPreviewDialogClass

    /// <summary>
    /// This project is set to target the x86 platform because the .NET print dialog does not 
    /// Seem to show when calling from a 64-bit application.
    /// </summary>
    internal class DocumentPreviewAndPrint : DocsExamplesBase
    {
        public static void Run()
        {
            Document doc = new Document(MyDir + "TestFile.doc");

            //ExStart:PrintDialog
            PrintDialog printDlg = new PrintDialog();
            // Initialize the print dialog with the number of pages in the document
            printDlg.AllowSomePages = true;
            printDlg.PrinterSettings.MinimumPage = 1;
            printDlg.PrinterSettings.MaximumPage = doc.PageCount;
            printDlg.PrinterSettings.FromPage = 1;
            printDlg.PrinterSettings.ToPage = doc.PageCount;
            // ExEnd:PrintDialog

            //ExStart:ShowDialog
            if (!printDlg.ShowDialog().Equals(DialogResult.OK))
                return;
            //ExEnd:ShowDialog

            //ExStart:AsposeWordsPrintDocument
            // Pass the printer settings from the dialog to the print document
            AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
            awPrintDoc.PrinterSettings = printDlg.PrinterSettings;
            //ExEnd:AsposeWordsPrintDocument

            //ExStart:ActivePrintPreviewDialog
            ActivePrintPreviewDialog previewDlg = new ActivePrintPreviewDialog();
            // Pass the Aspose.Words print document to the Print Preview dialog
            previewDlg.Document = awPrintDoc;
            // Specify additional parameters of the Print Preview dialog
            previewDlg.ShowInTaskbar = true;
            previewDlg.MinimizeBox = true;
            previewDlg.PrintPreviewControl.Zoom = 1;
            previewDlg.Document.DocumentName = "TestName.doc";
            previewDlg.WindowState = FormWindowState.Maximized;
            // Show the appropriately configured Print Preview dialog
            previewDlg.ShowDialog();
            //ExEnd:ActivePrintPreviewDialog
        }
    }
}
#endif