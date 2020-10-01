using System;
using System.Drawing.Printing;
using System.Windows.Forms;
using Aspose.Words.Rendering;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing.Printing
{
    class PrintCachePrinterSettings : TestDataHelper
    {
        [Test]
        public static void CachePrinterSettings()
        {
            //ExStart:CachePrinterSettings
            Document doc = new Document(MyDir + "Rendering.docx");

            // Build layout
            doc.UpdatePageLayout();

            // Create settings, setup printing
            PrinterSettings settings = new PrinterSettings();
            settings.PrinterName = "Microsoft XPS Document Writer";

            // Create AsposeWordsPrintDocument and cache settings
            AsposeWordsPrintDocument printDocument = new AsposeWordsPrintDocument(doc);
            printDocument.PrinterSettings = settings;
            printDocument.CachePrinterSettings();

            printDocument.Print();
            //ExEnd:CachePrinterSettings
        }

        [Test]
        public static void PrintProgressDialog()
        {
            // ExStart:PrintProgressDialog
            // Load the documents which store the shapes we want to render
            Document doc = new Document(MyDir + "Rendering.docx");
            // Obtain the settings of the default printer
            PrinterSettings settings = new PrinterSettings();

            // The standard print controller comes with no UI
            PrintController standardPrintController =
                new StandardPrintController();

            // Print the document using the custom print controller
            AsposeWordsPrintDocument prntDoc = new AsposeWordsPrintDocument(doc);
            prntDoc.PrinterSettings = settings;
            prntDoc.PrintController = standardPrintController;
            prntDoc.Print();
            //ExEnd:PrintProgressDialog
        }

        public static void Run()
        {
            // ExStart:PrintPreviewSettingsDialog
            Document doc = new Document(MyDir + "Rendering.docx");

            PrintDialog printDlg = new PrintDialog();

            // Initialize the print dialog with the number of pages in the document.
            printDlg.AllowSomePages = true;
            printDlg.PrinterSettings.MinimumPage = 1;
            printDlg.PrinterSettings.MaximumPage = doc.PageCount;
            printDlg.PrinterSettings.FromPage = 1;
            printDlg.PrinterSettings.ToPage = doc.PageCount;

            // Сheck if the user accepted the print settings and whether to proceed to document preview.
            if (printDlg.ShowDialog() != DialogResult.OK)
                return;

            // Create a special Aspose.Words implementation of the .NET PrintDocument class.
            // Pass the printer settings from the print dialog to the print document.
            AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
            awPrintDoc.PrinterSettings = printDlg.PrinterSettings;

            // Initialize the print preview dialog.
            PrintPreviewDialog previewDlg = new PrintPreviewDialog();

            // Pass the Aspose.Words print document to the print preview dialog.
            previewDlg.Document = awPrintDoc;

            // Specify additional parameters of the print preview dialog.
            previewDlg.ShowInTaskbar = true;
            previewDlg.MinimizeBox = true;
            previewDlg.PrintPreviewControl.Zoom = 1;
            previewDlg.Document.DocumentName = doc.OriginalFileName;
            previewDlg.WindowState = FormWindowState.Maximized;

            // Occur whenever the print preview dialog is first displayed.
            previewDlg.Shown += PreviewDlg_Shown;

            // Show the appropriately configured print preview dialog.
            previewDlg.ShowDialog();
            // ExEnd:PrintPreviewSettingsDialog
        }

        // ExStart:PrintPreviewSettingsDialogEvent
        private static void PreviewDlg_Shown(object sender, EventArgs e)
        {
            // Bring the print preview dialog on top when it is initially displayed.
            ((PrintPreviewDialog)sender).Activate();
        }
        // ExEnd:PrintPreviewSettingsDialogEvent
    }
}