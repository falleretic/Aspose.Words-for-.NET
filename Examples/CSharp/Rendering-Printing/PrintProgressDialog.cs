using Aspose.Words.Rendering;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class PrintProgressDialog : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            // ExStart:PrintProgressDialog
            // Load the documents which store the shapes we want to render
            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");
            // Obtain the settings of the default printer
            System.Drawing.Printing.PrinterSettings settings = new System.Drawing.Printing.PrinterSettings();

            // The standard print controller comes with no UI
            System.Drawing.Printing.PrintController standardPrintController =
                new System.Drawing.Printing.StandardPrintController();

            // Print the document using the custom print controller
            AsposeWordsPrintDocument prntDoc = new AsposeWordsPrintDocument(doc);
            prntDoc.PrinterSettings = settings;
            prntDoc.PrintController = standardPrintController;
            prntDoc.Print();
            //ExEnd:PrintProgressDialog
        }
    }
}