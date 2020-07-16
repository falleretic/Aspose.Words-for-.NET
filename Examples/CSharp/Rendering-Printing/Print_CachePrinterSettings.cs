using System.Drawing.Printing;
using Aspose.Words.Rendering;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class PrintCachePrinterSettings : TestDataHelper
    {
        [Test]
        public static void CachePrinterSettings()
        {
            //ExStart:CachePrinterSettings
            Document doc = new Document(MailMergeDir + "TestFile.doc");

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
    }
}