using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Load_Options
{
    internal class WorkingWithPdfLoadOptions : DocsExamplesBase
    {
        [Test]
        public static void LoadEncryptedPdf()
        {
            //ExStart:LoadEncryptedPdf  
            Document doc = new Document(MyDir + "Pdf Document.pdf");

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                EncryptionDetails = new PdfEncryptionDetails("Aspose", null, PdfEncryptionAlgorithm.RC4_40)
            };

            doc.Save(ArtifactsDir + "WorkingWithPdfLoadOptions.LoadEncryptedPdf.pdf", pdfSaveOptions);

            PdfLoadOptions pdfLoadOptions = new PdfLoadOptions { Password = "Aspose", LoadFormat = LoadFormat.Pdf };

            doc = new Document(ArtifactsDir + "WorkingWithPdfLoadOptions.LoadEncryptedPdf.pdf", pdfLoadOptions);
            //ExEnd:LoadEncryptedPdf
        }

        [Test]
        public static void LoadPageRangeOfPdf()
        {
            //ExStart:LoadPageRangeOfPdf  
            PdfLoadOptions pdfLoadOptions = new PdfLoadOptions { PageIndex = 0, PageCount = 1 };

            //ExStart:LoadPDF
            Document doc = new Document(MyDir + "Pdf Document.pdf", pdfLoadOptions);
            doc.Save(ArtifactsDir + "WorkingWithPdfLoadOptions.LoadPageRangeOfPdf.pdf");
            //ExEnd:LoadPDF
            //ExEnd:LoadPageRangeOfPdf
        }
    }
}
