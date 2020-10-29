using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace SiteExamples.File_Formats_and_Conversions.Load_Options
{
    internal class WorkingWithPdfLoadOptions : SiteExamplesBase
    {
        [Test]
        public static void LoadEncryptedPdf()
        {
            //ExStart:LoadEncryptedPdf  
            Document doc = new Document(MyDir + "Pdf Document.pdf");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.EncryptionDetails = new PdfEncryptionDetails("Aspose", null, PdfEncryptionAlgorithm.RC4_40);

            doc.Save(ArtifactsDir + "WorkingWithPdfLoadOptions.LoadEncryptedPdf.pdf", saveOptions);

            PdfLoadOptions loadOptions = new PdfLoadOptions();
            loadOptions.Password = "Aspose";
            loadOptions.LoadFormat = LoadFormat.Pdf;

            doc = new Document(ArtifactsDir + "WorkingWithPdfLoadOptions.LoadEncryptedPdf.pdf", loadOptions);
            //ExEnd:LoadEncryptedPdf
        }

        [Test]
        public static void LoadPageRangeOfPdf()
        {
            //ExStart:LoadPageRangeOfPdf  
            PdfLoadOptions pdfLoadOptions = new PdfLoadOptions();
            pdfLoadOptions.PageIndex = 0;
            pdfLoadOptions.PageCount = 1;

            //ExStart:LoadPDF
            Document doc = new Document(MyDir + "Pdf Document.pdf", pdfLoadOptions);
            doc.Save(ArtifactsDir + "out.pdf");
            //ExEnd:LoadPDF
            //ExEnd:LoadPageRangeOfPdf
        }
    }
}
