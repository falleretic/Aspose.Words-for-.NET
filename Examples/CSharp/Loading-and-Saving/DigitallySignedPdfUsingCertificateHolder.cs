using System;
using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class DigitallySignedPdfUsingCertificateHolder : TestDataHelper
    {
        public static void Run()
        {
            //ExStart:DigitallySignedPdfUsingCertificateHolder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Writeln("Test Signed PDF.");

            PdfSaveOptions options = new PdfSaveOptions();
            options.DigitalSignatureDetails = new PdfDigitalSignatureDetails(
                CertificateHolder.Create(LoadingSavingDir + "CioSrv1.pfx", "cinD96..arellA"), "reason", "location",
                DateTime.Now);

            doc.Save(ArtifactsDir + "DigitallySignedPdfUsingCertificateHolder.pdf", options);
            //ExEnd:DigitallySignedPdfUsingCertificateHolder
        }
    }
}