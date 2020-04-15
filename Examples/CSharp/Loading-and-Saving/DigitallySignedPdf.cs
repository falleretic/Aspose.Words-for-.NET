//ExStart:X509Certificates
using System.Security.Cryptography.X509Certificates;
//ExEnd:X509Certificates
using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class DigitallySignedPdf : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:DigitallySignedPdf
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Writeln("Test Signed PDF.");
            
            // Load the certificate from disk
            // The other constructor overloads can be used to load certificates from different locations
            X509Certificate2 cert = new X509Certificate2(LoadingSavingDir + "signature.pfx", "signature");

            // Pass the certificate and details to the save options class to sign with
            PdfSaveOptions options = new PdfSaveOptions();
            options.DigitalSignatureDetails = new PdfDigitalSignatureDetails();

            doc.Save(ArtifactsDir + "DigitallySignedPdf.pdf", options);
            //ExEnd:DigitallySignedPdf
        }
    }
}