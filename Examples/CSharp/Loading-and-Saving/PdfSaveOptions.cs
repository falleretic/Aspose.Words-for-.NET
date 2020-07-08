using System;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class Doc2Pdf : TestDataHelper
    {
        [Test]
        public static void DisplayDocTitleInWindowTitlebar()
        {
            //ExStart:DisplayDocTitleInWindowTitlebar
            Document doc = new Document(LoadingSavingDir + "Rendering.doc");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.DisplayDocTitle = true;

            doc.Save(ArtifactsDir + "DisplayDocTitleInWindowTitlebar.pdf", saveOptions);
            //ExEnd:DisplayDocTitleInWindowTitlebar
        }

        [Test]
        //ExStart:PdfRenderWarnings
        public static void PdfRenderWarnings()
        {
            Document doc = new Document(LoadingSavingDir + "PdfRenderWarnings.doc");

            // Set a SaveOptions object to not emulate raster operations
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.MetafileRenderingOptions = new MetafileRenderingOptions
            {
                EmulateRasterOperations = false,
                RenderingMode = MetafileRenderingMode.VectorWithFallback
            };

            // If Aspose.Words cannot correctly render some of the metafile records
            // to vector graphics then Aspose.Words renders this metafile to a bitmap
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            doc.Save(ArtifactsDir + "PdfRenderWarnings.pdf", saveOptions);

            // While the file saves successfully, rendering warnings that occurred during saving are collected here
            foreach (WarningInfo warningInfo in callback.mWarnings)
            {
                Console.WriteLine(warningInfo.Description);
            }
        }
        // ExStart:RenderMetafileToBitmap
        public class HandleDocumentWarnings : IWarningCallback
        {
            /// <summary>
            /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
            /// potential issue during document processing. The callback can be set to listen for warnings generated during
            /// document load and/or document save.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                // For now type of warnings about unsupported metafile records changed
                // from DataLoss/UnexpectedContent to MinorFormattingLoss
                if (info.WarningType == WarningType.MinorFormattingLoss)
                {
                    Console.WriteLine("Unsupported operation: " + info.Description);
                    mWarnings.Warning(info);
                }
            }

            public WarningInfoCollection mWarnings = new WarningInfoCollection();
        }
        //ExEnd:PdfRenderWarnings

        [Test]
        public static void DigitallySignedPdfUsingCertificateHolder()
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