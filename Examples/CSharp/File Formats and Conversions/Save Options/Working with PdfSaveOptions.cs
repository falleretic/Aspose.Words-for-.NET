using System;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class Doc2Pdf : TestDataHelper
    {
        [Test]
        public static void DisplayDocTitleInWindowTitlebar()
        {
            //ExStart:DisplayDocTitleInWindowTitlebar
            Document doc = new Document(LoadingSavingDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.DisplayDocTitle = true;

            doc.Save(ArtifactsDir + "DisplayDocTitleInWindowTitlebar.pdf", saveOptions);
            //ExEnd:DisplayDocTitleInWindowTitlebar
        }

        [Test]
        //ExStart:PdfRenderWarnings
        public static void PdfRenderWarnings()
        {
            Document doc = new Document(LoadingSavingDir + "WMF with image.docx");

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

        [Test]
        public static void EmbeddedAllFonts()
        {
            //ExStart:EmbeddAllFonts
            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");

            // Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true. The property below can be changed
            // Each time a document is rendered
            PdfSaveOptions options = new PdfSaveOptions();
            options.EmbedFullFonts = true;

            // The output PDF will be embedded with all fonts found in the document
            doc.Save(ArtifactsDir + "EmbeddedFontsInPdf.pdf", options);
            //ExEnd:EmbeddAllFonts
        }

        [Test]
        public static void EmbeddedSubsetFonts()
        {
            //ExStart:EmbeddSubsetFonts
            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");
            
            // To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false
            PdfSaveOptions options = new PdfSaveOptions();
            options.EmbedFullFonts = false;
            
            // The output PDF will contain subsets of the fonts in the document. Only the glyphs used
            // in the document are included in the PDF fonts
            doc.Save(ArtifactsDir + "EmbeddSubsetFonts.pdf", options);
            //ExEnd:EmbeddSubsetFonts
        }

        [Test]
        public static void SetFontEmbeddingMode()
        {
            // ExStart:SetFontEmbeddingMode
            // Load the document to render.
            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");

            // To disable embedding standard windows font use the PdfSaveOptions and set the EmbedStandardWindowsFonts property to false.
            PdfSaveOptions options = new PdfSaveOptions();
            options.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedNone;

            // The output PDF will be saved without embedding standard windows fonts.
            doc.Save(ArtifactsDir + "Rendering.DisableEmbedWindowsFonts.pdf");
            // ExEnd:SetFontEmbeddingMode
        }

        [Test]
        public static void AvoidEmbeddingCoreFonts()
        {
            //ExStart:AvoidEmbeddingCoreFonts
            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");

            // To disable embedding of core fonts and substitute PDF type 1 fonts set UseCoreFonts to true
            PdfSaveOptions options = new PdfSaveOptions();
            options.UseCoreFonts = true;

            // The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc
            doc.Save(ArtifactsDir + "AvoidEmbeddingCoreFonts.pdf", options);
            //ExEnd:AvoidEmbeddingCoreFonts
        }

        [Test]
        public static void SkipEmbeddedArialAndTimesRomanFonts()
        {
            //ExStart:SkipEmbeddedArialAndTimesRomanFonts
            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");
            
            // To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false
            // To disable embedding standard windows font use the PdfSaveOptions and set the EmbedStandardWindowsFonts property to false
            PdfSaveOptions options = new PdfSaveOptions();
            options.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll;

            // The output PDF will be saved without embedding standard windows fonts
            doc.Save(ArtifactsDir + "SkipEmbeddedArialAndTimesRomanFonts.pdf");
            //ExEnd:SkipEmbeddedArialAndTimesRomanFonts
        }

        [Test]
        public static void EscapeUriInPdf()
        {
            //ExStart:EscapeUriInPdf
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertHyperlink("Testlink", "https://www.google.com/search?q=%2Fthe%20test", false);
            builder.Writeln();
            builder.InsertHyperlink("https://www.google.com/search?q=%2Fthe%20test", "https://www.google.com/search?q=%2Fthe%20test", false);

            PdfSaveOptions options = new PdfSaveOptions();
            options.EscapeUri = false;

            doc.Save(ArtifactsDir + "loadOptions.pdf", options);
            //ExEnd:EscapeUriInPdf
        }

        [Test]
        public static void ExportHeaderFooterBookmarks()
        {
            //ExStart:ExportHeaderFooterBookmarks
            Document doc = new Document(RenderingPrintingDir + "Bookmarks in headers and footers.docx");

            PdfSaveOptions options = new PdfSaveOptions();
            options.OutlineOptions.DefaultBookmarksOutlineLevel = 1;
            options.HeaderFooterBookmarksExportMode = HeaderFooterBookmarksExportMode.First;

            doc.Save(ArtifactsDir + "ExportHeaderFooterBookmarks.pdf", options);
            //ExEnd:ExportHeaderFooterBookmarks
        }

        [Test]
        public static void ScaleWmfFontsToMetafileSize()
        {
            //ExStart:ScaleWmfFontsToMetafileSize
            Document doc = new Document(RenderingPrintingDir + "WMF with text.docx");

            MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions
            {
                ScaleWmfFontsToMetafileSize = false
            };

            // If Aspose.Words cannot correctly render some of the metafile records to vector graphics
            // then Aspose.Words renders this metafile to a bitmap
            PdfSaveOptions options = new PdfSaveOptions { MetafileRenderingOptions = metafileRenderingOptions };

            doc.Save(ArtifactsDir + "ScaleWmfFontsToMetafileSize.pdf", options);
            //ExEnd:ScaleWmfFontsToMetafileSize
        }

        [Test]
        public static void AdditionalTextPositioning()
        {
            //ExStart:AdditionalTextPositioning
            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");

            PdfSaveOptions options = new PdfSaveOptions();
            options.AdditionalTextPositioning = true;

            doc.Save(ArtifactsDir + "AdditionalTextPositioning.pdf", options);
            //ExEnd:AdditionalTextPositioning
        }

        [Test]
        public static void ConversionToPdf17()
        {
            //ExStart:ConversionToPDF17
            Document originalDoc = new Document(RenderingPrintingDir + "Rendering.docx");

            // Provide PDFSaveOption compliance to PDF17
            // or just convert without SaveOptions
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.Compliance = PdfCompliance.Pdf17;

            originalDoc.Save(ArtifactsDir + "ConversionToPdf17.pdf", pdfSaveOptions);
            //ExEnd:ConversionToPDF17
        }

        [Test]
        public static void DownsamplingImages()
        {
            // ExStart:DownsamplingImages
            // Open a document that contains images 
            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");

            // If we want to convert the document to .pdf, we can use a SaveOptions implementation to customize the saving process
            PdfSaveOptions options = new PdfSaveOptions();

            // We can set the output resolution to a different value
            // The first two images in the input document will be affected by this
            options.DownsampleOptions.Resolution = 36;

            // We can set a minimum threshold for downsampling 
            // This value will prevent the second image in the input document from being downsampled
            options.DownsampleOptions.ResolutionThreshold = 128;

            doc.Save(ArtifactsDir + "PdfSaveOptions.DownsampleOptions.pdf", options);
            // ExEnd:DownsamplingImages
        }

        [Test]
        public static void SaveToPdfWithOutline()
        {
            // ExStart:SaveToPdfWithOutline
            // Open a document
            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");

            PdfSaveOptions options = new PdfSaveOptions();
            options.OutlineOptions.HeadingsOutlineLevels = 3;
            options.OutlineOptions.ExpandedOutlineLevels = 1;

            doc.Save(ArtifactsDir + "Rendering.SaveToPdfWithOutline.pdf", options);
            // ExEnd:SaveToPdfWithOutline
        }

        [Test]
        public static void CustomPropertiesExport()
        {
            // ExStart:CustomPropertiesExport
            // Open a document
            Document doc = new Document();

            // Add a custom document property that doesn't use the name of some built in properties
            doc.CustomDocumentProperties.Add("Company", "My value");

            // Configure the PdfSaveOptions like this will display the properties
            // in the "Document Properties" menu of Adobe Acrobat Pro
            PdfSaveOptions options = new PdfSaveOptions();
            options.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

            doc.Save(ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf", options);
            // ExEnd:CustomPropertiesExport
        }

        [Test]
        public static void ExportDocumentStructure()
        {
            // ExStart:ExportDocumentStructure
            // Open a document
            Document doc = new Document(RenderingPrintingDir + "Paragraphs.docx");

            // Create a PdfSaveOptions object and configure it to preserve the logical structure that's in the input document
            // The file size will be increased and the structure will be visible in the "Content" navigation pane
            // of Adobe Acrobat Pro, while editing the .pdf
            PdfSaveOptions options = new PdfSaveOptions();
            options.ExportDocumentStructure = true;

            doc.Save(ArtifactsDir + "PdfSaveOptions.ExportDocumentStructure.pdf", options);
            // ExEnd:ExportDocumentStructure
        }

        [Test]
        public static void PdfImageComppression()
        {
            // ExStart:PdfImageComppression
            // Open a document
            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");

            PdfSaveOptions options = new PdfSaveOptions
            {
                ImageCompression = PdfImageCompression.Jpeg,
                PreserveFormFields = true
            };
            
            doc.Save(ArtifactsDir + "SaveOptions.PdfImageCompression.pdf", options);

            PdfSaveOptions optionsA1B = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1b,
                ImageCompression = PdfImageCompression.Jpeg,

                // Use JPEG compression at 50% quality to reduce file size
                JpegQuality = 100, 
                ImageColorSpaceExportMode = PdfImageColorSpaceExportMode.SimpleCmyk
            };
            
            doc.Save(ArtifactsDir + "SaveOptions.PdfImageComppression PDF_A_1_B.pdf", optionsA1B);
            // ExEnd:PdfImageComppression
        }

        [Test]
        public static void UpdateIfLastPrinted()
        {
            // ExStart:UpdateIfLastPrinted
            // Open a document
            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");

            SaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.UpdateLastPrintedProperty = false;

            doc.Save(ArtifactsDir + "PdfSaveOptions.UpdateIfLastPrinted.pdf", saveOptions);
            // ExEnd:UpdateIfLastPrinted
        }

        [Test]
        public static void EffectsRendering()
        {
            // ExStart:EffectsRendering
            // Open a document
            Document doc = new Document(RenderingPrintingDir + "Rendering.docx");

            SaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.Dml3DEffectsRenderingMode = Dml3DEffectsRenderingMode.Advanced;
            
            doc.Save(ArtifactsDir + "EffectsRendering.pdf", saveOptions);
            // ExEnd:EffectsRendering
        }

        [Test]
        public static void SetImageInterpolation()
        {
            // ExStart:SetImageInterpolation
            Document doc = new Document();

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.InterpolateImages = true;
            
            doc.Save(ArtifactsDir + "SetImageInterpolation.pdf", saveOptions);
            // ExEnd:SetImageInterpolation
        }
    }
}