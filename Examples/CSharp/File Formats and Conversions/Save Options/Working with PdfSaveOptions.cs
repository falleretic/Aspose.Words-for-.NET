using System;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.File_Formats_and_Conversions.Save_Options
{
    class Doc2Pdf : TestDataHelper
    {
        [Test]
        public static void DisplayDocTitleInWindowTitlebar()
        {
            //ExStart:DisplayDocTitleInWindowTitlebar
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.DisplayDocTitle = true;

            doc.Save(ArtifactsDir + "PdfSaveOptions.DisplayDocTitleInWindowTitlebar.pdf", saveOptions);
            //ExEnd:DisplayDocTitleInWindowTitlebar
        }

        [Test]
        //ExStart:PdfRenderWarnings
        public static void PdfRenderWarnings()
        {
            Document doc = new Document(MyDir + "WMF with image.docx");

            MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions();
            metafileRenderingOptions.EmulateRasterOperations = false;
            metafileRenderingOptions.RenderingMode = MetafileRenderingMode.VectorWithFallback;

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.MetafileRenderingOptions = metafileRenderingOptions;
            
            // If Aspose.Words cannot correctly render some of the metafile records
            // to vector graphics then Aspose.Words renders this metafile to a bitmap.
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            doc.Save(ArtifactsDir + "PdfSaveOptions.PdfRenderWarnings.pdf", saveOptions);

            // While the file saves successfully, rendering warnings that occurred during saving are collected here.
            foreach (WarningInfo warningInfo in callback.mWarnings)
            {
                Console.WriteLine(warningInfo.Description);
            }
        }

        //ExStart:RenderMetafileToBitmap
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
                // from DataLoss/UnexpectedContent to MinorFormattingLoss.
                if (info.WarningType == WarningType.MinorFormattingLoss)
                {
                    Console.WriteLine("Unsupported operation: " + info.Description);
                    mWarnings.Warning(info);
                }
            }

            public WarningInfoCollection mWarnings = new WarningInfoCollection();
        }
        //ExEnd:RenderMetafileToBitmap
        //ExEnd:PdfRenderWarnings

        [Test]
        public static void DigitallySignedPdfUsingCertificateHolder()
        {
            //ExStart:DigitallySignedPdfUsingCertificateHolder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Writeln("Test Signed PDF.");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.DigitalSignatureDetails = new PdfDigitalSignatureDetails(
                CertificateHolder.Create(MyDir + "CioSrv1.pfx", "cinD96..arellA"), "reason", "location",
                DateTime.Now);

            doc.Save(ArtifactsDir + "PdfSaveOptions.DigitallySignedPdfUsingCertificateHolder.pdf", saveOptions);
            //ExEnd:DigitallySignedPdfUsingCertificateHolder
        }

        [Test]
        public static void EmbeddedAllFonts()
        {
            //ExStart:EmbeddAllFonts
            Document doc = new Document(MyDir + "Rendering.docx");

            // Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true.
            // The property below can be changed each time a document is rendered.
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.EmbedFullFonts = true;

            // The output PDF will be embedded with all fonts found in the document.
            doc.Save(ArtifactsDir + "PdfSaveOptions.EmbeddedFontsInPdf.pdf", saveOptions);
            //ExEnd:EmbeddAllFonts
        }

        [Test]
        public static void EmbeddedSubsetFonts()
        {
            //ExStart:EmbeddSubsetFonts
            Document doc = new Document(MyDir + "Rendering.docx");
            
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.EmbedFullFonts = false;
            
            // The output PDF will contain subsets of the fonts in the document.
            // Only the glyphs used in the document are included in the PDF fonts.
            doc.Save(ArtifactsDir + "PdfSaveOptions.EmbeddSubsetFonts.pdf", saveOptions);
            //ExEnd:EmbeddSubsetFonts
        }

        [Test]
        public static void DisableEmbedWindowsFonts()
        {
            // ExStart:DisableEmbedWindowsFonts
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedNone;
            
            // The output PDF will be saved without embedding standard windows fonts.
            doc.Save(ArtifactsDir + "PdfSaveOptions.DisableEmbedWindowsFonts.pdf", saveOptions);
            // ExEnd:DisableEmbedWindowsFonts
        }

        [Test]
        public static void AvoidEmbeddingCoreFonts()
        {
            //ExStart:AvoidEmbeddingCoreFonts
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.UseCoreFonts = true;

            // The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
            doc.Save(ArtifactsDir + "PdfSaveOptions.AvoidEmbeddingCoreFonts.pdf", saveOptions);
            //ExEnd:AvoidEmbeddingCoreFonts
        }

        [Test]
        public static void SkipEmbeddedArialAndTimesRomanFonts()
        {
            //ExStart:SkipEmbeddedArialAndTimesRomanFonts
            Document doc = new Document(MyDir + "Rendering.docx");
            
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll;

            doc.Save(ArtifactsDir + "PdfSaveOptions.SkipEmbeddedArialAndTimesRomanFonts.pdf", saveOptions);
            //ExEnd:SkipEmbeddedArialAndTimesRomanFonts
        }

        [Test]
        public static void EscapeUri()
        {
            //ExStart:EscapeUri
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertHyperlink("Testlink", "https://www.google.com/search?q=%2Fthe%20test", false);
            builder.Writeln();
            builder.InsertHyperlink("https://www.google.com/search?q=%2Fthe%20test", "https://www.google.com/search?q=%2Fthe%20test", false);

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.EscapeUri = false;

            doc.Save(ArtifactsDir + "PdfSaveOptions.EscapeUri.pdf", saveOptions);
            //ExEnd:EscapeUri
        }

        [Test]
        public static void ExportHeaderFooterBookmarks()
        {
            //ExStart:ExportHeaderFooterBookmarks
            Document doc = new Document(MyDir + "Bookmarks in headers and footers.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.OutlineOptions.DefaultBookmarksOutlineLevel = 1;
            saveOptions.HeaderFooterBookmarksExportMode = HeaderFooterBookmarksExportMode.First;

            doc.Save(ArtifactsDir + "PdfSaveOptions.ExportHeaderFooterBookmarks.pdf", saveOptions);
            //ExEnd:ExportHeaderFooterBookmarks
        }

        [Test]
        public static void ScaleWmfFontsToMetafileSize()
        {
            //ExStart:ScaleWmfFontsToMetafileSize
            Document doc = new Document(MyDir + "WMF with text.docx");

            MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions();
            metafileRenderingOptions.ScaleWmfFontsToMetafileSize = false;
        
            // If Aspose.Words cannot correctly render some of the metafile records to vector graphics
            // then Aspose.Words renders this metafile to a bitmap.
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.MetafileRenderingOptions = metafileRenderingOptions;

            doc.Save(ArtifactsDir + "PdfSaveOptions.ScaleWmfFontsToMetafileSize.pdf", saveOptions);
            //ExEnd:ScaleWmfFontsToMetafileSize
        }

        [Test]
        public static void AdditionalTextPositioning()
        {
            //ExStart:AdditionalTextPositioning
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.AdditionalTextPositioning = true;

            doc.Save(ArtifactsDir + "PdfSaveOptions.AdditionalTextPositioning.pdf", saveOptions);
            //ExEnd:AdditionalTextPositioning
        }

        [Test]
        public static void ConversionToPdf17()
        {
            //ExStart:ConversionToPDF17
            Document originalDoc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.Compliance = PdfCompliance.Pdf17;

            originalDoc.Save(ArtifactsDir + "PdfSaveOptions.ConversionToPdf17.pdf", saveOptions);
            //ExEnd:ConversionToPDF17
        }

        [Test]
        public static void DownsamplingImages()
        {
            //ExStart:DownsamplingImages
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            // The first two images in the input document will be affected by this.
            saveOptions.DownsampleOptions.Resolution = 36;
            // We can set a minimum threshold for downsampling.
            // This value will prevent the second image in the input document from being downsampled.
            saveOptions.DownsampleOptions.ResolutionThreshold = 128;

            doc.Save(ArtifactsDir + "PdfSaveOptions.DownsamplingImages.pdf", saveOptions);
            //ExEnd:DownsamplingImages
        }

        [Test]
        public static void SaveToPdfWithOutline()
        {
            //ExStart:SaveToPdfWithOutline
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.OutlineOptions.HeadingsOutlineLevels = 3;
            saveOptions.OutlineOptions.ExpandedOutlineLevels = 1;

            doc.Save(ArtifactsDir + "PdfSaveOptions.SaveToPdfWithOutline.pdf", saveOptions);
            // ExEnd:SaveToPdfWithOutline
        }

        [Test]
        public static void CustomPropertiesExport()
        {
            //ExStart:CustomPropertiesExport
            Document doc = new Document();
            doc.CustomDocumentProperties.Add("Company", "My value");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

            doc.Save(ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf", saveOptions);
            //ExEnd:CustomPropertiesExport
        }

        [Test]
        public static void ExportDocumentStructure()
        {
            //ExStart:ExportDocumentStructure
            Document doc = new Document(MyDir + "Paragraphs.docx");

            // The file size will be increased and the structure will be visible in the "Content" navigation pane
            // of Adobe Acrobat Pro, while editing the .pdf.
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.ExportDocumentStructure = true;

            doc.Save(ArtifactsDir + "PdfSaveOptions.ExportDocumentStructure.pdf", saveOptions);
            //ExEnd:ExportDocumentStructure
        }

        [Test]
        public static void PdfImageComppression()
        {
            //ExStart:PdfImageComppression
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.ImageCompression = PdfImageCompression.Jpeg;
            saveOptions.PreserveFormFields = true;
        
            doc.Save(ArtifactsDir + "PdfSaveOptions.PdfImageCompression.pdf", saveOptions);

            PdfSaveOptions saveOptionsA1B = new PdfSaveOptions();
            saveOptionsA1B.Compliance = PdfCompliance.PdfA1b;
            saveOptionsA1B.ImageCompression = PdfImageCompression.Jpeg;
            // Use JPEG compression at 50% quality to reduce file size.
            saveOptionsA1B.JpegQuality = 100;
            saveOptionsA1B.ImageColorSpaceExportMode = PdfImageColorSpaceExportMode.SimpleCmyk;
        
            doc.Save(ArtifactsDir + "PdfSaveOptions.PdfImageComppression PDF_A_1_B.pdf", saveOptionsA1B);
            //ExEnd:PdfImageComppression
        }

        [Test]
        public static void UpdateLastPrintedProperty()
        {
            //ExStart:UpdateIfLastPrinted
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.UpdateLastPrintedProperty = false;

            doc.Save(ArtifactsDir + "PdfSaveOptions.UpdateIfLastPrinted.pdf", saveOptions);
            //ExEnd:UpdateIfLastPrinted
        }

        [Test]
        public static void Dml3DEffectsRendering()
        {
            //ExStart:Dml3DEffectsRendering
            Document doc = new Document(MyDir + "Rendering.docx");

            SaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.Dml3DEffectsRenderingMode = Dml3DEffectsRenderingMode.Advanced;
            
            doc.Save(ArtifactsDir + "PdfSaveOptions.Dml3DEffectsRendering.pdf", saveOptions);
            //ExEnd:Dml3DEffectsRendering
        }

        [Test]
        public static void InterpolateImages()
        {
            //ExStart:InterpolateImages
            Document doc = new Document();

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.InterpolateImages = true;
            
            doc.Save(ArtifactsDir + "PdfSaveOptions.InterpolateImages.pdf", saveOptions);
            //ExEnd:InterpolateImages
        }
    }
}