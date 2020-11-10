using System;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Save_Options
{
    internal class WorkingWithPdfSaveOptions : DocsExamplesBase
    {
        [Test]
        public static void DisplayDocTitleInWindowTitlebar()
        {
            //ExStart:DisplayDocTitleInWindowTitlebar
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { DisplayDocTitle = true };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.DisplayDocTitleInWindowTitlebar.pdf", pdfSaveOptions);
            //ExEnd:DisplayDocTitleInWindowTitlebar
        }

        [Test]
        //ExStart:PdfRenderWarnings
        public static void PdfRenderWarnings()
        {
            Document doc = new Document(MyDir + "WMF with image.docx");

            MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions
            {
                EmulateRasterOperations = false, RenderingMode = MetafileRenderingMode.VectorWithFallback
            };

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { MetafileRenderingOptions = metafileRenderingOptions };

            // If Aspose.Words cannot correctly render some of the metafile records
            // to vector graphics then Aspose.Words renders this metafile to a bitmap.
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.PdfRenderWarnings.pdf", pdfSaveOptions);

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

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                DigitalSignatureDetails = new PdfDigitalSignatureDetails(
                    CertificateHolder.Create(MyDir + "morzal.pfx", "aw"), "reason", "location",
                    DateTime.Now)
            };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.DigitallySignedPdfUsingCertificateHolder.pdf", pdfSaveOptions);
            //ExEnd:DigitallySignedPdfUsingCertificateHolder
        }

        [Test]
        public static void EmbeddedAllFonts()
        {
            //ExStart:EmbeddAllFonts
            Document doc = new Document(MyDir + "Rendering.docx");

            // The output PDF will be embedded with all fonts found in the document.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { EmbedFullFonts = true };
            
            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.EmbeddedFontsInPdf.pdf", pdfSaveOptions);
            //ExEnd:EmbeddAllFonts
        }

        [Test]
        public static void EmbeddedSubsetFonts()
        {
            //ExStart:EmbeddSubsetFonts
            Document doc = new Document(MyDir + "Rendering.docx");

            // The output PDF will contain subsets of the fonts in the document.
            // Only the glyphs used in the document are included in the PDF fonts.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { EmbedFullFonts = false };
            
            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.EmbeddSubsetFonts.pdf", pdfSaveOptions);
            //ExEnd:EmbeddSubsetFonts
        }

        [Test]
        public static void DisableEmbedWindowsFonts()
        {
            // ExStart:DisableEmbedWindowsFonts
            Document doc = new Document(MyDir + "Rendering.docx");

            // The output PDF will be saved without embedding standard windows fonts.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { FontEmbeddingMode = PdfFontEmbeddingMode.EmbedNone };
            
            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.DisableEmbedWindowsFonts.pdf", pdfSaveOptions);
            // ExEnd:DisableEmbedWindowsFonts
        }

        [Test]
        public static void SkipEmbeddedArialAndTimesRomanFonts()
        {
            //ExStart:SkipEmbeddedArialAndTimesRomanFonts
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.SkipEmbeddedArialAndTimesRomanFonts.pdf", pdfSaveOptions);
            //ExEnd:SkipEmbeddedArialAndTimesRomanFonts
        }

        [Test]
        public static void AvoidEmbeddingCoreFonts()
        {
            //ExStart:AvoidEmbeddingCoreFonts
            Document doc = new Document(MyDir + "Rendering.docx");

            // The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { UseCoreFonts = true };
            
            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.AvoidEmbeddingCoreFonts.pdf", pdfSaveOptions);
            //ExEnd:AvoidEmbeddingCoreFonts
        }
        
        [Test]
        public static void EscapeUri()
        {
            //ExStart:EscapeUri
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertHyperlink("Testlink", 
                "https://www.google.com/search?q=%2Fthe%20test", false);
            builder.Writeln();
            builder.InsertHyperlink("https://www.google.com/search?q=%2Fthe%20test", 
                "https://www.google.com/search?q=%2Fthe%20test", false);

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { EscapeUri = false };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.EscapeUri.pdf", pdfSaveOptions);
            //ExEnd:EscapeUri
        }

        [Test]
        public static void ExportHeaderFooterBookmarks()
        {
            //ExStart:ExportHeaderFooterBookmarks
            Document doc = new Document(MyDir + "Bookmarks in headers and footers.docx");

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.OutlineOptions.DefaultBookmarksOutlineLevel = 1;
            pdfSaveOptions.HeaderFooterBookmarksExportMode = HeaderFooterBookmarksExportMode.First;

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.ExportHeaderFooterBookmarks.pdf", pdfSaveOptions);
            //ExEnd:ExportHeaderFooterBookmarks
        }

        [Test]
        public static void ScaleWmfFontsToMetafileSize()
        {
            //ExStart:ScaleWmfFontsToMetafileSize
            Document doc = new Document(MyDir + "WMF with text.docx");

            MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions
            {
                ScaleWmfFontsToMetafileSize = false
            };

            // If Aspose.Words cannot correctly render some of the metafile records to vector graphics
            // then Aspose.Words renders this metafile to a bitmap.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { MetafileRenderingOptions = metafileRenderingOptions };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.ScaleWmfFontsToMetafileSize.pdf", pdfSaveOptions);
            //ExEnd:ScaleWmfFontsToMetafileSize
        }

        [Test]
        public static void AdditionalTextPositioning()
        {
            //ExStart:AdditionalTextPositioning
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { AdditionalTextPositioning = true };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.AdditionalTextPositioning.pdf", pdfSaveOptions);
            //ExEnd:AdditionalTextPositioning
        }

        [Test]
        public static void ConversionToPdf17()
        {
            //ExStart:ConversionToPDF17
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { Compliance = PdfCompliance.Pdf17 };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.ConversionToPdf17.pdf", pdfSaveOptions);
            //ExEnd:ConversionToPDF17
        }

        [Test]
        public static void DownsamplingImages()
        {
            //ExStart:DownsamplingImages
            Document doc = new Document(MyDir + "Rendering.docx");

            // We can set a minimum threshold for downsampling.
            // This value will prevent the second image in the input document from being downsampled.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                DownsampleOptions = { Resolution = 36, ResolutionThreshold = 128 }
            };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.DownsamplingImages.pdf", pdfSaveOptions);
            //ExEnd:DownsamplingImages
        }

        [Test]
        public static void SetOutlineOptions()
        {
            //ExStart:SetOutlineOptions
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.OutlineOptions.HeadingsOutlineLevels = 3;
            pdfSaveOptions.OutlineOptions.ExpandedOutlineLevels = 1;

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.SetOutlineOptions.pdf", pdfSaveOptions);
            // ExEnd:SetOutlineOptions
        }

        [Test]
        public static void CustomPropertiesExport()
        {
            //ExStart:CustomPropertiesExport
            Document doc = new Document();
            doc.CustomDocumentProperties.Add("Company", "Aspose");

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { CustomPropertiesExport = PdfCustomPropertiesExport.Standard };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.CustomPropertiesExport.pdf", pdfSaveOptions);
            //ExEnd:CustomPropertiesExport
        }

        [Test]
        public static void ExportDocumentStructure()
        {
            //ExStart:ExportDocumentStructure
            Document doc = new Document(MyDir + "Paragraphs.docx");

            // The file size will be increased and the structure will be visible in the "Content" navigation pane
            // of Adobe Acrobat Pro, while editing the .pdf.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { ExportDocumentStructure = true };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.ExportDocumentStructure.pdf", pdfSaveOptions);
            //ExEnd:ExportDocumentStructure
        }

        [Test]
        public static void PdfImageComppression()
        {
            //ExStart:PdfImageComppression
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                ImageCompression = PdfImageCompression.Jpeg, PreserveFormFields = true
            };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.PdfImageCompression.pdf", pdfSaveOptions);

            PdfSaveOptions pdfSaveOptionsA1B = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1b,
                ImageCompression = PdfImageCompression.Jpeg,
                JpegQuality = 100, // Use JPEG compression at 50% quality to reduce file size.
                ImageColorSpaceExportMode = PdfImageColorSpaceExportMode.SimpleCmyk
            };
            

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.PdfImageCompression.Pdf_A1b.pdf", pdfSaveOptionsA1B);
            //ExEnd:PdfImageComppression
        }

        [Test]
        public static void UpdateLastPrintedProperty()
        {
            //ExStart:UpdateIfLastPrinted
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { UpdateLastPrintedProperty = false };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.UpdateIfLastPrinted.pdf", pdfSaveOptions);
            //ExEnd:UpdateIfLastPrinted
        }

        [Test]
        public static void Dml3DEffectsRendering()
        {
            //ExStart:Dml3DEffectsRendering
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { Dml3DEffectsRenderingMode = Dml3DEffectsRenderingMode.Advanced };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.Dml3DEffectsRendering.pdf", pdfSaveOptions);
            //ExEnd:Dml3DEffectsRendering
        }

        [Test]
        public static void InterpolateImages()
        {
            //ExStart:InterpolateImages
            Document doc = new Document(MyDir + "Rendering.docx");

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { InterpolateImages = true };

            doc.Save(ArtifactsDir + "WorkingWithPdfSaveOptions.InterpolateImages.pdf", pdfSaveOptions);
            //ExEnd:InterpolateImages
        }
    }
}