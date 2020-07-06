using Aspose.Words.Saving;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class WorkingWithPdfSaveOptions : TestDataHelper
    {
        [Test]
        public static void EscapeUriInPdf()
        {
            //ExStart:EscapeUriInPdf
            Document doc = new Document(RenderingPrintingDir + "EscapeUri.docx");

            PdfSaveOptions options = new PdfSaveOptions();
            options.EscapeUri = false;

            doc.Save(ArtifactsDir + "loadOptions.pdf", options);
            //ExEnd:EscapeUriInPdf
        }

        [Test]
        public static void ExportHeaderFooterBookmarks()
        {
            //ExStart:ExportHeaderFooterBookmarks
            Document doc = new Document(RenderingPrintingDir + "TestFile.docx");

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
            Document doc = new Document(RenderingPrintingDir + "MetafileRendering.docx");

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
            Document doc = new Document(RenderingPrintingDir + "TestFile.docx");

            PdfSaveOptions options = new PdfSaveOptions();
            options.AdditionalTextPositioning = true;

            doc.Save(ArtifactsDir + "AdditionalTextPositioning.pdf", options);
            //ExEnd:AdditionalTextPositioning
        }

        [Test]
        public static void ConversionToPdf17()
        {
            //ExStart:ConversionToPDF17
            Document originalDoc = new Document(ChartsDir + "Document.docx");

            // Provide PDFSaveOption compliance to PDF17
            // or just convert without SaveOptions
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.Compliance = PdfCompliance.Pdf17;

            originalDoc.Save(ArtifactsDir + "ConversionToPdf17.pdf", pdfSaveOptions);
            //ExEnd:ConversionToPDF17
        }

        public static void DownsamplingImages(string dataDir)
        {
            // ExStart:DownsamplingImages
            // Open a document that contains images 
            Document doc = new Document(dataDir + "Rendering.doc");

            // If we want to convert the document to .pdf, we can use a SaveOptions implementation to customize the saving process
            PdfSaveOptions options = new PdfSaveOptions();

            // We can set the output resolution to a different value
            // The first two images in the input document will be affected by this
            options.DownsampleOptions.Resolution = 36;

            // We can set a minimum threshold for downsampling 
            // This value will prevent the second image in the input document from being downsampled
            options.DownsampleOptions.ResolutionThreshold = 128;

            doc.Save(dataDir + "PdfSaveOptions.DownsampleOptions.pdf", options);
            // ExEnd:DownsamplingImages
        }

        public static void SaveToPdfWithOutline(string dataDir)
        {
            // ExStart:SaveToPdfWithOutline
            // Open a document
            Document doc = new Document(dataDir + "Rendering.doc");

            PdfSaveOptions options = new PdfSaveOptions();
            options.OutlineOptions.HeadingsOutlineLevels = 3;
            options.OutlineOptions.ExpandedOutlineLevels = 1;

            doc.Save(dataDir + "Rendering.SaveToPdfWithOutline.pdf", options);
            // ExEnd:SaveToPdfWithOutline
        }

        public static void CustomPropertiesExport(string dataDir)
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

            doc.Save(dataDir + "PdfSaveOptions.CustomPropertiesExport.pdf", options);
            // ExEnd:CustomPropertiesExport
        }

        public static void ExportDocumentStructure(string dataDir)
        {
            // ExStart:ExportDocumentStructure
            // Open a document
            Document doc = new Document(dataDir + "Paragraphs.docx");

            // Create a PdfSaveOptions object and configure it to preserve the logical structure that's in the input document
            // The file size will be increased and the structure will be visible in the "Content" navigation pane
            // of Adobe Acrobat Pro, while editing the .pdf
            PdfSaveOptions options = new PdfSaveOptions();
            options.ExportDocumentStructure = true;

            doc.Save(dataDir + "PdfSaveOptions.ExportDocumentStructure.pdf", options);
            // ExEnd:ExportDocumentStructure
        }

        public static void PdfImageComppression(string dataDir)
        {
            // ExStart:PdfImageComppression
            // Open a document
            Document doc = new Document(dataDir + "SaveOptions.PdfImageCompression.rtf");

            PdfSaveOptions options = new PdfSaveOptions
            {
                ImageCompression = PdfImageCompression.Jpeg,
                PreserveFormFields = true
            };
            
            doc.Save(dataDir + "SaveOptions.PdfImageCompression.pdf", options);

            PdfSaveOptions optionsA1B = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1b,
                ImageCompression = PdfImageCompression.Jpeg,

                // Use JPEG compression at 50% quality to reduce file size
                JpegQuality = 100, 
                ImageColorSpaceExportMode = PdfImageColorSpaceExportMode.SimpleCmyk
            };
            
            doc.Save(dataDir + "SaveOptions.PdfImageComppression PDF_A_1_B.pdf", optionsA1B);
            // ExEnd:PdfImageComppression
            Console.WriteLine("\nFile saved at " + dataDir);
        }

        public static void UpdateIfLastPrinted(string dataDir)
        {
            // ExStart:UpdateIfLastPrinted
            // Open a document
            Document doc = new Document(dataDir + "Rendering.doc");

            SaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.UpdateLastPrintedProperty = false;

            doc.Save(dataDir + "PdfSaveOptions.UpdateIfLastPrinted.pdf", saveOptions);
            // ExEnd:UpdateIfLastPrinted
        }

        public static void EffectsRendering(string dataDir)
        {
            // ExStart:EffectsRendering
            // Open a document
            Document doc = new Document(dataDir + "Rendering.doc");

            SaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.Dml3DEffectsRenderingMode = Dml3DEffectsRenderingMode.Advanced;
            
            doc.Save(dataDir, saveOptions);
            // ExEnd:EffectsRendering
        }

        public static void SetImageInterpolation(string dataDir)
        {
            // ExStart:SetImageInterpolation
            Document doc = new Document(dataDir);

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.InterpolateImages = true;
            
            doc.Save(dataDir, saveOptions);
            // ExEnd:SetImageInterpolation
        }
    }
}