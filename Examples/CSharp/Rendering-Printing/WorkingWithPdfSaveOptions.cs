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
            Document doc = new Document(MailMergeDir + "EscapeUri.docx");

            PdfSaveOptions options = new PdfSaveOptions();
            options.EscapeUri = false;

            doc.Save(ArtifactsDir + "loadOptions.pdf", options);
            //ExEnd:EscapeUriInPdf
        }

        [Test]
        public static void ExportHeaderFooterBookmarks()
        {
            //ExStart:ExportHeaderFooterBookmarks
            Document doc = new Document(MailMergeDir + "TestFile.docx");

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
            Document doc = new Document(MailMergeDir + "MetafileRendering.docx");

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
            Document doc = new Document(MailMergeDir + "TestFile.docx");

            PdfSaveOptions options = new PdfSaveOptions();
            options.AdditionalTextPositioning = true;

            doc.Save(ArtifactsDir + "AdditionalTextPositioning.pdf", options);
            //ExEnd:AdditionalTextPositioning
        }

        [Test]
        public static void ConversionToPdf17()
        {
            //ExStart:ConversionToPDF17
            Document originalDoc = new Document(MailMergeDir + "Document.docx");

            // Provide PDFSaveOption compliance to PDF17
            // or just convert without SaveOptions
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.Compliance = PdfCompliance.Pdf17;

            originalDoc.Save(ArtifactsDir + "ConversionToPdf17.pdf", pdfSaveOptions);
            //ExEnd:ConversionToPDF17
        }
    }
}